using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace CrazyKTV_SongMgr
{
    /// <summary>What a 錢櫃查詢 looks for: the 錢櫃 tab's query menu, its other-query menu (錢櫃缺歌, 重複歌曲) and its date menu.</summary>
    enum CashboxSearchKind
    {
        SongName,
        SingerName,
        CashboxId,
        AllSongs,
        NewSongs,
        /// <summary>錢櫃缺歌: 錢櫃 songs the 歌庫 does not have. Looks at all of ktv_Cashbox, whatever the 語系 filter.</summary>
        Missing,
        /// <summary>重複歌曲: <see cref="DuplicateSongKey.Rule.IgnoreSongType"/>, as ktv_Cashbox has no 歌曲類別.</summary>
        Duplicate,
        /// <summary>The songs whose 更新日期 (Song_CreatDate) is one day.</summary>
        UpdatedOn,
    }

    /// <summary>
    /// 錢櫃查詢條件: everything one 錢櫃查詢 needs from the screen, captured once. <see cref="Value"/> is the text typed
    /// for SongName / SingerName / CashboxId and the row count for NewSongs; <see cref="Date"/> is the day for UpdatedOn.
    /// </summary>
    sealed class CashboxSearchCriteria
    {
        public CashboxSearchCriteria(CashboxSearchKind kind, string value = "", string lang = null, bool fuzzy = true, bool synonyms = false, DateTime date = default)
        {
            Kind = kind;
            Value = value ?? "";
            Lang = (string.IsNullOrEmpty(lang) || lang == SongSearchCriteria.AllLangs) ? null : lang;
            Fuzzy = fuzzy;
            Synonyms = synonyms;
            Date = date.Date;
        }

        public CashboxSearchKind Kind { get; }
        public string Value { get; }
        /// <summary>The 語系 to keep, or null for every 語系.</summary>
        public string Lang { get; }
        public bool Fuzzy { get; }
        public bool Synonyms { get; }
        public DateTime Date { get; }
    }

    /// <summary>
    /// 錢櫃查詢: finds the ktv_Cashbox songs a <see cref="CashboxSearchCriteria"/> asks for, in the 參考資料庫 of either
    /// kind. Text compares and values bind as in <see cref="SongSearch"/>.
    /// </summary>
    static class CashboxSearch
    {
        /// <summary>The columns of every result, in the order the 錢櫃 list shows them.</summary>
        public const string Columns = "Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate, Song_HaveSong";

        /// <summary>SongMgr's lists for a 錢櫃查詢; 錢櫃缺歌 also reads the 歌庫's songs once.</summary>
        public static SongSearchReference CaptureReference(CashboxSearchKind kind)
        {
            Func<DataRow, bool> hasLibrarySong = null;
            if (kind == CashboxSearchKind.Missing)
            {
                Cashbox.CreateSongDataTable();
                var lowCase = new List<string>(Cashbox.SongDataLowCaseList);
                var fuzzy = new List<string>(Cashbox.SongDataFuzzyList);
                Cashbox.DisposeSongDataTable();
                hasLibrarySong = row => CommonFunc.MatchCashboxSong("CashboxSong", row, "", "", "", lowCase, fuzzy) >= 0;
            }

            return new SongSearchReference(
                Global.SynonymousWordList,
                Global.SingerGroupList,
                (Global.SongAddSpecialStr ?? "").Split('|'),
                hasLibrarySong: hasLibrarySong);
        }

        public static SongSearchResult Run(CashboxSearchCriteria criteria, SongSearchReference reference, string songMgrDatabaseFile)
        {
            var q = new SongSearch.Query("ktv_Cashbox", Columns);
            Func<DataRow, bool> keep = null;
            string v = criteria.Value;

            switch (criteria.Kind)
            {
                case CashboxSearchKind.SongName:
                    if (criteria.Fuzzy)
                    {
                        var names = new List<string> { v };
                        if (criteria.Synonyms) names.AddRange(SongSearch.SynonymousNames(v, reference));
                        q.WhereContainsAny("Song_SongName", SongSearch.WithWidthForms(names));
                    }
                    else
                    {
                        q.WhereEqualsText("Song_SongName", v);
                    }
                    WhereLang(q, criteria).OrderBy("Cashbox_Id");
                    break;

                case CashboxSearchKind.SingerName:
                    if (criteria.Fuzzy)
                    {
                        var singers = new List<string> { v };
                        singers.AddRange(reference.OtherGroupNames(v));
                        q.WhereContainsAny("Song_Singer", SongSearch.WithWidthForms(singers));
                    }
                    else
                    {
                        q.WhereContainsAny("Song_Singer", new[] { v });
                        keep = row => SongSearch.IsSingerOf(row["Song_Singer"].ToString(), v, reference);
                    }
                    WhereLang(q, criteria).OrderBy("Cashbox_Id");
                    break;

                case CashboxSearchKind.CashboxId:
                    if (criteria.Fuzzy) WhereLang(q.WhereContainsAny("Cashbox_Id", new[] { v }), criteria);
                    else q.Where("Cashbox_Id = @CashboxId", ("@CashboxId", v));
                    q.OrderBy("Cashbox_Id");
                    break;

                case CashboxSearchKind.AllSongs:
                    WhereLang(q, criteria).OrderBy("Cashbox_Id");
                    break;

                case CashboxSearchKind.NewSongs:
                    if (!int.TryParse(v, NumberStyles.None, CultureInfo.InvariantCulture, out int top) || top <= 0)
                    {
                        throw new ArgumentException("新進歌曲的筆數必須是正整數: " + v);
                    }
                    WhereLang(q.Top(top), criteria).OrderBy("Song_CreatDate desc, Cashbox_Id desc");
                    break;

                case CashboxSearchKind.Missing:
                    q.OrderBy("Cashbox_Id");
                    keep = row => !reference.HasLibrarySong(row);
                    break;

                case CashboxSearchKind.Duplicate:
                    return new SongSearchResult(SongSearch.Groups(q.Run(songMgrDatabaseFile), row => DuplicateSongKey.Of(DuplicateSongKey.Rule.IgnoreSongType,
                        row["Song_Lang"].ToString(), row["Song_Singer"].ToString(), row["Song_SongName"].ToString(), "", reference), "Cashbox_Id"));

                case CashboxSearchKind.UpdatedOn:
                    q.Where("Song_CreatDate >= @From and Song_CreatDate < @To", ("@From", criteria.Date), ("@To", criteria.Date.AddDays(1)))
                     .OrderBy("Cashbox_Id");
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(criteria), criteria.Kind, "不支援的查詢種類");
            }

            DataTable found = q.Run(songMgrDatabaseFile);
            if (keep != null) SongSearch.RemoveRowsWhere(found, row => !keep(row));
            return new SongSearchResult(found);
        }

        private static SongSearch.Query WhereLang(SongSearch.Query q, CashboxSearchCriteria criteria)
        {
            return (criteria.Lang == null) ? q : q.Where("Song_Lang = @Lang", ("@Lang", criteria.Lang));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CrazyKTV_SongMgr
{
    /// <summary>What a 歌庫查詢 looks for: one kind per entry of the 歌庫查詢 menu, plus 我的最愛.</summary>
    enum SongSearchKind
    {
        SongName,
        SingerName,
        SongId,
        AllSongs,
        NewSongs,
        ChorusSongs,
        SongType,
        SingerType,
        SongTrack,
        CashboxId,
        Favorite,
    }

    /// <summary>
    /// 歌庫查詢條件: everything one 歌庫查詢 needs from the screen, captured once.
    /// <see cref="Value"/> means: the text typed for SongName / SingerName / SongId, the row count for NewSongs,
    /// the 歌曲類別 for SongType ("" for 無類別), the index into the 歌手類別 list for SingerType, the 聲道 number for
    /// SongTrack and the User_Id for Favorite; the other kinds ignore it.
    /// </summary>
    sealed class SongSearchCriteria
    {
        public SongSearchCriteria(SongSearchKind kind, string value = "", string lang = null, bool fuzzy = true, bool synonyms = false)
        {
            Kind = kind;
            Value = value ?? "";
            Lang = (string.IsNullOrEmpty(lang) || lang == AllLangs) ? null : lang;
            Fuzzy = fuzzy;
            Synonyms = synonyms;
        }

        public const string AllLangs = "全部";

        public SongSearchKind Kind { get; }
        public string Value { get; }
        /// <summary>The 語系 to keep, or null for every 語系.</summary>
        public string Lang { get; }
        /// <summary>模糊查詢: SongName / SingerName / SongId match anywhere in the text instead of the whole text.</summary>
        public bool Fuzzy { get; }
        /// <summary>同義字查詢: a fuzzy SongName search also looks for the name with its 同義字 swapped in.</summary>
        public bool Synonyms { get; }
    }

    /// <summary>
    /// The lists a 歌庫查詢 matches against, other than the 歌庫 itself: 同義字, 歌手群組, the special singer names
    /// that are never split as 合唱, and the 錢櫃 numbers. <see cref="Capture"/> reads them from SongMgr's settings
    /// and 參考資料庫; tests build one directly.
    /// </summary>
    sealed class SongSearchReference
    {
        public SongSearchReference(IEnumerable<string> synonymousWords = null, IEnumerable<string> singerGroups = null,
                                   IEnumerable<string> specialSingerNames = null, IEnumerable<string> cashboxIds = null)
        {
            SynonymousWords = (synonymousWords ?? Enumerable.Empty<string>()).ToList();
            SpecialSingerNames = (specialSingerNames ?? Enumerable.Empty<string>()).Where(s => s != "").ToList();
            CashboxIds = new HashSet<string>(cashboxIds ?? Enumerable.Empty<string>(), StringComparer.Ordinal);

            groupOf = new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (string group in singerGroups ?? Enumerable.Empty<string>())
            {
                List<string> names = group.Split(',').Select(n => n.Trim()).Where(n => n != "").ToList();
                foreach (string name in names)
                {
                    if (!groupOf.ContainsKey(name)) groupOf.Add(name, names);
                }
            }
        }

        private readonly Dictionary<string, IReadOnlyList<string>> groupOf;

        /// <summary>Each entry is one set of 同義字 joined by "|".</summary>
        public IReadOnlyList<string> SynonymousWords { get; }
        /// <summary>Singer names holding "&amp;" or "+" that are one singer, not a 合唱 (SongAddSpecialStr).</summary>
        public IReadOnlyList<string> SpecialSingerNames { get; }
        /// <summary>The 錢櫃 numbers, written with as many digits as the 歌庫's 歌曲編號.</summary>
        public ISet<string> CashboxIds { get; }

        /// <summary>The other names of <paramref name="singer"/>'s 歌手群組; empty when it is in none.</summary>
        public IEnumerable<string> OtherGroupNames(string singer)
        {
            if (!groupOf.TryGetValue(singer.Trim(), out IReadOnlyList<string> names)) return Enumerable.Empty<string>();
            return names.Where(n => !n.Equals(singer.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>The singers of a 合唱 written "A&amp;B" or "A+B", keeping the special singer names whole.</summary>
        public List<string> SplitChorus(string singer)
        {
            string marked = singer;
            foreach (string special in SpecialSingerNames)
            {
                marked = Regex.Replace(marked, Regex.Escape(special), "%%" + special + "%%", RegexOptions.IgnoreCase);
            }
            return Regex.Split(marked, "[&+](?=(?:[^%]*%%[^%]*%%)*(?![^%]*%%))")
                        .Select(s => Regex.Replace(s, @"%%|^\s*|\s*$", ""))
                        .ToList();
        }

        /// <summary>SongMgr's current lists; the 錢櫃 numbers are read only when <paramref name="kind"/> needs them.</summary>
        public static SongSearchReference Capture(SongSearchKind kind)
        {
            List<string> cashboxIds = null;
            if (kind == SongSearchKind.CashboxId)
            {
                string digits = (Global.SongMgrMaxDigitCode == "1") ? "D5" : "D6";
                var langs = new HashSet<string>(Global.CashboxSongLangList);
                using (DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, "select Cashbox_Id, Song_Lang from ktv_Cashbox", ""))
                {
                    cashboxIds = dt.AsEnumerable()
                                   .Where(r => langs.Contains(r["Song_Lang"].ToString()))
                                   .Select(r => Convert.ToInt32(r["Cashbox_Id"].ToString()).ToString(digits))
                                   .ToList();
                }
            }

            return new SongSearchReference(
                Global.SynonymousWordList,
                Global.SingerGroupList,
                (Global.SongAddSpecialStr ?? "").Split('|'),
                cashboxIds);
        }
    }

    /// <summary>The songs a 歌庫查詢 found, with the columns the 歌庫查詢 list shows.</summary>
    sealed class SongSearchResult
    {
        public SongSearchResult(DataTable songs) { Songs = songs; }

        public DataTable Songs { get; }
        public int Count => Songs.Rows.Count;
    }

    /// <summary>
    /// 歌庫查詢: finds the songs a <see cref="SongSearchCriteria"/> asks for, in either kind of 歌庫. Text is compared
    /// without regard to case on both Access and SQLite. Values are bound as parameters; a value that cannot mean
    /// anything for its kind (a row count that is not a number, say) throws <see cref="ArgumentException"/>.
    /// </summary>
    static class SongSearch
    {
        /// <summary>The columns of every result, in the order the 歌庫查詢 list shows them.</summary>
        public const string Columns = "Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_SongType, Song_Track, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState, Song_ReplayGain";

        private const int ChorusSingerType = 3;

        // Characters that have a half-width and a full-width form: a fuzzy search tries both.
        private static readonly Regex HasWidthForms = new Regex("[!-~！-～]");

        // The characters of a song name that can have 同義字.
        private static readonly Regex SynonymCandidate = new Regex(@"([⺀-㏿]|[一-鿌㐀-䶵﨎﨏﨑﨓﨔﨟﨡﨣﨤﨧-﨩]|[\ud840-\ud868][\udc00-\udfff]|\ud869[\udc00-\uded6\udf00-\udfff]|[\ud86a-\ud86c][\udc00-\udfff]|\ud86d[\udc00-\udf34\udf40-\udfff]|\ud86e[\udc00-\udc1d]|[가-퟿])");

        public static SongSearchResult Run(SongSearchCriteria criteria, SongSearchReference reference, string databaseFile)
        {
            var q = new Query();
            Func<DataRow, bool> keep = null;
            string v = criteria.Value;

            switch (criteria.Kind)
            {
                case SongSearchKind.SongName:
                    if (criteria.Fuzzy)
                    {
                        var names = new List<string> { v };
                        if (criteria.Synonyms) names.AddRange(SynonymousNames(v, reference));
                        q.WhereContainsAny("Song_SongName", WithWidthForms(names)).WhereLang(criteria).OrderBy("Song_Id");
                    }
                    else
                    {
                        q.WhereEqualsText("Song_SongName", v).WhereLang(criteria).OrderBy("Song_Id");
                    }
                    break;

                case SongSearchKind.SingerName:
                    if (criteria.Fuzzy)
                    {
                        var singers = new List<string> { v };
                        singers.AddRange(reference.OtherGroupNames(v));
                        q.WhereContainsAny("Song_Singer", WithWidthForms(singers)).WhereLang(criteria).OrderBy("Song_Id");
                    }
                    else
                    {
                        // The singer, alone or as one of the singers of a 合唱.
                        q.WhereContainsAny("Song_Singer", new[] { v }).WhereLang(criteria).OrderBy("Song_Id");
                        keep = row => IsSingerOf(row["Song_Singer"].ToString(), v, reference);
                    }
                    break;

                case SongSearchKind.SongId:
                    if (criteria.Fuzzy)
                    {
                        q.WhereContainsAny("Song_Id", new[] { v }).WhereLang(criteria).OrderBy("Song_Id");
                    }
                    else
                    {
                        // A 歌曲編號 is unique, so the 語系 filter does not apply.
                        q.Where("Song_Id = @SongId", ("@SongId", v)).OrderBy("Song_Id");
                    }
                    break;

                case SongSearchKind.AllSongs:
                    q.WhereLang(criteria).OrderBy("Song_Id");
                    break;

                case SongSearchKind.NewSongs:
                    if (!int.TryParse(v, NumberStyles.None, CultureInfo.InvariantCulture, out int top) || top <= 0)
                    {
                        throw new ArgumentException("新進歌曲的筆數必須是正整數: " + v);
                    }
                    q.Top(top).WhereLang(criteria).OrderBy("Song_CreatDate desc, Song_Id desc");
                    break;

                case SongSearchKind.ChorusSongs:
                    q.Where("Song_SingerType = @SingerType", ("@SingerType", ChorusSingerType)).WhereLang(criteria).OrderBy("Song_Id");
                    break;

                case SongSearchKind.SongType:
                    q.WhereEqualsText("Song_SongType", v).WhereLang(criteria).OrderBy("Song_Id");
                    break;

                case SongSearchKind.SingerType:
                    q.Where("Song_SingerType = @SingerType", ("@SingerType", ParseNumber(v, "歌手類別"))).WhereLang(criteria).OrderBy("Song_Id");
                    break;

                case SongSearchKind.SongTrack:
                    q.Where("Song_Track = @SongTrack", ("@SongTrack", ParseNumber(v, "歌曲聲道"))).WhereLang(criteria).OrderBy("Song_Id");
                    break;

                case SongSearchKind.CashboxId:
                    q.WhereLang(criteria).OrderBy("Song_Id");
                    keep = row => reference.CashboxIds.Contains(row["Song_Id"].ToString());
                    break;

                case SongSearchKind.Favorite:
                    q.Where("Song_Id in (select Song_Id from ktv_Favorite where User_Id = @UserId)", ("@UserId", v)).OrderBy("Song_Id");
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(criteria), criteria.Kind, "不支援的查詢種類");
            }

            DataTable songs = q.Run(databaseFile);
            if (keep != null) RemoveRowsWhere(songs, row => !keep(row));
            return new SongSearchResult(songs);
        }

        /// <summary>Whether <paramref name="singer"/> (a song's Song_Singer) is <paramref name="wanted"/>, alone or in a 合唱.</summary>
        private static bool IsSingerOf(string singer, string wanted, SongSearchReference reference)
        {
            if (singer.Equals(wanted, StringComparison.OrdinalIgnoreCase)) return true;
            return reference.SplitChorus(singer).Any(s => s.Equals(wanted.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>The song name with each of its characters that has 同義字 replaced by each of them.</summary>
        private static IEnumerable<string> SynonymousNames(string songName, SongSearchReference reference)
        {
            var words = new HashSet<string>(reference.SynonymousWords.SelectMany(s => s.Split('|')));
            var result = new List<string>();
            foreach (Match match in SynonymCandidate.Matches(songName))
            {
                if (!words.Contains(match.Value)) continue;
                foreach (string set in reference.SynonymousWords.Where(s => s.Contains(match.Value)))
                {
                    foreach (string word in set.Split('|'))
                    {
                        string name = Regex.Replace(songName, Regex.Escape(match.Value), word.Replace("$", "$$"), RegexOptions.IgnoreCase);
                        if (name != songName && !result.Contains(name)) result.Add(name);
                    }
                }
            }
            return result;
        }

        /// <summary>Each text, and its half-width and full-width forms when it has characters with both.</summary>
        private static List<string> WithWidthForms(IEnumerable<string> texts)
        {
            var result = new List<string>();
            void Add(string s) { if (!result.Contains(s)) result.Add(s); }
            foreach (string text in texts)
            {
                Add(text);
                if (HasWidthForms.IsMatch(text))
                {
                    Add(CommonFunc.ConvToNarrow(text));
                    Add(CommonFunc.ConvToWide(text));
                }
            }
            return result;
        }

        private static int ParseNumber(string value, string what)
        {
            if (!int.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int n))
            {
                throw new ArgumentException(what + "必須是數字: " + value);
            }
            return n;
        }

        internal static void RemoveRowsWhere(DataTable table, Func<DataRow, bool> remove)
        {
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                if (remove(table.Rows[i])) table.Rows.RemoveAt(i);
            }
        }

        /// <summary>
        /// One select on ktv_Song, written as Access SQL (rewritten for SQLite by <see cref="Db"/>) with every value
        /// bound as a parameter. Parameters are added in the order they appear, as OleDb binds them by position.
        /// </summary>
        private sealed class Query
        {
            private readonly List<string> conditions = new List<string>();
            private readonly List<(string Name, object Value)> parameters = new List<(string, object)>();
            private string order = "";
            private int top;

            public Query Top(int rows) { top = rows; return this; }

            public Query OrderBy(string columns) { order = columns; return this; }

            public Query Where(string condition, params (string Name, object Value)[] values)
            {
                conditions.Add(condition);
                parameters.AddRange(values);
                return this;
            }

            /// <summary>The column equals the text, whatever the case.</summary>
            public Query WhereEqualsText(string column, string text)
            {
                string p = NextName();
                return Where("LCase(" + column + ") = LCase(" + p + ")", (p, text));
            }

            /// <summary>The column holds any of the texts, whatever the case.</summary>
            public Query WhereContainsAny(string column, IEnumerable<string> texts)
            {
                var ors = new List<string>();
                var values = new List<(string, object)>();
                foreach (string text in texts)
                {
                    string p = NextName();
                    ors.Add("InStr(1,LCase(" + column + "),LCase(" + p + "),0) <>0");
                    values.Add((p, text));
                }
                return Where("(" + string.Join(" or ", ors) + ")", values.ToArray());
            }

            public Query WhereLang(SongSearchCriteria criteria)
            {
                if (criteria.Lang == null) return this;
                return Where("Song_Lang = @Lang", ("@Lang", criteria.Lang));
            }

            private int names;

            private string NextName() => "@p" + (names++).ToString(CultureInfo.InvariantCulture);

            public string Sql
            {
                get
                {
                    var sql = new StringBuilder("select ");
                    if (top > 0) sql.Append("top ").Append(top.ToString(CultureInfo.InvariantCulture)).Append(' ');
                    sql.Append(Columns).Append(" from ktv_Song");
                    if (conditions.Count > 0) sql.Append(" where ").Append(string.Join(" and ", conditions));
                    if (order != "") sql.Append(" order by ").Append(order);
                    return sql.ToString();
                }
            }

            public DataTable Run(string databaseFile)
            {
                using (DbConnection conn = Db.Open(databaseFile))
                using (DbCommand cmd = Db.Command(Sql, conn))
                {
                    foreach (var (name, value) in parameters) cmd.Parameters.AddWithValue(name, value);
                    using (DbDataAdapter da = Db.Adapter(cmd))
                    {
                        var dt = new DataTable();
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
    }
}

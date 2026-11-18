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
    /// <summary>
    /// What a 歌庫查詢 looks for: one kind per entry of the 歌庫查詢 menu, 我的最愛, and one per entry of the 異常查詢
    /// menu. The 異常查詢 kinds look at the whole 歌庫, whatever the 語系 filter.
    /// </summary>
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

        /// <summary>無檔案歌曲: the song's file is not on disk.</summary>
        MissingFile,
        /// <summary>同檔案歌曲: more than one song uses the same file.</summary>
        SameFile,
        /// <summary>重複歌曲 (<see cref="DuplicateSongKey.Rule.Full"/>).</summary>
        Duplicate,
        /// <summary>重複歌曲 (忽略歌曲類別).</summary>
        DuplicateIgnoreSongType,
        /// <summary>重複歌曲 (忽略括號及歌曲類別).</summary>
        DuplicateIgnoreBrackets,
        /// <summary>重複歌曲 (忽略歌手及歌曲類別).</summary>
        DuplicateIgnoreSinger,
        /// <summary>重複歌曲 (合唱歌曲): 合唱 songs with the same 語系 and 歌名, whoever sings them.</summary>
        DuplicateChorus,
        /// <summary>同名歌曲 (忽略括號及歌曲類別).</summary>
        SameNameIgnoreBrackets,
        /// <summary>檔名不符 (歌曲聲道): the file name does not carry the song's 聲道.</summary>
        TrackNotInFileName,
        /// <summary>歌手未在資料庫: the singer, or one singer of a 合唱, is not in ktv_Singer.</summary>
        SingerNotInLibrary,
        /// <summary>歌手類別不符: the song's 歌手類別 is not the one ktv_Singer gives its singer (合唱 left out).</summary>
        SingerTypeMismatch,
        /// <summary>語系類別不符 (錢櫃): 錢櫃 has the song under another 語系.</summary>
        LangMismatch,
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
                                   IEnumerable<string> specialSingerNames = null, IEnumerable<string> cashboxIds = null,
                                   string songInfoSeparator = "_", IEnumerable<string> trackWords = null,
                                   Func<string, string, string, bool> isCashboxLangMismatch = null,
                                   Func<DataRow, bool> hasLibrarySong = null)
        {
            HasLibrarySong = hasLibrarySong ?? (row => false);
            SynonymousWords = (synonymousWords ?? Enumerable.Empty<string>()).ToList();
            SpecialSingerNames = (specialSingerNames ?? Enumerable.Empty<string>()).Where(s => s != "").ToList();
            CashboxIds = new HashSet<string>(cashboxIds ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
            SongInfoSeparator = songInfoSeparator ?? "_";
            TrackWords = (trackWords ?? Enumerable.Empty<string>()).ToList();
            IsCashboxLangMismatch = isCashboxLangMismatch ?? ((lang, singer, name) => false);

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
        /// <summary>What separates the parts of a song's file name: "_" or "-" (SongMgrSongInfoSeparate).</summary>
        public string SongInfoSeparator { get; }
        /// <summary>The word for each 聲道 (Song_Track 0..5) that a song's file name carries.</summary>
        public IReadOnlyList<string> TrackWords { get; }
        /// <summary>Whether 錢櫃 has the song (語系, 歌手, 歌名) under another 語系 and not under its own.</summary>
        public Func<string, string, string, bool> IsCashboxLangMismatch { get; }
        /// <summary>Whether the 歌庫 has the 錢櫃 song in the row (Song_Lang, Song_Singer, Song_SongName), as 錢櫃比對 matches them.</summary>
        public Func<DataRow, bool> HasLibrarySong { get; }

        /// <summary>
        /// The name a singer is compared by: its 歌手群組's first name, or the singer itself when in none; trimmed and
        /// lower case.
        /// </summary>
        public string GroupOf(string singer)
        {
            string name = (singer ?? "").Trim();
            if (groupOf.TryGetValue(name, out IReadOnlyList<string> names) && names.Count > 0) name = names[0];
            return name.ToLowerInvariant();
        }

        /// <summary>The word a song's file name carries for <paramref name="track"/>.</summary>
        public string TrackWord(int track)
        {
            if (track < 0 || track > 5) track = 0;
            return track < TrackWords.Count ? TrackWords[track] : "";
        }

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

            Func<string, string, string, bool> langMismatch = null;
            if (kind == SongSearchKind.LangMismatch)
            {
                langMismatch = (lang, singer, name) =>
                {
                    int i = CommonFunc.MatchCashboxSong("CashboxLang", null, lang, singer, name, Global.CashboxSongDataLowCaseList, Global.CashboxSongDataFuzzyList);
                    if (i < 0 || lang == Global.CashboxSongDataLangList[i]) return false;
                    return Global.CashboxSongDataFullList.IndexOf(lang + "|" + singer.ToLower() + "|" + name.ToLower()) < 0;
                };
            }

            return new SongSearchReference(
                Global.SynonymousWordList,
                Global.SingerGroupList,
                (Global.SongAddSpecialStr ?? "").Split('|'),
                cashboxIds,
                (Global.SongMgrSongInfoSeparate == "1") ? "_" : "-",
                Global.CrazyktvSongTrackList,
                langMismatch);
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
            var q = new Query("ktv_Song", Columns);
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

                case SongSearchKind.MissingFile:
                    return new SongSearchResult(MissingFiles(q.OrderBy("Song_Id").Run(databaseFile)));

                case SongSearchKind.SameFile:
                    return new SongSearchResult(Groups(q.Run(databaseFile), row => (row["Song_FileName"] + "\u0001" + row["Song_Path"]).ToLowerInvariant()));

                case SongSearchKind.Duplicate:
                case SongSearchKind.DuplicateIgnoreSongType:
                case SongSearchKind.DuplicateIgnoreBrackets:
                case SongSearchKind.DuplicateIgnoreSinger:
                case SongSearchKind.DuplicateChorus:
                case SongSearchKind.SameNameIgnoreBrackets:
                {
                    DuplicateSongKey.Rule rule = RuleOf(criteria.Kind);
                    if (criteria.Kind == SongSearchKind.DuplicateChorus) q.Where("Song_SingerType = @SingerType", ("@SingerType", ChorusSingerType));
                    return new SongSearchResult(Groups(q.Run(databaseFile), row => DuplicateSongKey.Of(rule,
                        row["Song_Lang"].ToString(), row["Song_Singer"].ToString(), row["Song_SongName"].ToString(), row["Song_SongType"].ToString(), reference)));
                }

                case SongSearchKind.TrackNotInFileName:
                    q.OrderBy("Song_Id");
                    keep = row => !row["Song_FileName"].ToString().Contains(reference.SongInfoSeparator + reference.TrackWord(Convert.ToInt32(row["Song_Track"])));
                    break;

                case SongSearchKind.SingerNotInLibrary:
                {
                    var singers = new HashSet<string>(LibrarySingers(databaseFile).Keys, StringComparer.OrdinalIgnoreCase);
                    DataTable songs = q.Run(databaseFile);
                    RemoveRowsWhere(songs, row =>
                    {
                        string singer = row["Song_Singer"].ToString();
                        if (Convert.ToInt32(row["Song_SingerType"]) != ChorusSingerType) return singers.Contains(singer);
                        return reference.SplitChorus(singer).All(s => singers.Contains(s));
                    });
                    return new SongSearchResult(SortedBySinger(songs));
                }

                case SongSearchKind.SingerTypeMismatch:
                {
                    Dictionary<string, string> typeOf = LibrarySingers(databaseFile);
                    q.Where("Song_SingerType <> @SingerType", ("@SingerType", ChorusSingerType));
                    DataTable songs = q.Run(databaseFile);
                    RemoveRowsWhere(songs, row => typeOf.TryGetValue(row["Song_Singer"].ToString(), out string type) && type == row["Song_SingerType"].ToString());
                    return new SongSearchResult(SortedBySinger(songs));
                }

                case SongSearchKind.LangMismatch:
                    q.OrderBy("Song_Id");
                    keep = row => reference.IsCashboxLangMismatch(row["Song_Lang"].ToString(), row["Song_Singer"].ToString(), row["Song_SongName"].ToString());
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(criteria), criteria.Kind, "不支援的查詢種類");
            }

            DataTable found = q.Run(databaseFile);
            if (keep != null) RemoveRowsWhere(found, row => !keep(row));
            return new SongSearchResult(found);
        }

        public static DuplicateSongKey.Rule RuleOf(SongSearchKind kind)
        {
            switch (kind)
            {
                case SongSearchKind.Duplicate: return DuplicateSongKey.Rule.Full;
                case SongSearchKind.DuplicateIgnoreSongType: return DuplicateSongKey.Rule.IgnoreSongType;
                case SongSearchKind.DuplicateIgnoreBrackets: return DuplicateSongKey.Rule.IgnoreBracketsAndSongType;
                case SongSearchKind.DuplicateIgnoreSinger:
                case SongSearchKind.DuplicateChorus: return DuplicateSongKey.Rule.NameOnly;
                case SongSearchKind.SameNameIgnoreBrackets: return DuplicateSongKey.Rule.NameOnlyIgnoreBrackets;
                default: throw new ArgumentOutOfRangeException(nameof(kind), kind, "不是重複歌曲的查詢種類");
            }
        }

        /// <summary>The rows whose key some other row shares, each group together, groups in key order, by 歌曲編號 within one.</summary>
        internal static DataTable Groups(DataTable songs, Func<DataRow, string> keyOf, string idColumn = "Song_Id")
        {
            var keyed = songs.AsEnumerable().Select(row => (Key: keyOf(row), Row: row)).ToList();
            var dups = keyed.GroupBy(k => k.Key, StringComparer.Ordinal)
                            .Where(g => g.Count() > 1)
                            .OrderBy(g => g.Key, StringComparer.Ordinal)
                            .SelectMany(g => g.OrderBy(k => k.Row[idColumn].ToString(), StringComparer.Ordinal))
                            .Select(k => k.Row)
                            .ToList();

            DataTable result = songs.Clone();
            foreach (DataRow row in dups) result.ImportRow(row);
            return result;
        }

        private static DataTable SortedBySinger(DataTable songs)
        {
            DataTable result = songs.Clone();
            foreach (DataRow row in songs.AsEnumerable()
                                         .OrderBy(r => r["Song_Singer"].ToString(), StringComparer.OrdinalIgnoreCase)
                                         .ThenBy(r => r["Song_Id"].ToString(), StringComparer.Ordinal))
            {
                result.ImportRow(row);
            }
            return result;
        }

        /// <summary>ktv_Singer of the 歌庫: each singer's name and 歌手類別, whatever the case of the name.</summary>
        private static Dictionary<string, string> LibrarySingers(string databaseFile)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (DataTable dt = Db.GetDataTable(databaseFile, "select Singer_Name, Singer_Type from ktv_Singer", ""))
            {
                foreach (DataRow row in dt.Rows)
                {
                    string name = row["Singer_Name"].ToString();
                    if (!result.ContainsKey(name)) result.Add(name, row["Singer_Type"].ToString());
                }
            }
            return result;
        }

        /// <summary>The songs whose file is not on disk. Each folder is listed once instead of one File.Exists per song.</summary>
        private static DataTable MissingFiles(DataTable songs)
        {
            var filesIn = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> FilesIn(string folder)
            {
                if (!filesIn.TryGetValue(folder, out HashSet<string> files))
                {
                    files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        if (System.IO.Directory.Exists(folder))
                        {
                            foreach (string f in System.IO.Directory.EnumerateFiles(folder)) files.Add(System.IO.Path.GetFileName(f));
                        }
                    }
                    catch (Exception ex) when (ex is System.IO.IOException || ex is UnauthorizedAccessException)
                    {
                        // An unreadable folder counts as holding no files.
                    }
                    filesIn.Add(folder, files);
                }
                return files;
            }

            RemoveRowsWhere(songs, row =>
            {
                string folder, file;
                try
                {
                    string path = System.IO.Path.Combine(row["Song_Path"].ToString(), row["Song_FileName"].ToString());
                    folder = System.IO.Path.GetDirectoryName(path);
                    file = System.IO.Path.GetFileName(path);
                }
                catch (ArgumentException)
                {
                    return false;
                }
                if (string.IsNullOrEmpty(folder) || string.IsNullOrEmpty(file)) return false;
                return FilesIn(folder).Contains(file);
            });
            return songs;
        }

        /// <summary>Whether <paramref name="singer"/> (a song's Song_Singer) is <paramref name="wanted"/>, alone or in a 合唱.</summary>
        internal static bool IsSingerOf(string singer, string wanted, SongSearchReference reference)
        {
            if (singer.Equals(wanted, StringComparison.OrdinalIgnoreCase)) return true;
            return reference.SplitChorus(singer).Any(s => s.Equals(wanted.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>The song name with each of its characters that has 同義字 replaced by each of them.</summary>
        internal static IEnumerable<string> SynonymousNames(string songName, SongSearchReference reference)
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
        internal static List<string> WithWidthForms(IEnumerable<string> texts)
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

        internal static int ParseNumber(string value, string what)
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
        /// One select on one table (ktv_Song, ktv_Cashbox), written as Access SQL (rewritten for SQLite by
        /// <see cref="Db"/>) with every value bound as a parameter. Parameters are added in the order they appear, as
        /// OleDb binds them by position.
        /// </summary>
        internal sealed class Query
        {
            public Query(string table, string columns)
            {
                this.table = table;
                this.columns = columns;
            }

            private readonly string table;
            private readonly string columns;
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
                    sql.Append(columns).Append(" from ").Append(table);
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

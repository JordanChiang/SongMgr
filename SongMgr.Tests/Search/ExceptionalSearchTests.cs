using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Search
{
    /// <summary>
    /// A 歌庫 for the 異常查詢: near-duplicates that differ in case, 歌手群組, 合唱 order and separator, brackets and
    /// 歌曲類別, two songs on one file, one song whose file is missing or does not carry its 聲道, and singers in and
    /// out of ktv_Singer.
    /// </summary>
    public sealed class ExceptionalLibrary : ScratchLibrary
    {
        public ExceptionalLibrary()
        {
            Folder = Path.Combine(Root, "songs") + @"\";
            Directory.CreateDirectory(Folder);

            List<SongRecord> songs = Songs().ToList();
            foreach (SongRecord s in songs.Where(s => s.Id != "000003"))
            {
                System.IO.File.WriteAllText(Path.Combine(s.Path, s.FileName), "");
            }
            Build(songs, new (string, string)[0], new[] { ("周杰倫", "0"), ("費玉清", "0") });
        }

        public string Folder { get; private set; }

        IEnumerable<SongRecord> Songs()
        {
            yield return Song("000001", "國語", "0", "周杰倫", "晴天");
            yield return Song("000002", "國語", "0", "jay", "晴天");
            yield return Song("000003", "國語", "0", "周杰倫", "晴天(Live)", file: "000003_V9.mpg");
            yield return Song("000004", "國語", "0", "周杰倫", "晴天", type: "MV");
            yield return Song("000005", "國語", "3", "周杰倫&費玉清", "千里之外");
            yield return Song("000006", "國語", "3", "費玉清+周杰倫", "千里之外");
            yield return Song("000007", "國語", "3", "費玉清&蔡依林", "千里之外");
            yield return Song("000008", "台語", "0", "周杰倫", "晴天");
            yield return Song("000009", "國語", "0", "AB&C", "說好");
            yield return Song("000010", "國語", "3", "A&C", "說好");
            yield return Song("000011", "國語", "0", "周杰倫", "LOVE");
            yield return Song("000012", "國語", "1", "周杰倫", "love");
            yield return Song("000013", "國語", "0", "周杰倫", "同檔", file: "shared_V1.mpg");
            yield return Song("000014", "國語", "0", "周杰倫", "同檔二", file: "SHARED_V1.mpg", upperPath: true);
        }

        SongRecord Song(string id, string lang, string singerType, string singer, string name, string type = "", string file = null, bool upperPath = false) => new SongRecord
        {
            Id = id, Lang = lang, SingerType = singerType, Singer = singer, SongName = name, Track = "1", SongType = type,
            Volume = "100", WordCount = name.Length.ToString(), PlayCount = "0", MB = "12.5", CreatDate = "2024/01/01 03:04:05",
            FileName = file ?? id + "_V1.mpg", Path = upperPath ? Folder.ToUpperInvariant() : Folder,
            Spell = "", SpellNum = "", SongStroke = "0", PenStyle = "", PlayState = "0",
        };
    }

    /// <summary>
    /// The 異常查詢 kinds of <see cref="SongSearch"/> on an Access and a SQLite 歌庫, and the one definition of
    /// 重複歌曲 (<see cref="DuplicateSongKey"/>) they share.
    /// </summary>
    public class ExceptionalSearchTests : IClassFixture<ExceptionalLibrary>
    {
        readonly ExceptionalLibrary library;

        static readonly SongSearchReference Reference = new SongSearchReference(
            singerGroups: new[] { "周杰倫,JAY" },
            songInfoSeparator: "_",
            trackWords: new[] { "V0", "V1", "V2", "V3", "V4", "V5" },
            isCashboxLangMismatch: (lang, singer, name) => lang == "國語" && name == "說好");

        public ExceptionalSearchTests(ExceptionalLibrary library) { this.library = library; }

        List<string> Ids(string db, SongSearchKind kind)
        {
            DataTable songs = SongSearch.Run(new SongSearchCriteria(kind, lang: "台語"), Reference, library.File(db)).Songs;
            return songs.AsEnumerable().Select(r => r["Song_Id"].ToString()).ToList();
        }

        public static IEnumerable<object[]> Both => new[] { new object[] { "Access" }, new object[] { "SQLite" } };

        [Theory, MemberData(nameof(Both))]
        public void Duplicate_compares_name_singer_set_and_type_whatever_the_case(string db)
        {
            // LOVE / love by the same singer; a 合唱 in another order and with "+"; 周杰倫 and JAY of one 歌手群組.
            // AB&C and A&C are not the same singers; 晴天(Live), 晴天 MV and 台語 晴天 stay apart.
            Assert.Equal(new[] { "000011", "000012", "000005", "000006", "000001", "000002" }, Ids(db, SongSearchKind.Duplicate));
        }

        [Theory, MemberData(nameof(Both))]
        public void Duplicate_ignoring_song_type(string db)
        {
            Assert.Equal(new[] { "000011", "000012", "000005", "000006", "000001", "000002", "000004" }, Ids(db, SongSearchKind.DuplicateIgnoreSongType));
        }

        [Theory, MemberData(nameof(Both))]
        public void Duplicate_ignoring_brackets_and_song_type(string db)
        {
            Assert.Equal(new[] { "000011", "000012", "000005", "000006", "000001", "000002", "000003", "000004" }, Ids(db, SongSearchKind.DuplicateIgnoreBrackets));
        }

        [Theory, MemberData(nameof(Both))]
        public void Duplicate_ignoring_singer_and_song_type(string db)
        {
            Assert.Equal(new[] { "000011", "000012", "000005", "000006", "000007", "000001", "000002", "000004", "000009", "000010" }, Ids(db, SongSearchKind.DuplicateIgnoreSinger));
        }

        [Theory, MemberData(nameof(Both))]
        public void Duplicate_chorus_songs_by_name(string db)
        {
            // Each song once: the old query added a 合唱 found twice.
            Assert.Equal(new[] { "000005", "000006", "000007" }, Ids(db, SongSearchKind.DuplicateChorus));
        }

        [Theory, MemberData(nameof(Both))]
        public void Same_name_ignoring_brackets(string db)
        {
            Assert.Equal(new[] { "000011", "000012", "000005", "000006", "000007", "000001", "000002", "000003", "000004", "000009", "000010" }, Ids(db, SongSearchKind.SameNameIgnoreBrackets));
        }

        [Theory, MemberData(nameof(Both))]
        public void Same_file_whatever_the_case(string db)
        {
            Assert.Equal(new[] { "000013", "000014" }, Ids(db, SongSearchKind.SameFile));
        }

        [Theory, MemberData(nameof(Both))]
        public void Missing_file(string db)
        {
            Assert.Equal(new[] { "000003" }, Ids(db, SongSearchKind.MissingFile));
        }

        [Theory, MemberData(nameof(Both))]
        public void Track_not_in_file_name(string db)
        {
            Assert.Equal(new[] { "000003" }, Ids(db, SongSearchKind.TrackNotInFileName));
        }

        [Theory, MemberData(nameof(Both))]
        public void Singer_not_in_library(string db)
        {
            // A 合唱 counts when any of its singers is missing; ordered by singer.
            Assert.Equal(new[] { "000010", "000009", "000002", "000007" }, Ids(db, SongSearchKind.SingerNotInLibrary));
        }

        [Theory, MemberData(nameof(Both))]
        public void Singer_type_mismatch(string db)
        {
            Assert.Equal(new[] { "000009", "000002", "000012" }, Ids(db, SongSearchKind.SingerTypeMismatch));
        }

        [Theory, MemberData(nameof(Both))]
        public void Lang_mismatch_asks_the_cashbox_reference(string db)
        {
            Assert.Equal(new[] { "000009", "000010" }, Ids(db, SongSearchKind.LangMismatch));
        }

        [Theory]
        [InlineData("周杰倫&費玉清", "費玉清 + 周杰倫")]
        [InlineData("JAY", "周杰倫")]
        [InlineData("JAY&費玉清", "費玉清&周杰倫")]
        [InlineData("周杰倫&周杰倫", "周杰倫")]
        public void Singer_key_is_the_set_of_singers(string a, string b)
        {
            Assert.Equal(DuplicateSongKey.SingerKey(a, Reference), DuplicateSongKey.SingerKey(b, Reference));
        }

        [Theory]
        [InlineData("AB&C", "A&C")]
        [InlineData("周杰倫&費玉清", "周杰倫")]
        [InlineData("周杰倫&費玉清", "周杰倫&費玉清&蔡依林")]
        public void Different_singers_are_different_keys(string a, string b)
        {
            Assert.NotEqual(DuplicateSongKey.SingerKey(a, Reference), DuplicateSongKey.SingerKey(b, Reference));
        }

        [Fact]
        public void Index_finds_duplicates_in_a_growing_list()
        {
            var list = new List<string> { "國語|周杰倫&費玉清|千里之外|", "國語|周杰倫|晴天|mv" };
            var index = new DuplicateSongKey.Index(list, Reference);

            Assert.Equal(new[] { 0 }, index.Find("國語", "費玉清+周杰倫", "千里之外", ""));
            Assert.Equal(new[] { 1 }, index.Find("國語", "JAY", " 晴天 ", "MV"));
            Assert.Empty(index.Find("國語", "周杰倫", "晴天", ""));

            list.Add("國語|jay|晴天|");
            list.Add("國語|周杰倫|晴天|");
            Assert.Equal(new[] { 2, 3 }, index.Find("國語", "周杰倫", "晴天", ""));
            Assert.Empty(index.Find("國語", "AB&C", "說好", ""));
        }

        [Fact]
        public void Empty_song_type_matches_only_empty()
        {
            Assert.NotEqual(DuplicateSongKey.Of(DuplicateSongKey.Rule.Full, "國語", "周杰倫", "晴天", "", Reference),
                            DuplicateSongKey.Of(DuplicateSongKey.Rule.Full, "國語", "周杰倫", "晴天", "MV", Reference));
            Assert.Equal(DuplicateSongKey.Of(DuplicateSongKey.Rule.Full, "國語", "周杰倫", " 晴天 ", "mv", Reference),
                         DuplicateSongKey.Of(DuplicateSongKey.Rule.Full, "國語", "周杰倫", "晴天", "MV", Reference));
        }
    }
}

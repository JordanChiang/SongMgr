using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Search
{
    /// <summary>
    /// One 歌庫 written to CrazySong.mdb and exported to CrazySong.db, searched through both: the songs are chosen to
    /// differ in case, width, 合唱 order and separator, 歌手群組 and 同義字.
    /// </summary>
    public sealed class SearchLibrary : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("search");

        public SearchLibrary()
        {
            string savedSong = Global.CrazyktvDatabaseFile, savedMgr = Global.CrazyktvSongMgrDatabaseFile;
            TestEnv.UseScratchDatabases(root);
            Access = Global.CrazyktvDatabaseFile;
            Global.CrazyktvDatabaseFile = savedSong;
            Global.CrazyktvSongMgrDatabaseFile = savedMgr;

            using (SongLibraryStore store = SongLibraryStore.OpenFile(Access))
            {
                foreach (SongRecord s in Songs()) store.InsertSong(s, SongColumns.AllButReplayGain);
            }
            using (DbConnection conn = Db.Open(Access))
            {
                foreach (string id in new[] { "000002", "000005" })
                {
                    using (DbCommand cmd = Db.Command("insert into ktv_Favorite (User_Id, Song_Id) values ('0001', @SongId)", conn))
                    {
                        cmd.Parameters.AddWithValue("@SongId", id);
                        cmd.ExecuteNonQuery();
                    }
                }
                using (DbCommand cmd = Db.Command("insert into ktv_Favorite (User_Id, Song_Id) values ('0002', '000001')", conn)) cmd.ExecuteNonQuery();
            }

            SQLite = Path.Combine(root, "CrazySong.db");
            var results = SqliteExport.Run(new[] { new SqliteExport.Job(Access, SQLite, SqliteExport.CrazySongDBVer) });
            Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());
        }

        public string Access { get; }
        public string SQLite { get; }

        public string File(string kind) => kind == "Access" ? Access : SQLite;

        public void Dispose()
        {
            System.Data.SQLite.SQLiteConnection.ClearAllPools();
            TestEnv.DeleteDir(root);
        }

        static IEnumerable<SongRecord> Songs()
        {
            yield return Song("000001", "國語", "0", "周杰倫", "晴天", created: "2024/01/01");
            yield return Song("000002", "國語", "1", "蔡依林", "日不落", type: "MV", created: "2024/01/02");
            yield return Song("000003", "國語", "0", "周杰倫", "Love Song", created: "2024/01/03");
            yield return Song("000004", "國語", "0", "JAY", "七里香", created: "2024/01/04");
            yield return Song("000005", "國語", "3", "周杰倫&費玉清", "千里之外", track: "2", created: "2024/01/05");
            yield return Song("000006", "國語", "3", "費玉清+周杰倫", "千里之外", created: "2024/01/06");
            yield return Song("000007", "國語", "2", "S.H.E", "不想長大", created: "2024/01/07");
            yield return Song("000008", "國語", "2", "Ｆ４", "流星雨", created: "2024/01/08");
            yield return Song("000009", "國語", "0", "周杰倫", "love song", type: "Live", created: "2024/01/09");
            yield return Song("000010", "國語", "2", "Me&You", "你我", created: "2024/01/10");
            yield return Song("000011", "國語", "0", "江蕙", "臺北的天空", created: "2024/01/11");
            yield return Song("000012", "國語", "0", "O'Neil", "It's me", created: "2024/01/12");
            yield return Song("100001", "台語", "3", "江蕙&陳雷", "愛人醉落去", created: "2024/01/13");
        }

        static SongRecord Song(string id, string lang, string singerType, string singer, string name, string type = "", string track = "1", string created = "2024/01/01") => new SongRecord
        {
            Id = id, Lang = lang, SingerType = singerType, Singer = singer, SongName = name, Track = track, SongType = type,
            Volume = "100", WordCount = name.Length.ToString(), PlayCount = "0", MB = "12.5", CreatDate = created + " 03:04:05",
            FileName = id + ".mpg", Path = @"D:\KTV\" + lang + @"\", Spell = "", SpellNum = "", SongStroke = "0", PenStyle = "", PlayState = "0",
        };
    }

    /// <summary>
    /// 歌庫查詢 (<see cref="SongSearch"/>) on an Access and a SQLite 歌庫: every kind finds the same songs on both,
    /// text compares without regard to case, and values are parameters, never SQL.
    /// </summary>
    public class SongSearchTests : IClassFixture<SearchLibrary>
    {
        readonly SearchLibrary library;

        static readonly SongSearchReference Reference = new SongSearchReference(
            synonymousWords: new[] { "台|臺" },
            singerGroups: new[] { "周杰倫,JAY" },
            specialSingerNames: new[] { "Me&You" },
            cashboxIds: new[] { "000001", "000005", "999999" });

        public SongSearchTests(SearchLibrary library) { this.library = library; }

        List<string> Ids(string db, SongSearchKind kind, string value = "", string lang = null, bool fuzzy = true, bool synonyms = false)
        {
            SongSearchResult result = SongSearch.Run(new SongSearchCriteria(kind, value, lang, fuzzy, synonyms), Reference, library.File(db));
            Assert.Equal(result.Count, result.Songs.Rows.Count);
            return result.Songs.AsEnumerable().Select(r => r["Song_Id"].ToString()).ToList();
        }

        public static IEnumerable<object[]> Both => new[] { new object[] { "Access" }, new object[] { "SQLite" } };

        [Theory, MemberData(nameof(Both))]
        public void SongName_fuzzy_ignores_case(string db)
        {
            Assert.Equal(new[] { "000003", "000009" }, Ids(db, SongSearchKind.SongName, "LOVE"));
        }

        [Theory, MemberData(nameof(Both))]
        public void SongName_exact_ignores_case(string db)
        {
            Assert.Equal(new[] { "000003", "000009" }, Ids(db, SongSearchKind.SongName, "love song", fuzzy: false));
            Assert.Empty(Ids(db, SongSearchKind.SongName, "love", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void SongName_with_synonyms_finds_the_other_word(string db)
        {
            Assert.Empty(Ids(db, SongSearchKind.SongName, "台北"));
            Assert.Equal(new[] { "000011" }, Ids(db, SongSearchKind.SongName, "台北", synonyms: true));
        }

        [Theory, MemberData(nameof(Both))]
        public void Fuzzy_search_tries_half_and_full_width(string db)
        {
            Assert.Equal(new[] { "000008" }, Ids(db, SongSearchKind.SingerName, "F4"));
            Assert.Equal(new[] { "000007" }, Ids(db, SongSearchKind.SingerName, "Ｓ.Ｈ.Ｅ"));
        }

        [Theory, MemberData(nameof(Both))]
        public void Quotes_in_the_value_are_text_not_sql(string db)
        {
            Assert.Equal(new[] { "000012" }, Ids(db, SongSearchKind.SingerName, "o'neil"));
            Assert.Equal(new[] { "000012" }, Ids(db, SongSearchKind.SongName, "It's me", fuzzy: false));
            Assert.Empty(Ids(db, SongSearchKind.SongName, "x' or '1'='1", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void SingerName_fuzzy_includes_the_singer_group(string db)
        {
            Assert.Equal(new[] { "000001", "000003", "000004", "000005", "000006", "000009" }, Ids(db, SongSearchKind.SingerName, "周杰倫"));
            Assert.Equal(new[] { "000001", "000003", "000004", "000005", "000006", "000009" }, Ids(db, SongSearchKind.SingerName, "jay"));
        }

        [Theory, MemberData(nameof(Both))]
        public void SingerName_exact_finds_the_singer_alone_or_in_a_chorus(string db)
        {
            // Both 合唱 separators count; the 歌手群組 does not apply to an exact search.
            Assert.Equal(new[] { "000001", "000003", "000005", "000006", "000009" }, Ids(db, SongSearchKind.SingerName, "周杰倫", fuzzy: false));
            Assert.Equal(new[] { "000005", "000006" }, Ids(db, SongSearchKind.SingerName, "費玉清", fuzzy: false));
            Assert.Equal(new[] { "000007" }, Ids(db, SongSearchKind.SingerName, "s.h.e", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void SingerName_exact_keeps_special_names_whole(string db)
        {
            Assert.Empty(Ids(db, SongSearchKind.SingerName, "Me", fuzzy: false));
            Assert.Equal(new[] { "000010" }, Ids(db, SongSearchKind.SingerName, "Me&You", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void Lang_filter_applies(string db)
        {
            Assert.Equal(new[] { "100001" }, Ids(db, SongSearchKind.AllSongs, lang: "台語"));
            Assert.Equal(new[] { "000011", "100001" }, Ids(db, SongSearchKind.SingerName, "江蕙"));
            Assert.Equal(new[] { "000011" }, Ids(db, SongSearchKind.SingerName, "江蕙", lang: "國語"));
            Assert.Equal(13, Ids(db, SongSearchKind.AllSongs, lang: SongSearchCriteria.AllLangs).Count);
        }

        [Theory, MemberData(nameof(Both))]
        public void SongId_fuzzy_and_exact(string db)
        {
            Assert.Equal(new[] { "000001", "000010", "000011", "000012", "100001" }, Ids(db, SongSearchKind.SongId, "1"));
            Assert.Equal(new[] { "000005" }, Ids(db, SongSearchKind.SongId, "000005", fuzzy: false));
            // An exact 歌曲編號 ignores the 語系 filter.
            Assert.Equal(new[] { "000005" }, Ids(db, SongSearchKind.SongId, "000005", lang: "台語", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void NewSongs_are_the_latest_first(string db)
        {
            Assert.Equal(new[] { "100001", "000012", "000011" }, Ids(db, SongSearchKind.NewSongs, "3"));
            Assert.Equal(new[] { "000012", "000011" }, Ids(db, SongSearchKind.NewSongs, "2", lang: "國語"));
        }

        [Theory]
        [InlineData("0")]
        [InlineData("-1")]
        [InlineData("abc")]
        [InlineData("1; drop table ktv_Song")]
        public void NewSongs_needs_a_positive_count(string count)
        {
            Assert.Throws<ArgumentException>(() => SongSearch.Run(new SongSearchCriteria(SongSearchKind.NewSongs, count), Reference, library.Access));
        }

        [Theory, MemberData(nameof(Both))]
        public void ChorusSongs(string db)
        {
            Assert.Equal(new[] { "000005", "000006", "100001" }, Ids(db, SongSearchKind.ChorusSongs));
        }

        [Theory, MemberData(nameof(Both))]
        public void SongType_ignores_case_and_empty_means_none(string db)
        {
            Assert.Equal(new[] { "000009" }, Ids(db, SongSearchKind.SongType, "live"));
            Assert.Equal(11, Ids(db, SongSearchKind.SongType, "").Count);
        }

        [Theory, MemberData(nameof(Both))]
        public void SingerType_and_SongTrack(string db)
        {
            Assert.Equal(new[] { "000002" }, Ids(db, SongSearchKind.SingerType, "1"));
            Assert.Equal(new[] { "000005" }, Ids(db, SongSearchKind.SongTrack, "2"));
            Assert.Throws<ArgumentException>(() => Ids(db, SongSearchKind.SongTrack, "2 or 1=1"));
        }

        [Theory, MemberData(nameof(Both))]
        public void CashboxId_keeps_songs_on_a_cashbox_number(string db)
        {
            Assert.Equal(new[] { "000001", "000005" }, Ids(db, SongSearchKind.CashboxId));
        }

        [Theory, MemberData(nameof(Both))]
        public void Favorite_lists_the_users_songs(string db)
        {
            Assert.Equal(new[] { "000002", "000005" }, Ids(db, SongSearchKind.Favorite, "0001"));
            Assert.Equal(new[] { "000001" }, Ids(db, SongSearchKind.Favorite, "0002"));
            Assert.Empty(Ids(db, SongSearchKind.Favorite, "9999"));
        }

        [Fact]
        public void Every_result_has_the_columns_the_list_shows()
        {
            DataTable access = SongSearch.Run(new SongSearchCriteria(SongSearchKind.AllSongs), Reference, library.Access).Songs;
            DataTable sqlite = SongSearch.Run(new SongSearchCriteria(SongSearchKind.AllSongs), Reference, library.SQLite).Songs;
            string[] expected = SongSearch.Columns.Split(',').Select(c => c.Trim()).ToArray();
            Assert.Equal(expected, access.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
            Assert.Equal(expected, sqlite.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        }
    }
}

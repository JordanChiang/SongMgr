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
    /// A 參考資料庫 whose ktv_Cashbox holds only the songs below, as CrazySongMgr.mdb and exported to CrazySongMgr.db.
    /// </summary>
    public sealed class CashboxLibrary : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("cashbox");

        public CashboxLibrary()
        {
            string savedSong = Global.CrazyktvDatabaseFile, savedMgr = Global.CrazyktvSongMgrDatabaseFile;
            TestEnv.UseScratchDatabases(root);
            Access = Global.CrazyktvSongMgrDatabaseFile;
            Global.CrazyktvDatabaseFile = savedSong;
            Global.CrazyktvSongMgrDatabaseFile = savedMgr;

            using (DbConnection conn = Db.Open(Access))
            {
                using (DbCommand cmd = Db.Command("delete * from ktv_Cashbox", conn)) cmd.ExecuteNonQuery();
                foreach (var (id, lang, singer, name, date) in Songs)
                {
                    using (DbCommand cmd = Db.Command("insert into ktv_Cashbox (Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate, Song_HaveSong) values (@Id, @Lang, @Singer, @Name, @Date, @Have)", conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", id);
                        cmd.Parameters.AddWithValue("@Lang", lang);
                        cmd.Parameters.AddWithValue("@Singer", singer);
                        cmd.Parameters.AddWithValue("@Name", name);
                        cmd.Parameters.AddWithValue("@Date", date);
                        cmd.Parameters.AddWithValue("@Have", false);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            SQLite = Path.Combine(root, "CrazySongMgr.db");
            var results = SqliteExport.Run(new[] { new SqliteExport.Job(Access, SQLite, SqliteExport.CrazySongMgrDBVer, true) });
            Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());
        }

        static readonly (string Id, string Lang, string Singer, string Name, DateTime Date)[] Songs =
        {
            ("00001", "國語", "周杰倫", "晴天", new DateTime(2024, 1, 1)),
            ("00002", "國語", "JAY", "晴天", new DateTime(2024, 1, 1, 23, 59, 59)),
            ("00003", "國語", "周杰倫", "Love Song", new DateTime(2024, 1, 2)),
            ("00004", "國語", "周杰倫", "love song", new DateTime(2024, 1, 2)),
            ("00005", "國語", "周杰倫&費玉清", "千里之外", new DateTime(2024, 1, 3)),
            ("00006", "國語", "費玉清+周杰倫", "千里之外", new DateTime(2024, 1, 3)),
            ("00007", "國語", "O'Neil", "It's me", new DateTime(2024, 1, 4)),
            ("00008", "國語", "Ｆ４", "流星雨", new DateTime(2024, 1, 5)),
            ("00009", "國語", "江蕙", "臺北的天空", new DateTime(2024, 1, 6)),
            ("10001", "台語", "江蕙&陳雷", "愛人醉落去", new DateTime(2024, 1, 7)),
        };

        public string Access { get; }
        public string SQLite { get; }

        public string File(string kind) => kind == "Access" ? Access : SQLite;

        public void Dispose()
        {
            System.Data.SQLite.SQLiteConnection.ClearAllPools();
            TestEnv.DeleteDir(root);
        }
    }

    /// <summary>錢櫃查詢 (<see cref="CashboxSearch"/>) on an Access and a SQLite 參考資料庫: every kind finds the same songs on both.</summary>
    public class CashboxSearchTests : IClassFixture<CashboxLibrary>
    {
        readonly CashboxLibrary library;

        static readonly SongSearchReference Reference = new SongSearchReference(
            synonymousWords: new[] { "台|臺" },
            singerGroups: new[] { "周杰倫,JAY" },
            hasLibrarySong: row => row["Song_Singer"].ToString() != "江蕙");

        public CashboxSearchTests(CashboxLibrary library) { this.library = library; }

        List<string> Ids(string db, CashboxSearchKind kind, string value = "", string lang = null, bool fuzzy = true, bool synonyms = false, DateTime date = default)
        {
            DataTable songs = CashboxSearch.Run(new CashboxSearchCriteria(kind, value, lang, fuzzy, synonyms, date), Reference, library.File(db)).Songs;
            return songs.AsEnumerable().Select(r => r["Cashbox_Id"].ToString()).ToList();
        }

        public static IEnumerable<object[]> Both => new[] { new object[] { "Access" }, new object[] { "SQLite" } };

        [Theory, MemberData(nameof(Both))]
        public void SongName_fuzzy_exact_and_synonyms(string db)
        {
            Assert.Equal(new[] { "00003", "00004" }, Ids(db, CashboxSearchKind.SongName, "LOVE"));
            Assert.Equal(new[] { "00003", "00004" }, Ids(db, CashboxSearchKind.SongName, "love song", fuzzy: false));
            Assert.Empty(Ids(db, CashboxSearchKind.SongName, "台北"));
            Assert.Equal(new[] { "00009" }, Ids(db, CashboxSearchKind.SongName, "台北", synonyms: true));
        }

        [Theory, MemberData(nameof(Both))]
        public void SingerName_fuzzy_includes_the_singer_group_and_width_forms(string db)
        {
            Assert.Equal(new[] { "00001", "00002", "00003", "00004", "00005", "00006" }, Ids(db, CashboxSearchKind.SingerName, "jay"));
            Assert.Equal(new[] { "00008" }, Ids(db, CashboxSearchKind.SingerName, "F4"));
        }

        [Theory, MemberData(nameof(Both))]
        public void SingerName_exact_finds_the_singer_in_a_chorus(string db)
        {
            Assert.Equal(new[] { "00005", "00006" }, Ids(db, CashboxSearchKind.SingerName, "費玉清", fuzzy: false));
            Assert.Equal(new[] { "00009", "10001" }, Ids(db, CashboxSearchKind.SingerName, "江蕙", fuzzy: false));
            Assert.Equal(new[] { "00009" }, Ids(db, CashboxSearchKind.SingerName, "江蕙", lang: "國語", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void Quotes_are_text_not_sql(string db)
        {
            Assert.Equal(new[] { "00007" }, Ids(db, CashboxSearchKind.SingerName, "o'neil"));
            Assert.Equal(new[] { "00007" }, Ids(db, CashboxSearchKind.SongName, "It's me", fuzzy: false));
            Assert.Empty(Ids(db, CashboxSearchKind.SongName, "x' or '1'='1", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void CashboxId_fuzzy_and_exact(string db)
        {
            Assert.Equal(new[] { "00001", "10001" }, Ids(db, CashboxSearchKind.CashboxId, "0001"));
            Assert.Equal(new[] { "00001" }, Ids(db, CashboxSearchKind.CashboxId, "0001", lang: "國語"));
            Assert.Equal(new[] { "10001" }, Ids(db, CashboxSearchKind.CashboxId, "10001", lang: "國語", fuzzy: false));
        }

        [Theory, MemberData(nameof(Both))]
        public void AllSongs_and_NewSongs(string db)
        {
            Assert.Equal(10, Ids(db, CashboxSearchKind.AllSongs).Count);
            Assert.Equal(new[] { "10001" }, Ids(db, CashboxSearchKind.AllSongs, lang: "台語"));
            Assert.Equal(new[] { "10001", "00009" }, Ids(db, CashboxSearchKind.NewSongs, "2"));
            Assert.Throws<ArgumentException>(() => Ids(db, CashboxSearchKind.NewSongs, "2 or 1=1"));
        }

        [Theory, MemberData(nameof(Both))]
        public void Missing_lists_songs_the_library_does_not_have(string db)
        {
            Assert.Equal(new[] { "00009" }, Ids(db, CashboxSearchKind.Missing, lang: "台語"));
        }

        [Theory, MemberData(nameof(Both))]
        public void Duplicate_uses_the_one_definition(string db)
        {
            // love song / Love Song, a 合唱 in another order, 周杰倫 and JAY of one 歌手群組; each group together.
            Assert.Equal(new[] { "00003", "00004", "00005", "00006", "00001", "00002" }, Ids(db, CashboxSearchKind.Duplicate));
        }

        [Theory, MemberData(nameof(Both))]
        public void UpdatedOn_is_the_whole_day(string db)
        {
            Assert.Equal(new[] { "00001", "00002" }, Ids(db, CashboxSearchKind.UpdatedOn, date: new DateTime(2024, 1, 1)));
            Assert.Equal(new[] { "00005", "00006" }, Ids(db, CashboxSearchKind.UpdatedOn, date: new DateTime(2024, 1, 3, 15, 0, 0)));
            Assert.Empty(Ids(db, CashboxSearchKind.UpdatedOn, date: new DateTime(2023, 12, 31)));
        }
    }
}

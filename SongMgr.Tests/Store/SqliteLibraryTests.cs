using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Store
{
    /// <summary>
    /// A SQLite 歌庫: CrazySongEmpty.mdb exported to CrazySong.db the way 匯出 SQLite 資料庫 does, then loaded
    /// as Global.CrazyktvDatabaseFile and read and written through <see cref="Db"/>, <see cref="SongLibraryStore"/>
    /// and <see cref="YoutubeSongStore"/>. Also the Access SQL rewrites and picking the provider by extension.
    /// </summary>
    public class SqliteLibraryTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("sqlite");
        readonly string savedSong = Global.CrazyktvDatabaseFile;
        readonly string savedMgr = Global.CrazyktvSongMgrDatabaseFile;
        readonly string db;

        public SqliteLibraryTests()
        {
            TestEnv.UseScratchDatabases(root);
            string mdb = Global.CrazyktvDatabaseFile;
            db = Path.Combine(root, "CrazySong.db");
            var results = SqliteExport.Run(new[] { new SqliteExport.Job(mdb, db, SqliteExport.CrazySongDBVer) });
            Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());
            Global.CrazyktvDatabaseFile = db;
        }

        public void Dispose()
        {
            System.Data.SQLite.SQLiteConnection.ClearAllPools();
            Global.CrazyktvDatabaseFile = savedSong;
            Global.CrazyktvSongMgrDatabaseFile = savedMgr;
            TestEnv.DeleteDir(root);
        }

        [Theory]
        [InlineData(@"D:\KTV\CrazySong.mdb", false)]
        [InlineData(@"D:\KTV\CrazySong.db", true)]
        [InlineData(@"D:\KTV\CrazySong.DB", true)]
        [InlineData(@"D:\KTV\CrazySong.sqlite", true)]
        [InlineData(@"D:\KTV\CrazySong.sqlite3", true)]
        [InlineData(@"D:\KTV\CrazySong", false)]
        public void IsSQLite_PicksProviderByExtension(string file, bool expected)
        {
            Assert.Equal(expected, Db.IsSQLite(file));
        }

        [Theory]
        [InlineData("delete * from ktv_Singer", "delete from ktv_Singer")]
        [InlineData("select * from ktv_Song where InStr(1,LCase(Song_SongName),LCase('Love (Live), 版'),0) <>0",
                    "select * from ktv_Song where instr(lower(Song_SongName), lower('Love (Live), 版')) <>0")]
        [InlineData("select top 10 Song_Id from ktv_Song order by Song_CreatDate desc", "select Song_Id from ktv_Song order by Song_CreatDate desc limit 10")]
        [InlineData("update ktv_Song set [Song_Volume] = @SongVolume where Song_Id = @SongId", "update ktv_Song set [Song_Volume] = @SongVolume where Song_Id = @SongId")]
        public void ToSQLite_RewritesAccessOnlySql(string access, string sqlite)
        {
            Assert.Equal(sqlite, Db.ToSQLite(access));
        }

        [Fact]
        public void ExportedLibrary_HasTheTablesAndColumnsTheStatusCheckNeeds()
        {
            List<string> tables = Db.GetTableList(db);
            foreach (string t in new[] { "ktv_Favorite", "ktv_Langauage", "ktv_Phonetics", "ktv_Remote", "ktv_Singer", "ktv_SingerName", "ktv_Song", "ktv_Swan", "ktv_User", "ktv_Words" })
            {
                Assert.Contains(t, tables);
            }
            Assert.Contains("Song_ReplayGain", Db.GetColumnList(db, "", "ktv_Song"));
            Assert.Contains("Langauage_KeyWord", Db.GetColumnList(db, "", "ktv_Langauage"));
        }

        [Fact]
        public void SongLibraryStore_InsertsUpdatesAndDeletesSongs()
        {
            using (SongLibraryStore store = SongLibraryStore.OpenOrThrow())
            {
                Assert.Equal(1, store.InsertSong(Song("000001", "晴天"), SongColumns.AllButReplayGain));
                Assert.Equal(1, store.InsertSong(Song("000002", "Love Story"), SongColumns.AllButReplayGain));
                Assert.Equal(1, store.UpdateSong(new SongRecord { Id = "000001", Volume = "77", Track = "2" }, SongColumns.Volume | SongColumns.Track));
                Assert.Equal(1, store.DeleteSong("000002"));
            }

            DataTable dt = Db.GetDataTable(db, "select * from ktv_Song");
            DataRow row = dt.Rows.Cast<DataRow>().Single();
            Assert.Equal("000001", row["Song_Id"]);
            Assert.Equal(77, Convert.ToInt32(row["Song_Volume"]));
            Assert.Equal(2, Convert.ToInt32(row["Song_Track"]));
            Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5), row.Field<DateTime>("Song_CreatDate"));
        }

        [Fact]
        public void InStrQuery_FindsSongsIgnoringCase()
        {
            using (SongLibraryStore store = SongLibraryStore.OpenOrThrow())
            {
                store.InsertSong(Song("000001", "Love Story"), SongColumns.AllButReplayGain);
                store.InsertSong(Song("000002", "晴天"), SongColumns.AllButReplayGain);
            }

            DataTable dt = Db.GetDataTable(db, "select Song_Id from ktv_Song where InStr(1,LCase(Song_SongName),LCase('LOVE'),0) <>0");

            Assert.Equal(new[] { "000001" }, dt.AsEnumerable().Select(r => r["Song_Id"].ToString()));
        }

        [Fact]
        public void Singers_DeleteAllAndMaxIdWorkOnSqlite()
        {
            using (SongLibraryStore store = SongLibraryStore.Open(SingerTable.Library))
            {
                store.DeleteAllSingers(SingerTable.Library);
                store.InsertSinger(SingerTable.Library, "3", "周杰倫", "0", "ㄓㄐㄌ", "8", "524", "丨一");
                store.InsertSinger(SingerTable.Library, "7", "蔡依林", "1", "ㄘㄧㄌ", "17", "612", "一丨");
            }

            Assert.Equal(7, CommonFunc.GetMaxSingerId("ktv_Singer", db));
            Assert.Equal(new[] { "周杰倫", "蔡依林" }, Db.GetDataTable(db, "select Singer_Name from ktv_Singer order by Singer_Id").AsEnumerable().Select(r => r["Singer_Name"].ToString()));
        }

        [Fact]
        public void YoutubeSongStore_CreatesEditsAndDeletesOnSqlite()
        {
            Assert.False(YoutubeSongStore.TableExists(db));
            YoutubeSongStore.CreateTable(db);
            Assert.True(YoutubeSongStore.TableExists(db));

            using (var conn = Db.Open(db))
            {
                using (var cmd = Db.Command("insert into YoutubeSongList (Id, Name, Url, PlayCount, Volume) values ('Y0001', 'Sistar - Touch My Body', 'https://youtu.be/x', 0, 80)", conn)) cmd.ExecuteNonQuery();
            }

            Assert.Equal(3, Convert.ToInt32(YoutubeSongStore.Query(db, "").Rows[0]["Track"]));
            Assert.Equal(1, YoutubeSongStore.Update(db, "Y0001", YoutubeColumn.Volume, "55"));
            Assert.Equal(55, Convert.ToInt32(YoutubeSongStore.Query(db, "sistar").Rows[0]["Volume"]));
            Assert.Equal(1, YoutubeSongStore.Delete(db, new[] { "Y0001" }));
            Assert.Empty(YoutubeSongStore.Query(db, "").Rows.Cast<DataRow>());
        }

        [Fact]
        public void SqliteExport_SkipsALibraryThatIsAlreadySqlite()
        {
            List<SqliteExport.Job> jobs = SqliteExport.Jobs(db, Global.CrazyktvSongMgrDatabaseFile);

            Assert.Equal(new[] { Path.Combine(root, "CrazySongMgr.db") }, jobs.Select(j => j.Dest));
            Assert.Empty(SqliteExport.Jobs(db, Path.Combine(root, "CrazySongMgr.db")));
        }

        [Fact]
        public void Compact_VacuumsSqlite()
        {
            Db.Compact(db);

            Assert.Contains("ktv_Song", Db.GetTableList(db));
        }

        [Fact]
        public void Open_MissingSqliteFileThrowsWithoutCreatingIt()
        {
            string missing = Path.Combine(root, "missing.db");

            Assert.Throws<FileNotFoundException>(() => Db.Open(missing));
            Assert.False(File.Exists(missing));
        }

        static SongRecord Song(string id, string name) => new SongRecord
        {
            Id = id, Lang = "國語", SingerType = "0", Singer = "周杰倫", SongName = name, Track = "1", SongType = "",
            Volume = "100", WordCount = "2", PlayCount = "0", MB = "12.5", CreatDate = "2024/01/02 03:04:05",
            FileName = "周杰倫_" + name + ".mpg", Path = @"D:\KTV\國語\", Spell = "ㄑㄊ", SpellNum = "42",
            SongStroke = "12", PenStyle = "丨一", PlayState = "0",
        };
    }
}

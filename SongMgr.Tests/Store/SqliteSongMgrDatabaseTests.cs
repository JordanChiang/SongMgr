using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Store
{
    /// <summary>
    /// CrazySongMgr.db: the bundled CrazySongMgr.mdb exported to SQLite and used as
    /// Global.CrazyktvSongMgrDatabaseFile. Pairing it with a SQLite 歌庫 only, reading the 注音 table,
    /// writing 錢櫃 dates, and the 錢櫃 date query.
    /// </summary>
    public class SqliteSongMgrDatabaseTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("sqlitemgr");
        readonly string savedSong = Global.CrazyktvDatabaseFile;
        readonly string savedMgr = Global.CrazyktvSongMgrDatabaseFile;
        readonly string mdb;
        readonly string db;

        public SqliteSongMgrDatabaseTests()
        {
            TestEnv.UseScratchDatabases(root);
            mdb = Global.CrazyktvSongMgrDatabaseFile;
            db = Path.Combine(root, "CrazySongMgr.db");
            var results = SqliteExport.Run(new[] { new SqliteExport.Job(mdb, db, SqliteExport.CrazySongMgrDBVer) });
            Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());
        }

        public void Dispose()
        {
            System.Data.SQLite.SQLiteConnection.ClearAllPools();
            Global.CrazyktvDatabaseFile = savedSong;
            Global.CrazyktvSongMgrDatabaseFile = savedMgr;
            Global.PhoneticsWordList = new List<string>();
            TestEnv.DeleteDir(root);
        }

        [Fact]
        public void SongMgrDatabaseFor_PairsTheSameKindAsTheSongDatabase()
        {
            // Both CrazySongMgr.mdb and CrazySongMgr.db exist: an Access 歌庫 keeps the .mdb.
            Assert.Equal(mdb, Db.SongMgrDatabaseFor(Path.Combine(root, "CrazySong.mdb"), mdb));
            Assert.Equal(mdb, Db.SongMgrDatabaseFor(Path.Combine(root, "CrazySong.mdb"), db));
            Assert.Equal(db, Db.SongMgrDatabaseFor(Path.Combine(root, "CrazySong.db"), mdb));
            Assert.Equal(db, Db.SongMgrDatabaseFor(Path.Combine(root, "CrazySong.sqlite"), db));

            // Never falls back to the other kind when the paired file is missing.
            string lone = Path.Combine(root, "Lone", "CrazySongMgr.mdb");
            Directory.CreateDirectory(Path.GetDirectoryName(lone));
            File.Copy(mdb, lone);
            Assert.Equal(Path.ChangeExtension(lone, ".db"), Db.SongMgrDatabaseFor(Path.Combine(root, "CrazySong.db"), lone));
        }

        [Fact]
        public void SongNameSpell_IsTheSameFromSqliteAsFromAccess()
        {
            Global.PhoneticsWordList = new List<string>();
            List<string> fromAccess = CommonFunc.GetSongNameSpell("周杰倫 晴天 Love");

            Global.CrazyktvSongMgrDatabaseFile = db;
            Global.PhoneticsWordList = new List<string>();
            List<string> fromSqlite = CommonFunc.GetSongNameSpell("周杰倫 晴天 Love");

            Assert.Equal(fromAccess, fromSqlite);
            Assert.NotEqual("", fromSqlite[0]);
        }

        [Fact]
        public void CashboxDate_WrittenAsTextReadsBackAsDateTime()
        {
            using (DbConnection conn = Db.Open(db))
            {
                using (DbCommand cmd = Db.Command("insert into ktv_Cashbox ( Cashbox_Id, Song_Lang, Song_SongName, Song_Singer, Song_CreatDate ) values ( @CashboxId, @SongLang, @SongSongName, @SongSinger, @SongCreatDate )", conn))
                {
                    cmd.Parameters.AddWithValue("@CashboxId", "SQLITE1");
                    cmd.Parameters.AddWithValue("@SongLang", "國語");
                    cmd.Parameters.AddWithValue("@SongSongName", "晴天");
                    cmd.Parameters.AddWithValue("@SongSinger", "周杰倫");
                    cmd.Parameters.AddWithValue("@SongCreatDate", Db.DateParam(conn, "2024/01/02"));
                    cmd.ExecuteNonQuery();
                }
                using (DbCommand cmd = Db.Command("update ktv_Version set CashboxUpdDate = @CashboxUpdDate where Id = @Id", conn))
                {
                    cmd.Parameters.AddWithValue("@CashboxUpdDate", Db.DateParam(conn, "2026/11/01"));
                    cmd.Parameters.AddWithValue("@Id", "1");
                    cmd.ExecuteNonQuery();
                }
            }

            DataRow song = Db.GetDataTable(db, "select * from ktv_Cashbox where Cashbox_Id = 'SQLITE1'").Rows.Cast<DataRow>().Single();
            Assert.Equal(new DateTime(2024, 1, 2), song.Field<DateTime>("Song_CreatDate"));
            Assert.Equal(new DateTime(2026, 11, 1), DateTime.Parse(Db.GetDataTable(db, "select * from ktv_Version").Rows[0]["CashboxUpdDate"].ToString()));
        }

        [Fact]
        public void DateParam_LeavesAccessTextAlone()
        {
            using (DbConnection conn = Db.Open(mdb))
            {
                Assert.Equal("2024/01/02", Db.DateParam(conn, "2024/01/02"));
            }
        }

        [Fact]
        public void CashboxDateQuery_MatchesIsoDatesOnSqlite()
        {
            Global.CrazyktvSongMgrDatabaseFile = db;
            using (DbConnection conn = Db.Open(db))
            {
                using (DbCommand cmd = Db.Command("insert into ktv_Cashbox ( Cashbox_Id, Song_Lang, Song_SongName, Song_Singer, Song_CreatDate ) values ( 'SQLITE2', '國語', '晴天', '周杰倫', @SongCreatDate )", conn))
                {
                    cmd.Parameters.AddWithValue("@SongCreatDate", Db.DateParam(conn, "2024/1/2"));
                    cmd.ExecuteNonQuery();
                }
            }

            string sql = Cashbox.GetSongQuerySqlStr("SongDate", new DateTime(2024, 1, 2).ToShortDateString(), null);

            Assert.Contains("SQLITE2", Db.GetDataTable(db, sql).AsEnumerable().Select(r => r["Cashbox_Id"].ToString()));

            // the 更新日期 list of 錢櫃資料, newest first
            DataTable dates = Db.GetDataTable(db, "SELECT First(Song_CreatDate) AS Song_CreatDate, Count(Song_CreatDate) AS Song_CreatDateCount FROM ktv_Cashbox GROUP BY Song_CreatDate HAVING (Count(Song_CreatDate)>0) ORDER BY First(Song_CreatDate) desc");
            List<DateTime> list = dates.AsEnumerable().Select(r => r.Field<DateTime>("Song_CreatDate")).ToList();
            Assert.Contains(new DateTime(2024, 1, 2), list);
            Assert.Equal(list.OrderByDescending(d => d), list);
        }
    }
}

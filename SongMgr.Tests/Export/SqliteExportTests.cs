using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Export
{
    /// <summary>
    /// 匯出 SQLite 資料庫 on scratch copies of the bundled CrazySongEmpty.mdb / CrazySongMgr.mdb: both .db files,
    /// their version, the Access column order and primary key, whole-number Song_MB, and a failed export
    /// leaving the previous .db alone.
    /// </summary>
    public class SqliteExportTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("export");
        readonly string song;
        readonly string mgr;

        public SqliteExportTests()
        {
            song = Path.Combine(root, "CrazySong.mdb");
            mgr = Path.Combine(root, "CrazySongMgr.mdb");
            File.Copy(Path.Combine(TestEnv.FixtureDir, "CrazySongEmpty.mdb"), song);
            File.Copy(Path.Combine(TestEnv.FixtureDir, "CrazySongMgr.mdb"), mgr);
            File.SetAttributes(song, FileAttributes.Normal);
            File.SetAttributes(mgr, FileAttributes.Normal);
        }

        public void Dispose()
        {
            SQLiteConnection.ClearAllPools();
            TestEnv.DeleteDir(root);
        }

        [Fact]
        public void Jobs_WriteDbFilesBesideCrazySong()
        {
            List<SqliteExport.Job> jobs = SqliteExport.Jobs(song, Path.Combine(root, "SongMgr", "CrazySongMgr.mdb"));

            Assert.Equal(Path.Combine(root, "CrazySong.db"), jobs[0].Dest);
            Assert.Equal(Path.Combine(root, "CrazySongMgr.db"), jobs[1].Dest);
            Assert.Equal(SqliteExport.CrazySongDBVer, jobs[0].DBVer);
            Assert.Equal(SqliteExport.CrazySongMgrDBVer, jobs[1].DBVer);
        }

        [Fact]
        public void Run_ExportsBothDatabasesWithVersionAndRowCounts()
        {
            InsertSong("000001", 136.7f);

            List<SqliteExport.Result> results = SqliteExport.Run(SqliteExport.Jobs(song, mgr));

            Assert.All(results, r => Assert.True(r.Succeeded, r.Error?.ToString()));
            foreach (SqliteExport.Result result in results)
            {
                Assert.True(File.Exists(result.Job.Dest));
                Assert.False(File.Exists(result.Job.Dest + ".tmp"));
                using (SQLiteConnection conn = OpenDb(result.Job.Dest))
                {
                    Assert.Equal((long)result.Job.DBVer, Scalar(conn, "select DBVer from ktv_DBVersion"));
                    Assert.Equal(1L, Scalar(conn, "select count(*) from ktv_DBVersion"));
                    Assert.Equal((long)result.Job.DBVer, Scalar(conn, "pragma user_version"));

                    foreach (var table in result.Tables)
                    {
                        Assert.Equal((long)AccessRowCount(result.Job.Source, table.Table), Scalar(conn, "select count(*) from \"" + table.Table + "\""));
                    }
                }
                Assert.Equal(AccessTables(result.Job.Source).OrderBy(t => t), result.Tables.Select(t => t.Table).OrderBy(t => t));
            }
            Assert.Contains(results[0].Tables, t => t.Table == "ktv_Song" && t.Rows == 1);
            Assert.Contains(results[1].Tables, t => t.Table == "ktv_Cashbox");
        }

        [Fact]
        public void Run_KeepsAccessColumnOrderAndPrimaryKey()
        {
            SqliteExport.Run(SqliteExport.Jobs(song, mgr));

            using (SQLiteConnection conn = OpenDb(Path.Combine(root, "CrazySong.db")))
            {
                var columns = TableInfo(conn, "ktv_Song");
                Assert.Equal(AccessColumns(song, "ktv_Song"), columns.Select(c => c.Name));
                Assert.Equal(new[] { "Song_Id" }, columns.Where(c => c.Pk > 0).Select(c => c.Name));
                Assert.Equal("INTEGER", columns.Single(c => c.Name == "Song_MB").Type);
            }
        }

        [Fact]
        public void Run_RoundsSongMBToWholeNumber()
        {
            InsertSong("000001", 136.7f);
            InsertSong("000002", 141.5f);
            InsertSong("000003", 0.4f);

            SqliteExport.Run(SqliteExport.Jobs(song, mgr));

            using (SQLiteConnection conn = OpenDb(Path.Combine(root, "CrazySong.db")))
            {
                Assert.Equal(137L, Scalar(conn, "select Song_MB from ktv_Song where Song_Id = '000001'"));
                Assert.Equal(142L, Scalar(conn, "select Song_MB from ktv_Song where Song_Id = '000002'"));
                Assert.Equal(0L, Scalar(conn, "select Song_MB from ktv_Song where Song_Id = '000003'"));
                Assert.Equal("integer", Scalar(conn, "select typeof(Song_MB) from ktv_Song where Song_Id = '000001'"));
            }
        }

        [Fact]
        public void Run_ExportsYoutubeSongList()
        {
            YoutubeSongStore.CreateTable(song);
            Execute(song, "insert into YoutubeSongList (Id, Name, Url, PlayCount, Length, FileSize, Volume, WordCount) values ('Y0001', '快樂崇拜', 'abcdefghijk', 1, 221, 90721283, 30, 0)");

            SqliteExport.Run(SqliteExport.Jobs(song, mgr));

            using (SQLiteConnection conn = OpenDb(Path.Combine(root, "CrazySong.db")))
            {
                Assert.Equal("abcdefghijk", Scalar(conn, "select Url from YoutubeSongList where Id = 'Y0001'"));
                Assert.Equal(3L, Scalar(conn, "select Track from YoutubeSongList where Id = 'Y0001'"));
                Assert.Equal(new[] { "Id" }, TableInfo(conn, "YoutubeSongList").Where(c => c.Pk > 0).Select(c => c.Name));
            }
        }

        [Fact]
        public void Run_FailureLeavesExistingDbAndReportsIt()
        {
            string dest = Path.Combine(root, "CrazySong.db");
            File.WriteAllText(dest, "previous export");

            List<SqliteExport.Result> results;
            using (new FileStream(dest, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                results = SqliteExport.Run(SqliteExport.Jobs(song, mgr));
            }

            Assert.False(results[0].Succeeded);
            Assert.True(results[1].Succeeded, results[1].Error?.ToString());
            Assert.Equal("previous export", File.ReadAllText(dest));
            Assert.False(File.Exists(dest + ".tmp"));

            string message = SqliteExport.Describe(results);
            Assert.Contains("匯出失敗: " + dest, message);
            Assert.Contains("匯出完成: " + Path.Combine(root, "CrazySongMgr.db"), message);
        }

        [Fact]
        public void OverwritePrompt_WarnsOnlyWhenCrazySongDbIsReplaced()
        {
            List<SqliteExport.Job> jobs = SqliteExport.Jobs(song, mgr);
            Assert.Equal("", SqliteExport.OverwritePrompt(jobs, song));

            File.WriteAllText(jobs[1].Dest, "");
            string mgrOnly = SqliteExport.OverwritePrompt(jobs, song);
            Assert.Contains(jobs[1].Dest, mgrOnly);
            Assert.DoesNotContain("播放器", mgrOnly);

            File.WriteAllText(jobs[0].Dest, "");
            string both = SqliteExport.OverwritePrompt(jobs, song);
            Assert.Contains(jobs[0].Dest, both);
            Assert.Contains("播放器寫進 CrazySong.db 的點播次數、YouTube 紀錄、我的最愛,會被 CrazySong.mdb 的內容取代", both);
        }

        // ---------- helpers ----------

        void InsertSong(string id, float mb)
        {
            using (OleDbConnection conn = new OleDbConnection(DbHelper.Access.ConnectionString(song)))
            {
                conn.Open();
                using (OleDbCommand cmd = new OleDbCommand("insert into ktv_Song (Song_Id, Song_Lang, Song_Singer, Song_SongName, Song_MB, Song_CreatDate, Song_FileName, Song_Path) values (@Id, '國語', '周杰倫', '晴天', @MB, @Date, '晴天.mpg', 'D:\\KTV\\')", conn))
                {
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.Parameters.AddWithValue("@MB", mb);
                    cmd.Parameters.AddWithValue("@Date", new DateTime(2026, 1, 2, 3, 4, 5).ToString("yyyy/MM/dd HH:mm:ss"));
                    cmd.ExecuteNonQuery();
                }
            }
        }

        static void Execute(string mdb, string sql)
        {
            using (OleDbConnection conn = new OleDbConnection(DbHelper.Access.ConnectionString(mdb)))
            {
                conn.Open();
                using (OleDbCommand cmd = new OleDbCommand(sql, conn)) cmd.ExecuteNonQuery();
            }
        }

        static List<string> AccessTables(string mdb)
        {
            using (OleDbConnection conn = new OleDbConnection(DbHelper.Access.ConnectionString(mdb)))
            {
                conn.Open();
                return conn.GetSchema("Tables", new string[] { null, null, null, "TABLE" }).AsEnumerable().Select(r => r["TABLE_NAME"].ToString()).ToList();
            }
        }

        static List<string> AccessColumns(string mdb, string table)
        {
            using (OleDbConnection conn = new OleDbConnection(DbHelper.Access.ConnectionString(mdb)))
            {
                conn.Open();
                return conn.GetSchema("Columns", new string[] { null, null, table, null }).AsEnumerable()
                    .OrderBy(r => Convert.ToInt64(r["ORDINAL_POSITION"])).Select(r => r["COLUMN_NAME"].ToString()).ToList();
            }
        }

        static int AccessRowCount(string mdb, string table)
        {
            using (OleDbConnection conn = new OleDbConnection(DbHelper.Access.ConnectionString(mdb)))
            {
                conn.Open();
                using (OleDbCommand cmd = new OleDbCommand("select count(*) from [" + table + "]", conn)) return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        static SQLiteConnection OpenDb(string path)
        {
            SQLiteConnection conn = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            conn.Open();
            return conn;
        }

        static object Scalar(SQLiteConnection conn, string sql)
        {
            using (SQLiteCommand cmd = new SQLiteCommand(sql, conn)) return cmd.ExecuteScalar();
        }

        static List<(string Name, string Type, long Pk)> TableInfo(SQLiteConnection conn, string table)
        {
            var columns = new List<(string, string, long)>();
            using (SQLiteCommand cmd = new SQLiteCommand("pragma table_info(\"" + table + "\")", conn))
            using (SQLiteDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read()) columns.Add((reader.GetString(1), reader.GetString(2), reader.GetInt64(5)));
            }
            return columns;
        }
    }
}

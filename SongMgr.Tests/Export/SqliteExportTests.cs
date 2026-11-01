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
    /// their version, the Access column order and primary key, whole-number Song_MB, an existing .db left alone
    /// when its content is the same or the user keeps it, and a failed export leaving the previous .db alone.
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
        public void Jobs_WriteEachDbBesideItsOwnMdb()
        {
            List<SqliteExport.Job> jobs = SqliteExport.Jobs(song, Path.Combine(root, "SongMgr", "CrazySongMgr.mdb"));

            Assert.Equal(Path.Combine(root, "CrazySong.db"), jobs[0].Dest);
            Assert.Equal(Path.Combine(root, "SongMgr", "CrazySongMgr.db"), jobs[1].Dest);
            Assert.Equal(SqliteExport.CrazySongDBVer, jobs[0].DBVer);
            Assert.Equal(SqliteExport.CrazySongMgrDBVer, jobs[1].DBVer);
            Assert.False(jobs[0].IsSongMgr);
            Assert.True(jobs[1].IsSongMgr);
        }

        [Fact]
        public void Build_CreatesMissingDbFiles()
        {
            List<SqliteExport.Result> results = SqliteExport.Build(SqliteExport.Jobs(song, mgr));

            Assert.All(results, r => Assert.Equal(SqliteExport.Outcome.Created, r.Outcome));
            Assert.All(results, r => Assert.True(File.Exists(r.Job.Dest)));
            Assert.All(results, r => Assert.False(File.Exists(r.Job.TempFile)));

            string message = SqliteExport.Describe(results);
            Assert.Contains("已建立: " + Path.Combine(root, "CrazySong.db") + " (資料庫版本 1)", message);
            Assert.Contains("SongMgr 下次啟動時會改用此檔案", message);
        }

        [Fact]
        public void Build_SameContentLeavesExistingDbUntouched()
        {
            InsertSong("000001", 136.7f);
            SqliteExport.Run(SqliteExport.Jobs(song, mgr));
            string dest = Path.Combine(root, "CrazySong.db");
            byte[] before = File.ReadAllBytes(dest);
            DateTime written = File.GetLastWriteTimeUtc(dest);

            List<SqliteExport.Result> results = SqliteExport.Build(SqliteExport.Jobs(song, mgr));

            Assert.All(results, r => Assert.Equal(SqliteExport.Outcome.Unchanged, r.Outcome));
            Assert.All(results, r => Assert.False(File.Exists(r.Job.TempFile)));
            Assert.Equal(before, File.ReadAllBytes(dest));
            Assert.Equal(written, File.GetLastWriteTimeUtc(dest));
            Assert.Contains("內容相同,未變更: " + dest, SqliteExport.Describe(results));
        }

        [Fact]
        public void Build_DifferentContentWaitsAndKeepLeavesExistingDb()
        {
            SqliteExport.Run(SqliteExport.Jobs(song, mgr));
            InsertSong("000001", 136.7f);

            List<SqliteExport.Result> results = SqliteExport.Build(SqliteExport.Jobs(song, mgr));
            SqliteExport.Result songResult = results[0];

            Assert.Equal(SqliteExport.Outcome.Different, songResult.Outcome);
            Assert.Equal(new[] { "ktv_Song: 資料不同 (現有 0 筆 → 匯出 1 筆)" }, songResult.Differences);
            Assert.True(File.Exists(songResult.Job.TempFile));
            Assert.Equal(SqliteExport.Outcome.Unchanged, results[1].Outcome);

            SqliteExport.Keep(songResult);

            Assert.Equal(SqliteExport.Outcome.Kept, songResult.Outcome);
            Assert.False(File.Exists(songResult.Job.TempFile));
            using (SQLiteConnection conn = OpenDb(songResult.Job.Dest)) Assert.Equal(0L, Scalar(conn, "select count(*) from ktv_Song"));
            Assert.Contains("未覆蓋,保留現有檔案: " + songResult.Job.Dest, SqliteExport.Describe(results));
        }

        [Fact]
        public void Build_DifferentContentReplaceOverwritesExistingDb()
        {
            SqliteExport.Run(SqliteExport.Jobs(song, mgr));
            InsertSong("000001", 136.7f);

            SqliteExport.Result songResult = SqliteExport.Build(SqliteExport.Jobs(song, mgr))[0];
            SqliteExport.Replace(songResult);

            Assert.Equal(SqliteExport.Outcome.Replaced, songResult.Outcome);
            Assert.False(File.Exists(songResult.Job.TempFile));
            using (SQLiteConnection conn = OpenDb(songResult.Job.Dest)) Assert.Equal(1L, Scalar(conn, "select count(*) from ktv_Song"));
            Assert.Contains("已覆蓋: " + songResult.Job.Dest, SqliteExport.Describe(new List<SqliteExport.Result> { songResult }));
        }

        [Fact]
        public void Compare_FindsWhatThePlayerChangedInTheDb()
        {
            InsertSong("000001", 136.7f);
            SqliteExport.Run(SqliteExport.Jobs(song, mgr));
            string dest = Path.Combine(root, "CrazySong.db");
            using (SQLiteConnection conn = OpenDb(dest))
            {
                foreach (string sql in new[] {
                    "update ktv_Song set Song_PlayCount = 5 where Song_Id = '000001'",
                    "create table PlayerOnly (Id TEXT)",
                    "pragma user_version = 0" })
                {
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery();
                }
            }

            SqliteExport.Result songResult = SqliteExport.Build(SqliteExport.Jobs(song, mgr))[0];

            Assert.Equal(new[] { "資料庫版本: 0 → 1", "ktv_Song: 資料不同 (現有 1 筆 → 匯出 1 筆)", "PlayerOnly: 歌庫沒有此資料表 (覆蓋後會移除)" }, songResult.Differences);
            SqliteExport.Keep(songResult);
        }

        [Fact]
        public void Compare_ReportsAnUnreadableExistingFile()
        {
            string dest = Path.Combine(root, "CrazySong.db");
            File.WriteAllText(dest, "not a database");

            SqliteExport.Result songResult = SqliteExport.Build(SqliteExport.Jobs(song, mgr))[0];

            Assert.Equal(SqliteExport.Outcome.Different, songResult.Outcome);
            Assert.StartsWith("無法讀取現有檔案: ", Assert.Single(songResult.Differences));
            SqliteExport.Keep(songResult);
            Assert.Equal("not a database", File.ReadAllText(dest));
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
            Assert.Contains("已建立: " + Path.Combine(root, "CrazySongMgr.db"), message);
        }

        [Fact]
        public void OverwritePrompt_ListsDifferencesAndWarnsOnlyForCrazySongDb()
        {
            SqliteExport.Run(SqliteExport.Jobs(song, mgr));
            InsertSong("000001", 136.7f);
            Execute(mgr, "delete from ktv_Version");

            List<SqliteExport.Result> results = SqliteExport.Build(SqliteExport.Jobs(song, mgr));
            string songPrompt = SqliteExport.OverwritePrompt(results[0]);
            string mgrPrompt = SqliteExport.OverwritePrompt(results[1]);
            results.ForEach(SqliteExport.Keep);

            Assert.Contains(results[0].Job.Dest, songPrompt);
            Assert.Contains("ktv_Song: 資料不同 (現有 0 筆 → 匯出 1 筆)", songPrompt);
            Assert.Contains("播放器寫進 CrazySong.db 的點播次數、YouTube 紀錄、我的最愛,會被 CrazySong.mdb 的內容取代", songPrompt);
            Assert.EndsWith("是否要覆蓋?", songPrompt);

            Assert.Contains("ktv_Version: 資料不同", mgrPrompt);
            Assert.DoesNotContain("播放器", mgrPrompt);
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

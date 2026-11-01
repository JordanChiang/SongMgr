using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Store
{
    /// <summary>
    /// YouTube 歌曲 on a scratch copy of CrazySongEmpty.mdb: creating YoutubeSongList the way UltimateKtv does,
    /// listing and searching it, editing each editable column, refusing bad values, and deleting rows.
    /// </summary>
    public class YoutubeSongStoreTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("youtube");
        readonly string db;

        public YoutubeSongStoreTests()
        {
            db = Path.Combine(root, "CrazySong.mdb");
            File.Copy(Path.Combine(TestEnv.FixtureDir, "CrazySongEmpty.mdb"), db);
            File.SetAttributes(db, FileAttributes.Normal);
        }

        public void Dispose() => TestEnv.DeleteDir(root);

        [Fact]
        public void CreateTable_MatchesUltimateKtvDefinition()
        {
            Assert.False(YoutubeSongStore.TableExists(db));

            YoutubeSongStore.CreateTable(db);

            Assert.True(YoutubeSongStore.TableExists(db));
            using (OleDbConnection conn = Open())
            {
                var columns = conn.GetSchema("Columns", new string[] { null, null, YoutubeSongStore.TableName, null }).AsEnumerable()
                    .OrderBy(r => Convert.ToInt64(r["ORDINAL_POSITION"])).ToList();
                Assert.Equal(new[] { "Id", "Name", "CreatDate", "Url", "Singer", "Lang", "PlayCount", "Length", "FileName", "FileSize", "Track", "Volume", "WordCount" },
                    columns.Select(r => r["COLUMN_NAME"].ToString()));
                Assert.Equal(5L, Convert.ToInt64(columns.Single(r => r["COLUMN_NAME"].ToString() == "Id")["CHARACTER_MAXIMUM_LENGTH"]));
                Assert.Equal(255L, Convert.ToInt64(columns.Single(r => r["COLUMN_NAME"].ToString() == "Name")["CHARACTER_MAXIMUM_LENGTH"]));
                Assert.Equal(50L, Convert.ToInt64(columns.Single(r => r["COLUMN_NAME"].ToString() == "Singer")["CHARACTER_MAXIMUM_LENGTH"]));
                Assert.Equal(16L, Convert.ToInt64(columns.Single(r => r["COLUMN_NAME"].ToString() == "Lang")["CHARACTER_MAXIMUM_LENGTH"]));
            }

            Execute("insert into YoutubeSongList (Id, Name) values ('Y0001', 'a')");
            Assert.Equal(3, Convert.ToInt32(Cell("Y0001", "Track")));
        }

        [Fact]
        public void Query_ListsEveryRowOrderedById()
        {
            Seed();

            DataTable dt = YoutubeSongStore.Query(db, "");

            Assert.Equal(new[] { "Y0001", "Y0002", "Y0003" }, dt.AsEnumerable().Select(r => r["Id"].ToString()));
            Assert.Equal(new[] { "Id", "Name", "Url", "Length", "PlayCount", "Track", "Volume", "FileName", "FileSize", "CreatDate" },
                dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        }

        [Theory]
        [InlineData("sistar", new[] { "Y0002", "Y0003" })]
        [InlineData("  [空耳]  ", new[] { "Y0003" })]
        [InlineData("快樂", new[] { "Y0001" })]
        [InlineData("沒有這首", new string[0])]
        public void Query_FiltersByNameIgnoringCase(string keyword, string[] expected)
        {
            Seed();

            DataTable dt = YoutubeSongStore.Query(db, keyword);

            Assert.Equal(expected, dt.AsEnumerable().Select(r => r["Id"].ToString()));
        }

        [Theory]
        [InlineData("Name", " 快樂崇拜 ", "快樂崇拜")]
        [InlineData("Track", 0, "0")]
        [InlineData("Track", "3", "3")]
        [InlineData("Volume", "100", "100")]
        [InlineData("PlayCount", 0, "0")]
        public void Update_WritesOneColumnOfOneRow(string columnName, object input, string stored)
        {
            var column = Enum.Parse<YoutubeColumn>(columnName);
            Seed();

            Assert.Equal(1, YoutubeSongStore.Update(db, "Y0001", column, input));

            Assert.Equal(stored, Cell("Y0001", column.ToString()).ToString());
            Assert.Equal("Sistar - Give It To Me", Cell("Y0002", "Name"));
        }

        [Theory]
        [InlineData("Name", "")]
        [InlineData("Volume", "101")]
        [InlineData("Volume", "abc")]
        [InlineData("PlayCount", "-1")]
        [InlineData("Track", "6")]
        [InlineData("Track", "1")]
        [InlineData("Track", "2")]
        public void Update_RefusesValuesOutsideTheColumnRules(string columnName, string input)
        {
            var column = Enum.Parse<YoutubeColumn>(columnName);
            Seed();

            Assert.NotEqual("", YoutubeSongStore.Validate(column, input, out _));
            Assert.Throws<ArgumentException>(() => YoutubeSongStore.Update(db, "Y0001", column, input));
        }

        [Fact]
        public void Delete_RemovesOnlyTheGivenRows()
        {
            Seed();

            Assert.Equal(2, YoutubeSongStore.Delete(db, new[] { "Y0001", "Y0003" }));

            Assert.Equal(new[] { "Y0002" }, YoutubeSongStore.Query(db, "").AsEnumerable().Select(r => r["Id"].ToString()));
        }

        [Fact]
        public void Formatting_ShowsMinutesSecondsWholeMBAndWatchUrl()
        {
            Assert.Equal("3:41", YoutubeSongStore.FormatLength(221));
            Assert.Equal("48:03", YoutubeSongStore.FormatLength(2883));
            Assert.Equal("0:00", YoutubeSongStore.FormatLength(0));
            Assert.Equal("", YoutubeSongStore.FormatLength(DBNull.Value));
            Assert.Equal("87", YoutubeSongStore.FormatMB(90721283));
            Assert.Equal("", YoutubeSongStore.FormatMB(DBNull.Value));
            Assert.Equal("https://www.youtube.com/watch?v=1bQ5t-VjM60", YoutubeSongStore.WatchUrl("1bQ5t-VjM60"));
            Assert.Equal("", YoutubeSongStore.WatchUrl(DBNull.Value));
        }

        // ---------- helpers ----------

        /// <summary>Rows shaped like UltimateKtv's RecordYoutubeSong writes them.</summary>
        void Seed()
        {
            YoutubeSongStore.CreateTable(db);
            Insert("Y0001", "潘瑋柏 - 快樂崇拜 ft. 張韶涵", "aaaaaaaaaaa", "03:42", 221);
            Insert("Y0002", "Sistar - Give It To Me", "bbbbbbbbbbb", "", 201);
            Insert("Y0003", "[空耳] Shake it  sistar", "ccccccccccc", "", 218);
        }

        void Insert(string id, string name, string url, string lang, int length)
        {
            using (OleDbConnection conn = Open())
            using (OleDbCommand cmd = new OleDbCommand("insert into YoutubeSongList (Id, Name, CreatDate, Url, Singer, Lang, PlayCount, Length, FileName, FileSize, Track, Volume, WordCount) values (@Id, @Name, @Date, @Url, '', @Lang, 1, @Length, @File, 90721283, 3, 30, 0)", conn))
            {
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@Name", name);
                cmd.Parameters.AddWithValue("@Date", "2026/08/08 21:42:48");
                cmd.Parameters.AddWithValue("@Url", url);
                cmd.Parameters.AddWithValue("@Lang", lang);
                cmd.Parameters.AddWithValue("@Length", length);
                cmd.Parameters.AddWithValue("@File", name + "_" + url + ".mp4");
                cmd.ExecuteNonQuery();
            }
        }

        object Cell(string id, string column)
        {
            using (OleDbConnection conn = Open())
            using (OleDbCommand cmd = new OleDbCommand("select [" + column + "] from YoutubeSongList where Id = @Id", conn))
            {
                cmd.Parameters.AddWithValue("@Id", id);
                return cmd.ExecuteScalar();
            }
        }

        void Execute(string sql)
        {
            using (OleDbConnection conn = Open())
            using (OleDbCommand cmd = new OleDbCommand(sql, conn)) cmd.ExecuteNonQuery();
        }

        OleDbConnection Open()
        {
            OleDbConnection conn = new OleDbConnection(SongLibraryStore.ConnectionString(db));
            conn.Open();
            return conn;
        }
    }
}

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
    /// YouTube 我的最愛 (UltimateKtv's YoutubeFavorite) on an Access and a SQLite 歌庫: adding and removing them through
    /// <see cref="SongLibraryStore"/>, listing them with the user's songs in a 我的最愛 歌庫查詢, removing a 最愛用戶,
    /// replacing a 錢櫃排行 list, and 我的最愛匯出 / 匯入.
    /// </summary>
    public class YoutubeFavoriteTests : IDisposable
    {
        const string VideoA = "dQw4w9WgXcQ";
        const string VideoB = "abcDEF12_-z";

        readonly string root = TestEnv.NewTempDir("ytfav");
        readonly string savedSong = Global.CrazyktvDatabaseFile;
        readonly string savedMgr = Global.CrazyktvSongMgrDatabaseFile;
        readonly string access;

        public YoutubeFavoriteTests()
        {
            TestEnv.UseScratchDatabases(root);
            access = Global.CrazyktvDatabaseFile;
            using (SongLibraryStore store = SongLibraryStore.OpenFile(access))
            {
                store.InsertSong(Song("000001", "晴天"), SongColumns.AllButReplayGain);
                store.InsertSong(Song("000002", "七里香"), SongColumns.AllButReplayGain);
                store.AddFavoriteUser("0001", "小明");
                store.AddFavoriteUser("0002", "小華");
                store.AddFavorites("0001", new[] { "000002" });
            }
        }

        public void Dispose()
        {
            System.Data.SQLite.SQLiteConnection.ClearAllPools();
            Global.CrazyktvDatabaseFile = savedSong;
            Global.CrazyktvSongMgrDatabaseFile = savedMgr;
            TestEnv.DeleteDir(root);
        }

        public static IEnumerable<object[]> Both => new[] { new object[] { "Access" }, new object[] { "SQLite" } };

        /// <summary>The scratch 歌庫, as .mdb or exported to .db.</summary>
        string Library(string kind)
        {
            if (kind == "Access") return access;
            string db = Path.Combine(root, "CrazySong.db");
            if (!File.Exists(db))
            {
                var results = SqliteExport.Run(new[] { new SqliteExport.Job(access, db, SqliteExport.CrazySongDBVer) });
                Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());
            }
            return db;
        }

        [Theory, MemberData(nameof(Both))]
        public void Add_creates_the_table_like_UltimateKtv_and_skips_repeats_and_bad_ids(string kind)
        {
            string db = Library(kind);
            Assert.False(YoutubeFavorites.TableExists(db));

            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                Assert.Equal(2, store.AddYoutubeFavorites("0001", new[]
                {
                    new YoutubeFavoriteItem(VideoA, "第一首", 212, new DateTime(2026, 1, 2, 3, 4, 5)),
                    new YoutubeFavoriteItem(VideoA, "重複", 0),
                    new YoutubeFavoriteItem("Y0017", "點唱紀錄編號", 0),
                    new YoutubeFavoriteItem("", "空白", 0),
                    new YoutubeFavoriteItem(VideoB, new string('歌', 300), -5),
                }));
                Assert.Equal(0, store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "已在清單", 0) }));
                Assert.Equal(1, store.AddYoutubeFavorites("0002", new[] { new YoutubeFavoriteItem(VideoA, "別人的", 0) }));
            }

            Assert.True(YoutubeFavorites.TableExists(db));
            Assert.Equal(new[] { "CreatDate", "Length", "Name", "User_Id", "VideoId" }, Db.GetColumnList(db, "", YoutubeFavorites.TableName).OrderBy(c => c, StringComparer.Ordinal));

            DataTable rows = Rows(db);
            Assert.Equal(3, rows.Rows.Count);
            DataRow a = rows.AsEnumerable().Single(r => r["User_Id"].ToString() == "0001" && r["VideoId"].ToString() == VideoA);
            Assert.Equal("第一首", a["Name"].ToString());
            Assert.Equal(212, Convert.ToInt32(a["Length"]));
            Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), YoutubeFavorites.ReadDate(a["CreatDate"]));
            DataRow b = rows.AsEnumerable().Single(r => r["VideoId"].ToString() == VideoB);
            Assert.Equal(255, b["Name"].ToString().Length);
            Assert.Equal(0, Convert.ToInt32(b["Length"]));
        }

        [Theory, MemberData(nameof(Both))]
        public void Favorite_search_lists_the_users_songs_then_their_youtube_favorites(string kind)
        {
            string db = Library(kind);
            YoutubeSongStore.CreateTable(db);
            Execute(db, "insert into YoutubeSongList (Id, Name, Url, PlayCount, Volume, FileSize) values ('Y0001', '下載過', '" + VideoB + "', 7, 80, 10485760)");
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                store.AddYoutubeFavorites("0001", new[]
                {
                    new YoutubeFavoriteItem(VideoB, "後加入", 100, new DateTime(2026, 2, 1)),
                    new YoutubeFavoriteItem(VideoA, "先加入", 200, new DateTime(2026, 1, 1)),
                });
                store.AddYoutubeFavorites("0002", new[] { new YoutubeFavoriteItem(VideoA, "別人的", 0) });
            }

            DataTable songs = SongSearch.Run(new SongSearchCriteria(SongSearchKind.Favorite, "0001"), new SongSearchReference(), db).Songs;

            Assert.Equal(new[] { "000002", VideoA, VideoB }, songs.AsEnumerable().Select(r => r["Song_Id"].ToString()));
            Assert.False(YoutubeFavorites.IsFavoriteRow(songs.Rows[0]["Song_Lang"], songs.Rows[0]["Song_Path"]));

            DataRow first = songs.Rows[1];
            Assert.True(YoutubeFavorites.IsFavoriteRow(first["Song_Lang"], first["Song_Path"]));
            Assert.Equal("先加入", first["Song_SongName"].ToString());
            Assert.Equal("https://www.youtube.com/watch?v=" + VideoA, first["Song_FileName"].ToString());
            Assert.Equal(new DateTime(2026, 1, 1), DateTime.Parse(first["Song_CreatDate"].ToString()));
            Assert.Equal(0, Convert.ToInt32(first["Song_PlayCount"]));

            DataRow downloaded = songs.Rows[2];
            Assert.Equal(7, Convert.ToInt32(downloaded["Song_PlayCount"]));
            Assert.Equal(80, Convert.ToInt32(downloaded["Song_Volume"]));
            Assert.Equal(10.0, Convert.ToDouble(downloaded["Song_MB"]));

            Assert.Equal(new[] { VideoA }, SongSearch.Run(new SongSearchCriteria(SongSearchKind.Favorite, "0002"), new SongSearchReference(), db).Songs.AsEnumerable().Select(r => r["Song_Id"].ToString()));
        }

        [Theory, MemberData(nameof(Both))]
        public void Favorite_search_without_the_table_lists_only_songs(string kind)
        {
            string db = Library(kind);
            DataTable songs = SongSearch.Run(new SongSearchCriteria(SongSearchKind.Favorite, "0001"), new SongSearchReference(), db).Songs;
            Assert.Equal(new[] { "000002" }, songs.AsEnumerable().Select(r => r["Song_Id"].ToString()));
        }

        [Theory, MemberData(nameof(Both))]
        public void Remove_takes_videos_off_one_users_list(string kind)
        {
            string db = Library(kind);
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                Assert.Equal(0, store.RemoveYoutubeFavorites("0001", new[] { VideoA }));
                store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "a", 0), new YoutubeFavoriteItem(VideoB, "b", 0) });
                store.AddYoutubeFavorites("0002", new[] { new YoutubeFavoriteItem(VideoA, "a", 0) });

                Assert.Equal(1, store.RemoveYoutubeFavorites("0001", new[] { VideoA }));
            }
            Assert.Equal(new[] { "0001|" + VideoB, "0002|" + VideoA }, Keys(db));
        }

        [Theory, MemberData(nameof(Both))]
        public void Removing_a_user_removes_their_youtube_favorites_but_replacing_a_list_keeps_them(string kind)
        {
            string db = Library(kind);
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "a", 0) });
                store.AddYoutubeFavorites("0002", new[] { new YoutubeFavoriteItem(VideoB, "b", 0) });

                Assert.Equal(1, store.RemoveFavoriteUser("0001"));
                Assert.Equal(1, store.ReplaceFavoriteUser("0002", "小華", new[] { "000001" }));
            }
            Assert.Equal(new[] { "0002|" + VideoB }, Keys(db));
        }

        [Fact]
        public void Sqlite_export_carries_the_table_and_its_two_column_key()
        {
            using (SongLibraryStore store = SongLibraryStore.OpenFile(access))
            {
                store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "a", 212, new DateTime(2026, 1, 2, 3, 4, 5)) });
                // UltimateKtv's 點唱紀錄 users: 今日點唱 and one 開機 session.
                store.AddYoutubeFavorites("****", new[] { new YoutubeFavoriteItem(VideoA, "a", 212) });
                store.AddYoutubeFavorites("11/23 21:17", new[] { new YoutubeFavoriteItem(VideoA, "a", 212) });
            }

            string db = Library("SQLite");

            Assert.Equal(new[] { "****|" + VideoA, "0001|" + VideoA, "11/23 21:17|" + VideoA }, Keys(db));
            Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), YoutubeFavorites.ReadDate(Rows(db).AsEnumerable().Single(r => r["User_Id"].ToString() == "0001")["CreatDate"]));
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                Assert.Equal(1, store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoB, "b", 0) }));
            }
            Assert.ThrowsAny<DbException>(() => Execute(db, "insert into YoutubeFavorite (User_Id, VideoId) values ('0001', '" + VideoA + "')"));
        }

        [Fact]
        public void Removing_a_user_works_without_the_table()
        {
            using (SongLibraryStore store = SongLibraryStore.OpenFile(access))
            {
                Assert.Equal(1, store.RemoveFavoriteUser("0001"));
            }
            Assert.False(YoutubeFavorites.TableExists(access));
        }

        [Theory, MemberData(nameof(Both))]
        public void Import_of_a_backup_with_youtube_lines_replaces_the_youtube_favorites(string kind)
        {
            string db = Library(kind);
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                store.AddYoutubeFavorites("0002", new[] { new YoutubeFavoriteItem(VideoB, "舊的", 0) });
            }

            DataRow exported = ExportRow(VideoA, "歌名|有直線", 212, new DateTime(2026, 3, 4, 5, 6, 7));
            string line = CrazyKTV_SongMgr.SongMaintenance.YoutubeFavoriteLine("0001", exported);
            Assert.Equal("YoutubeFavorite|0001|" + VideoA + "|212|2026-03-04 05:06:07|歌名|有直線", line);

            string[] backup = { "ktv_User|0001|小明", "ktv_Favorite|0001|國語|周杰倫|晴天", line };
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                Assert.Equal((1, 2), CrazyKTV_SongMgr.SongMaintenance.ImportFavorites(store, backup, AllSongs(db)));
            }

            DataRow row = Rows(db).AsEnumerable().Single();
            Assert.Equal("0001", row["User_Id"].ToString());
            Assert.Equal(VideoA, row["VideoId"].ToString());
            Assert.Equal("歌名|有直線", row["Name"].ToString());
            Assert.Equal(212, Convert.ToInt32(row["Length"]));
            Assert.Equal(new DateTime(2026, 3, 4, 5, 6, 7), YoutubeFavorites.ReadDate(row["CreatDate"]));
        }

        [Theory, MemberData(nameof(Both))]
        public void Import_of_an_older_backup_keeps_the_youtube_favorites_of_the_users_it_brings_back(string kind)
        {
            string db = Library(kind);
            using (SongLibraryStore store = SongLibraryStore.OpenFile(db))
            {
                store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "a", 0) });
                store.AddYoutubeFavorites("0002", new[] { new YoutubeFavoriteItem(VideoB, "b", 0) });

                Assert.Equal((1, 1), CrazyKTV_SongMgr.SongMaintenance.ImportFavorites(store, new[] { "ktv_User|0001|小明", "ktv_Favorite|0001|國語|周杰倫|晴天" }, AllSongs(db)));
            }
            Assert.Equal(new[] { "0001|" + VideoA }, Keys(db));
        }

        [Fact]
        public void Import_rejects_a_cut_youtube_line_and_changes_nothing()
        {
            int users = Db.GetDataTable(access, "select User_Id from ktv_User", "").Rows.Count;
            using (SongLibraryStore store = SongLibraryStore.OpenFile(access))
            {
                store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "a", 0) });
                Assert.Throws<FormatException>(() => CrazyKTV_SongMgr.SongMaintenance.ImportFavorites(store, new[] { "ktv_User|0003|小英", "YoutubeFavorite|0003|" + VideoB }, AllSongs(access)));
            }
            Assert.Equal(new[] { "0001|" + VideoA }, Keys(access));
            Assert.Equal(users, Db.GetDataTable(access, "select User_Id from ktv_User", "").Rows.Count);
        }

        static SongRecord Song(string id, string name) => new SongRecord
        {
            Id = id, Lang = "國語", SingerType = "0", Singer = "周杰倫", SongName = name, Track = "1", SongType = "",
            Volume = "100", WordCount = name.Length.ToString(), PlayCount = "0", MB = "12.5", CreatDate = "2024/01/01 03:04:05",
            FileName = id + ".mpg", Path = @"D:\KTV\國語\", Spell = "", SpellNum = "", SongStroke = "0", PenStyle = "", PlayState = "0",
        };

        static DataTable AllSongs(string db) => SongSearch.Run(new SongSearchCriteria(SongSearchKind.AllSongs), new SongSearchReference(), db).Songs;

        static DataTable Rows(string db) => Db.GetDataTable(db, "select User_Id, VideoId, Name, Length, CreatDate from YoutubeFavorite", "");

        static List<string> Keys(string db) => Rows(db).AsEnumerable().Select(r => r["User_Id"] + "|" + r["VideoId"]).OrderBy(k => k, StringComparer.Ordinal).ToList();

        static DataRow ExportRow(string videoId, string name, int length, DateTime created)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("VideoId", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Length", typeof(int));
            dt.Columns.Add("CreatDate", typeof(DateTime));
            return dt.Rows.Add(videoId, name, length, created);
        }

        static void Execute(string db, string sql)
        {
            using (DbConnection conn = Db.Open(db))
            using (DbCommand cmd = Db.Command(sql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }
    }
}

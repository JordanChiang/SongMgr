using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.EndToEnd
{
    /// <summary>
    /// YouTube 我的最愛 in the real program: putting YouTube 歌曲 on a 最愛用戶's list from the 歌庫查詢 YouTube list,
    /// seeing them in that user's 我的最愛 歌庫查詢 after the songs, their 右鍵選單, 從我的最愛移除, and in 歌庫維護 ›
    /// 我的最愛 the export, import and removing the 最愛用戶. Checked in YoutubeFavorite, UltimateKtv's table.
    /// </summary>
    [Trait("Category", "EndToEnd")]
    public class YoutubeFavoriteEndToEndTests
    {
        const string VideoA = "dQw4w9WgXcQ";
        const string VideoB = "abcDEF12_-z";

        static void Arrange(AppSession app)
        {
            app.WriteSettings(6);
            app.InsertSong("100001", "國語", "周杰倫", "晴天");
            app.Execute("insert into ktv_User (User_Id, User_Name) values (@a, @b)", "0001", "小明");
            // A second user: the 最愛 drop-down starts on the first one, and picking the one shown runs no query.
            app.Execute("insert into ktv_User (User_Id, User_Name) values (@a, @b)", "0002", "小華");
            app.InsertFavorite("0001", "100001");
            YoutubeSongStore.CreateTable(app.Database);
            app.Execute("insert into YoutubeSongList (Id, Name, Url, Length, PlayCount, Volume, FileSize) values (@a, @b, @c, 221, 3, 80, 10485760)", "Y0001", "快樂崇拜", VideoA);
            app.Execute("insert into YoutubeSongList (Id, Name, Url, Length, PlayCount, Volume, FileSize) values (@a, @b, @c, 95, 0, 100, 0)", "Y0002", "小幸運", VideoB);
        }

        static List<string> YoutubeFavorites(AppSession app)
        {
            if (!CrazyKTV_SongMgr.YoutubeFavorites.TableExists(app.Database)) return new List<string>();
            using (DataTable dt = Db.GetDataTable(app.Database, "select User_Id, VideoId, Name, Length from YoutubeFavorite order by User_Id, VideoId", ""))
            {
                return dt.AsEnumerable().Select(r => r["User_Id"] + ":" + r["VideoId"] + ":" + r["Name"] + ":" + r["Length"]).ToList();
            }
        }

        [DesktopFact]
        public void 歌庫查詢_puts_YouTube_歌曲_on_我的最愛_lists_them_and_removes_them()
        {
            using (var app = new AppSession("ytfav-query"))
            {
                Arrange(app);
                app.Start();

                // YouTube 歌曲 list → 右鍵 加入我的最愛 → 小明: the table is created and the video written with its 名稱 and 長度.
                app.Choose("SongQuery_QueryType_ComboBox", "YouTube 歌曲");
                app.Click("查詢", onTab: "歌庫查詢");
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "總共查詢到 2 筆YouTube 歌曲。", 30, "the YouTube 歌曲 list");
                app.ClickRow("SongQuery_DataGridView", "Y0001");
                app.RightClickRow("SongQuery_DataGridView", "Y0001");
                app.ClickMenu("加入我的最愛", "小明");
                AppSession.WaitUntil(() => YoutubeFavorites(app).Count == 1, 20, "快樂崇拜 on 小明's 我的最愛");
                Assert.Equal(new[] { "0001:" + VideoA + ":快樂崇拜:221" }, YoutubeFavorites(app));
                Assert.Equal("已將 1 筆 YouTube 歌曲加入【小明】的我的最愛。", app.Value("SongQuery_QueryStatus_Label"));

                // The same video again is skipped.
                app.RightClickRow("SongQuery_DataGridView", "Y0001");
                app.ClickMenu("加入我的最愛", "小明");
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "已將 0 筆 YouTube 歌曲加入【小明】的我的最愛, 1 筆已在清單中。", 20, "the repeat to be skipped");

                app.ClickRow("SongQuery_DataGridView", "Y0002");
                app.RightClickRow("SongQuery_DataGridView", "Y0002");
                app.ClickMenu("加入我的最愛", "小明");
                AppSession.WaitUntil(() => YoutubeFavorites(app).Count == 2, 20, "小幸運 on 小明's 我的最愛");

                // 其它查詢 › 最愛 → 小明: the song, then the two videos.
                app.SelectSubTab("SongQuery_TabControl", "其它查詢", "歌曲查詢", "其它查詢");
                app.Focus("SongQuery_FavoriteQuery_ComboBox");
                app.Choose("SongQuery_FavoriteQuery_ComboBox", "小華");
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "查無最愛歌曲,請重新查詢...", 30, "小華's empty 我的最愛");
                app.Focus("SongQuery_FavoriteQuery_ComboBox");
                app.Choose("SongQuery_FavoriteQuery_ComboBox", "小明");
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "總共查詢到 3 筆屬於『小明』的最愛歌曲。", 30, "小明's 我的最愛");
                app.Snap("favorite-list");

                // 右鍵 on a video: its own menu. 複製影片網址 puts its address on the clipboard.
                app.ClickRow("SongQuery_DataGridView", VideoA);
                app.RightClickRow("SongQuery_DataGridView", VideoA);
                app.ClickMenu("複製影片網址");
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "已複製影片網址: https://www.youtube.com/watch?v=" + VideoA, 20, "the address copied");

                // 從我的最愛移除 takes only that video off; the song and the other video stay.
                app.RightClickRow("SongQuery_DataGridView", VideoA);
                app.ClickMenu("從我的最愛移除");
                app.Answer("刪除提示", "是(Y)");
                AppSession.WaitUntil(() => YoutubeFavorites(app).Count == 1, 20, "快樂崇拜 off 小明's 我的最愛");
                Assert.Equal(new[] { "0001:" + VideoB + ":小幸運:95" }, YoutubeFavorites(app));
                Assert.Equal(new[] { "0001:100001" }, app.Favorites());
            }
        }

        [DesktopFact]
        public void 歌庫維護_exports_imports_and_removes_YouTube_我的最愛_with_the_user()
        {
            using (var app = new AppSession("ytfav-maint"))
            {
                Arrange(app);
                using (SongLibraryStore store = SongLibraryStore.OpenFile(app.Database))
                {
                    store.AddYoutubeFavorites("0001", new[] { new YoutubeFavoriteItem(VideoA, "快樂崇拜", 221, new DateTime(2026, 11, 1, 20, 30, 0)) });
                }
                app.Start();

                // 匯出我的最愛 writes the video as a YoutubeFavorite line.
                app.GoToTab("歌庫維護");
                app.SelectSubTab("SongMaintenance_TabControl", "我的最愛", "維護雜項", "我的最愛", "自訂語系", "FFmpeg", "資料庫結構");
                app.Click("匯出我的最愛", onTab: "歌庫維護");
                string backup = Path.Combine(app.AppDir, "SongMgr", "Backup", "Favorite.txt");
                AppSession.WaitUntil(() => File.Exists(backup) && File.ReadAllLines(backup).Any(l => l.StartsWith("YoutubeFavorite|")), 30, "Favorite.txt");
                Assert.Contains("YoutubeFavorite|0001|" + VideoA + "|221|2026-11-01 20:30:00|快樂崇拜", File.ReadAllLines(backup));
                Assert.Contains("ktv_Favorite|0001|國語|周杰倫|晴天", File.ReadAllLines(backup));

                // 匯入我的最愛 after the video was taken off brings it back, with its date.
                app.Execute("delete from YoutubeFavorite");
                app.Click("匯入我的最愛", onTab: "歌庫維護");
                app.Answer("確認提示", "是(Y)");
                AppSession.WaitUntil(() => YoutubeFavorites(app).Count == 1, 60, "the import");
                Assert.Equal(new[] { "0001:" + VideoA + ":快樂崇拜:221" }, YoutubeFavorites(app));
                Assert.Equal(new[] { "0001:100001" }, app.Favorites());
                using (DataTable dt = Db.GetDataTable(app.Database, "select CreatDate from YoutubeFavorite", ""))
                {
                    Assert.Equal(new DateTime(2026, 11, 1, 20, 30, 0), CrazyKTV_SongMgr.YoutubeFavorites.ReadDate(dt.Rows[0][0]));
                }

                // 移除 小明 removes his YouTube 我的最愛 too.
                app.SelectListItem("SongMaintenance_Favorite_ListBox", "小明");
                app.Click("移除", onTab: "歌庫維護");
                app.Answer("確認提示", "是(Y)");
                AppSession.WaitUntil(() => YoutubeFavorites(app).Count == 0, 20, "小明 and his YouTube 我的最愛 removed");
                Assert.DoesNotContain(app.Favorites(), f => f.StartsWith("0001:"));
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.EndToEnd
{
    /// <summary>A desktop test that also needs a real yt-dlp: the one SONGMGR_YTDLP names. It asks YouTube, so it needs the internet.</summary>
    public sealed class YtDlpDesktopFactAttribute : FactAttribute
    {
        public static string YtDlp => Environment.GetEnvironmentVariable("SONGMGR_YTDLP") is string s && File.Exists(s) ? s : null;

        public YtDlpDesktopFactAttribute()
        {
            if (AppSession.SkipReason != null) Skip = AppSession.SkipReason;
            else if (YtDlp == null) Skip = "SONGMGR_YTDLP does not name a yt-dlp.exe (UltimateKtv's Tools\\yt-dlp.exe)";
        }
    }

    /// <summary>
    /// 檢查影片 in the real program, with a real yt-dlp asking YouTube: a video that is there, one that is gone but
    /// downloaded (kept), one that is gone and not downloaded (removed after 移除提示), and one without a video ID
    /// (無法確認, kept). The 下載目錄 comes from a settings.json with // comments, as UltimateKtv writes it.
    /// </summary>
    [Trait("Category", "EndToEnd")]
    public class YoutubeCheckEndToEndTests
    {
        const string Available = "dQw4w9WgXcQ";
        const string GoneDownloaded = "xxxxxxxxxxx";
        const string GoneMissing = "zzzzzzzzzzz";

        static List<string> YoutubeIds(AppSession app)
        {
            using (DataTable dt = Db.GetDataTable(app.Database, "select Id from YoutubeSongList order by Id", ""))
            {
                return dt.AsEnumerable().Select(r => r["Id"].ToString()).ToList();
            }
        }

        [YtDlpDesktopFact]
        public void 檢查影片_removes_only_gone_videos_without_a_downloaded_file()
        {
            using (var app = new AppSession("ytcheck"))
            {
                app.WriteSettings(6);
                YoutubeSongStore.CreateTable(app.Database);
                app.Execute("insert into YoutubeSongList (Id, Name, Url, FileName) values (@a, @b, @c, @d)", "Y0001", "還在", Available, "還在_" + Available + ".mp4");
                app.Execute("insert into YoutubeSongList (Id, Name, Url, FileName) values (@a, @b, @c, @d)", "Y0002", "失效但有檔", GoneDownloaded, "失效但有檔_" + GoneDownloaded + ".mp4");
                app.Execute("insert into YoutubeSongList (Id, Name, Url, FileName) values (@a, @b, @c, @d)", "Y0003", "失效沒檔", GoneMissing, "失效沒檔_" + GoneMissing + ".mp4");
                app.Execute("insert into YoutubeSongList (Id, Name, Url, FileName) values (@a, @b, @c, @d)", "Y0004", "沒有影片ID", "", "匯入的檔案.mp4");

                // UltimateKtv's settings.json and 下載目錄, and its yt-dlp in Tools\.
                File.WriteAllText(Path.Combine(app.AppDir, "settings.json"),
                    "{\r\n  // 點播螢幕顯示\r\n  \"ShowPreviewInMainWindow\": true,\r\n\r\n  // Youtube 下載目錄: 設定 YouTube 影片下載的預設目錄。\r\n  \"YoutubeDownloadDir\": \"影片快取\",\r\n}\r\n", new UTF8Encoding(false));
                string cache = Path.Combine(app.AppDir, "影片快取");
                Directory.CreateDirectory(cache);
                File.WriteAllText(Path.Combine(cache, "失效但有檔_" + GoneDownloaded + ".mp4"), "not really a video");
                Directory.CreateDirectory(Path.Combine(app.AppDir, "Tools"));
                File.Copy(YtDlpDesktopFactAttribute.YtDlp, Path.Combine(app.AppDir, "Tools", "yt-dlp.exe"), true);

                app.Start();
                app.Choose("SongQuery_QueryType_ComboBox", "YouTube 歌曲");
                app.Click("查詢", onTab: "歌庫查詢");
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "總共查詢到 4 筆YouTube 歌曲。", 30, "the YouTube 歌曲 list");

                app.Click("檢查影片", onTab: "歌庫查詢");
                int shot = 0;
                var watch = System.Diagnostics.Stopwatch.StartNew();
                AppSession.WaitUntil(() =>
                {
                    if (watch.Elapsed.TotalSeconds > 4 * (shot + 1) && shot < 10) app.Snap("checking-" + ++shot);
                    return app.Dialogs().Any(d => d.Contains("移除提示"));
                }, 240, "檢查影片 to finish and ask 移除提示");
                app.Snap("check-results");
                Assert.Equal(new[] { "Y0001", "Y0002", "Y0003", "Y0004" }, YoutubeIds(app)); // nothing removed before the answer
                app.Answer("移除提示", "是(Y)");

                AppSession.WaitUntil(() => !YoutubeIds(app).Contains("Y0003"), 20, "失效沒檔 removed");
                Assert.Equal(new[] { "Y0001", "Y0002", "Y0004" }, YoutubeIds(app));
                AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "檢查完成: 可用 1 筆,失效但本機可播 1 筆,失效無檔已移除 1 筆,無法確認 1 筆。", 20, "the summary");
                Assert.True(File.Exists(Path.Combine(cache, "失效但有檔_" + GoneDownloaded + ".mp4")));
            }
        }
    }
}

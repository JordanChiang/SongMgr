using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Store
{
    /// <summary>
    /// 檢查影片 (<see cref="YoutubeAvailability"/>): reading yt-dlp's answer, which YouTube 歌曲 may be removed, finding
    /// UltimateKtv's 下載目錄 in its settings.json, and running yt-dlp (here a .cmd that answers like it).
    /// </summary>
    public class YoutubeAvailabilityTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("ytcheck");

        public void Dispose() => TestEnv.DeleteDir(root);

        [Fact]
        public void Exit_code_0_is_available()
        {
            var r = YoutubeAvailability.Classify(0, "", hasLocalFile: false);
            Assert.Equal(YoutubeVideoState.Available, r.State);
            Assert.False(r.Removable);
            Assert.Equal("可用 (未下載)", r.Text);
            Assert.Equal("可用", YoutubeAvailability.Classify(0, "", hasLocalFile: true).Text);
        }

        [Theory]
        [InlineData("ERROR: [youtube] xxxxxxxxxxx: This video is unavailable", "This video is unavailable")]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: Video unavailable. This video has been removed by the uploader", "Video unavailable. This video has been removed by the uploader")]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: Private video. Sign in if you've been granted access to this video", "Private video. Sign in if you've been granted access to this video")]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: Video unavailable. This video is no longer available because the YouTube account associated with this video has been terminated.", "Video unavailable. This video is no longer available because the YouTube account associated with this video has been terminated.")]
        public void A_gone_video_without_a_local_file_may_be_removed(string error, string reason)
        {
            var r = YoutubeAvailability.Classify(1, "WARNING: something\r\n" + error + "\r\n", hasLocalFile: false);
            Assert.Equal(YoutubeVideoState.Unavailable, r.State);
            Assert.Equal(reason, r.Reason);
            Assert.True(r.Removable);
            Assert.Equal("已失效 (無本機檔)", r.Text);
        }

        [Fact]
        public void A_gone_video_whose_file_is_still_here_is_kept()
        {
            var r = YoutubeAvailability.Classify(1, "ERROR: [youtube] xxxxxxxxxxx: This video is unavailable", hasLocalFile: true);
            Assert.Equal(YoutubeVideoState.Unavailable, r.State);
            Assert.False(r.Removable);
            Assert.Equal("已失效 (本機可播)", r.Text);
        }

        [Theory]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: Sign in to confirm your age. This video may be inappropriate for some users.")]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: Sign in to confirm you're not a bot")]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: The uploader has not made this video available in your country")]
        [InlineData("ERROR: Unable to download API page: <urlopen error [Errno 11001] getaddrinfo failed> (caused by TransportError)")]
        [InlineData("ERROR: [youtube] dQw4w9WgXcQ: Requested format is not available")]
        public void Anything_else_is_unknown_and_never_removed(string error)
        {
            var r = YoutubeAvailability.Classify(1, error, hasLocalFile: false);
            Assert.Equal(YoutubeVideoState.Unknown, r.State);
            Assert.False(r.Removable);
            Assert.StartsWith("無法確認: ", r.Text);
        }

        [Fact]
        public void An_exit_code_without_output_is_unknown_with_the_code()
        {
            var r = YoutubeAvailability.Classify(2, "", hasLocalFile: false);
            Assert.Equal(YoutubeVideoState.Unknown, r.State);
            Assert.Equal("yt-dlp 結束代碼 2", r.Reason);
        }

        [Fact]
        public void DownloadDir_reads_settings_json_with_comments()
        {
            File.WriteAllText(Path.Combine(root, "settings.json"), "{\r\n  // 點播螢幕顯示\r\n  \"ShowPreviewInMainWindow\": true,\r\n\r\n  // Youtube 下載目錄\r\n  \"YoutubeDownloadDir\": \"影片快取\",\r\n}", new UTF8Encoding(false));
            Assert.Equal(Path.Combine(root, "影片快取"), YoutubeAvailability.DownloadDir(root, out bool fromSettings));
            Assert.True(fromSettings);
        }

        [Fact]
        public void DownloadDir_keeps_a_rooted_path()
        {
            File.WriteAllText(Path.Combine(root, "settings.json"), "{ \"YoutubeDownloadDir\": \"D:\\\\yt\" }");
            Assert.Equal(@"D:\yt", YoutubeAvailability.DownloadDir(root, out bool fromSettings));
            Assert.True(fromSettings);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("{ \"NewSongDays\": 10 }")]
        [InlineData("{ \"YoutubeDownloadDir\": \"  \" }")]
        [InlineData("not json {")]
        public void DownloadDir_falls_back_to_UltimateKtvs_default(string settings)
        {
            if (settings != null) File.WriteAllText(Path.Combine(root, "settings.json"), settings);
            Assert.Equal(Path.Combine(root, "YoutubeDownloads"), YoutubeAvailability.DownloadDir(root, out bool fromSettings));
            Assert.False(fromSettings);
        }

        [Fact]
        public void HasLocalFile_looks_for_FileName_in_the_download_dir()
        {
            string dir = Path.Combine(root, "YoutubeDownloads");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "快樂崇拜_dQw4w9WgXcQ.mp4"), "x");

            Assert.True(YoutubeAvailability.HasLocalFile(dir, "快樂崇拜_dQw4w9WgXcQ.mp4"));
            Assert.False(YoutubeAvailability.HasLocalFile(dir, "別的_abcDEF12_-z.mp4"));
            Assert.False(YoutubeAvailability.HasLocalFile(dir, ""));
            Assert.False(YoutubeAvailability.HasLocalFile(dir, DBNull.Value));
            Assert.False(YoutubeAvailability.HasLocalFile(dir, "bad|name.mp4"));
        }

        [Fact]
        public void FindYtDlp_uses_UltimateKtvs_Tools_folder()
        {
            Assert.Null(YoutubeAvailability.FindYtDlp(root));
            Directory.CreateDirectory(Path.Combine(root, "Tools"));
            File.WriteAllText(Path.Combine(root, "Tools", "yt-dlp_x86.exe"), "");
            Assert.Equal(Path.Combine(root, "Tools", "yt-dlp_x86.exe"), YoutubeAvailability.FindYtDlp(root));
            if (Environment.Is64BitOperatingSystem)
            {
                File.WriteAllText(Path.Combine(root, "Tools", "yt-dlp.exe"), "");
                Assert.Equal(Path.Combine(root, "Tools", "yt-dlp.exe"), YoutubeAvailability.FindYtDlp(root));
            }
        }

        [Fact]
        public void Check_runs_yt_dlp_and_reads_its_answer()
        {
            string fake = FakeYtDlp("@echo off\r\necho %* | findstr /c:\"v=xxxxxxxxxxx\" >nul && (echo ERROR: [youtube] xxxxxxxxxxx: This video is unavailable 1>&2 & exit /b 1)\r\necho dQw4w9WgXcQ\r\nexit /b 0\r\n");

            var ok = YoutubeAvailability.Check(fake, "dQw4w9WgXcQ", false, TimeSpan.FromSeconds(30), CancellationToken.None);
            Assert.Equal(YoutubeVideoState.Available, ok.State);

            var gone = YoutubeAvailability.Check(fake, "xxxxxxxxxxx", false, TimeSpan.FromSeconds(30), CancellationToken.None);
            Assert.Equal(YoutubeVideoState.Unavailable, gone.State);
            Assert.Equal("This video is unavailable", gone.Reason);
        }

        [Fact]
        public void Check_does_not_ask_yt_dlp_without_a_valid_video_id_or_yt_dlp()
        {
            string fake = FakeYtDlp("@echo off\r\nexit /b 1\r\n");
            Assert.Equal("沒有影片 ID", YoutubeAvailability.Check(fake, DBNull.Value, true, TimeSpan.FromSeconds(5), CancellationToken.None).Reason);
            Assert.Equal("影片 ID 格式不正確", YoutubeAvailability.Check(fake, "Y0017", true, TimeSpan.FromSeconds(5), CancellationToken.None).Reason);
            var missing = YoutubeAvailability.Check(Path.Combine(root, "none.exe"), "dQw4w9WgXcQ", false, TimeSpan.FromSeconds(5), CancellationToken.None);
            Assert.Equal(YoutubeVideoState.Unknown, missing.State);
            Assert.Equal("找不到 yt-dlp", missing.Reason);
        }

        [Fact]
        public void Check_gives_up_when_yt_dlp_is_too_slow()
        {
            string fake = FakeYtDlp("@echo off\r\nping -n 30 127.0.0.1 >nul\r\nexit /b 0\r\n");
            var r = YoutubeAvailability.Check(fake, "dQw4w9WgXcQ", false, TimeSpan.FromSeconds(1), CancellationToken.None);
            Assert.Equal(YoutubeVideoState.Unknown, r.State);
            Assert.Equal("逾時", r.Reason);
        }

        [Theory]
        [InlineData(200, "Available", "")]
        [InlineData(null, "Unknown", "YouTube 沒有回應")]
        [InlineData(503, "Unknown", "YouTube 暫時無法回應 (HTTP 503)")]
        public void An_oEmbed_answer_settles_available_or_unknown(int? status, string state, string reason)
        {
            YoutubeCheckResult? r = YoutubeAvailability.FromOEmbed(status, hasLocalFile: false);
            Assert.NotNull(r);
            Assert.Equal(state, r.Value.State.ToString());
            Assert.Equal(reason, r.Value.Reason);
            Assert.False(r.Value.Removable);
        }

        [Theory]
        [InlineData(400)] // a video that does not exist
        [InlineData(401)] // private, or not embeddable
        [InlineData(403)]
        [InlineData(404)]
        [InlineData(429)] // oEmbed's own "too many requests": nothing about the video
        public void Any_other_oEmbed_answer_is_left_to_yt_dlp(int status)
        {
            Assert.Null(YoutubeAvailability.FromOEmbed(status, hasLocalFile: false));
        }

        [Fact]
        public void CheckAll_asks_yt_dlp_only_about_videos_oEmbed_did_not_settle()
        {
            string log = Path.Combine(root, "yt-dlp.log");
            // The fake yt-dlp finds ddddddddddd and says every other video is gone.
            string fake = FakeYtDlp("@echo off\r\necho %*>>\"" + log + "\"\r\necho %* | findstr /c:\"v=ddddddddddd\" >nul && (echo ddddddddddd & exit /b 0)\r\necho ERROR: [youtube] bbbbbbbbbbb: This video is unavailable 1>&2\r\nexit /b 1\r\n");
            var statuses = new Dictionary<string, int?> { ["aaaaaaaaaaa"] = 200, ["bbbbbbbbbbb"] = 400, ["ccccccccccc"] = null, ["ddddddddddd"] = 429, ["eeeeeeeeeee"] = 503 };
            var songs = new List<(object VideoId, bool HasLocalFile)>
            {
                ("aaaaaaaaaaa", false), ("bbbbbbbbbbb", false), ("ccccccccccc", false), ("ddddddddddd", false), ("eeeeeeeeeee", false), (DBNull.Value, true), ("Y0017", false),
            };
            var results = new System.Collections.Concurrent.ConcurrentDictionary<int, YoutubeCheckResult>();

            YoutubeAvailability.CheckAll(songs, (id, _) => statuses[id], fake, TimeSpan.FromSeconds(30), CancellationToken.None, (i, r) => results[i] = r);

            Assert.Equal(7, results.Count);
            Assert.Equal(YoutubeVideoState.Available, results[0].State);
            Assert.Equal(YoutubeVideoState.Unavailable, results[1].State);
            Assert.True(results[1].Removable);
            Assert.Equal("YouTube 沒有回應", results[2].Reason);
            Assert.Equal(YoutubeVideoState.Available, results[3].State); // 429: yt-dlp settled it
            Assert.Equal("YouTube 暫時無法回應 (HTTP 503)", results[4].Reason);
            Assert.Equal("沒有影片 ID", results[5].Reason);
            Assert.Equal("影片 ID 格式不正確", results[6].Reason);
            string[] calls = File.ReadAllLines(log);
            Assert.Equal(2, calls.Length);
            Assert.Contains(calls, c => c.Contains("v=bbbbbbbbbbb"));
            Assert.Contains(calls, c => c.Contains("v=ddddddddddd"));
        }

        [Fact]
        public void CheckAll_cancelled_reports_nothing_more()
        {
            var songs = new List<(object VideoId, bool HasLocalFile)> { ("aaaaaaaaaaa", false), ("bbbbbbbbbbb", false) };
            var cancel = new CancellationTokenSource();
            cancel.Cancel();
            int reported = 0;
            YoutubeAvailability.CheckAll(songs, (id, _) => 200, null, TimeSpan.FromSeconds(5), cancel.Token, (i, r) => Interlocked.Increment(ref reported));
            Assert.Equal(0, reported);
        }

        string FakeYtDlp(string script)
        {
            string path = Path.Combine(root, "yt-dlp-" + Guid.NewGuid().ToString("N") + ".cmd");
            File.WriteAllText(path, script, Encoding.ASCII);
            return path;
        }
    }
}

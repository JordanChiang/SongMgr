using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace CrazyKTV_SongMgr
{
    /// <summary>YouTube's oEmbed answer for a video: the HTTP status, or null when no answer came (no network, timeout).</summary>
    delegate int? OEmbedProbe(string videoId, CancellationToken cancel);

    /// <summary>What YouTube and yt-dlp say about a YouTube 歌曲's video.</summary>
    enum YoutubeVideoState
    {
        /// <summary>yt-dlp could read the video.</summary>
        Available,
        /// <summary>yt-dlp said the video is gone for good: removed, private, its account closed.</summary>
        Unavailable,
        /// <summary>Anything else: no network, too many requests, a sign-in or region wall, no video ID, yt-dlp missing or too slow.</summary>
        Unknown,
    }

    /// <summary>One YouTube 歌曲 checked: the video's state, whether its downloaded file is still there, and what to do.</summary>
    struct YoutubeCheckResult
    {
        public YoutubeCheckResult(YoutubeVideoState state, bool hasLocalFile, string reason)
        {
            State = state;
            HasLocalFile = hasLocalFile;
            Reason = reason ?? "";
        }

        public YoutubeVideoState State { get; }
        public bool HasLocalFile { get; }
        public string Reason { get; }

        /// <summary>
        /// Only a video yt-dlp says is gone and whose file is not on this computer is of no more use. A gone video whose
        /// file is still here keeps playing from it, and that file is the only copy left; an Unknown is never removed.
        /// </summary>
        public bool Removable => State == YoutubeVideoState.Unavailable && !HasLocalFile;

        /// <summary>The 檢查結果 column.</summary>
        public string Text
        {
            get
            {
                switch (State)
                {
                    case YoutubeVideoState.Available: return HasLocalFile ? "可用" : "可用 (未下載)";
                    case YoutubeVideoState.Unavailable: return HasLocalFile ? "已失效 (本機可播)" : "已失效 (無本機檔)";
                    default: return (Reason == "") ? "無法確認" : "無法確認: " + Reason;
                }
            }
        }
    }

    /// <summary>
    /// 檢查影片: whether each YouTube 歌曲's video is still on YouTube, and whether its downloaded file is in UltimateKtv's
    /// YoutubeDownloadDir (settings.json next to the program). Never downloads.
    /// Asking yt-dlp (UltimateKtv's, in Tools\) takes about 2.5 s a video, 1 s of it starting yt-dlp; YouTube's oEmbed
    /// answers in about 0.1 s. So oEmbed goes first, a few at a time: its 200 means the video is there. Anything else
    /// (a gone video, but also a private one, one that cannot be embedded, an age gate) goes on to yt-dlp, which alone
    /// may call a video gone.
    /// </summary>
    static class YoutubeAvailability
    {
        public const string DefaultDownloadDir = "YoutubeDownloads";

        /// <summary>oEmbed requests at a time.</summary>
        public const int OEmbedParallel = 4;

        private static readonly HttpClient Http = CreateHttp();

        private static HttpClient CreateHttp()
        {
            var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("CrazyKTV_SongMgr");
            return http;
        }

        /// <summary>YouTube's oEmbed status for the video (200 when it is there), or null when YouTube did not answer.</summary>
        public static int? OEmbedStatus(string videoId, CancellationToken cancel)
        {
            string url = "https://www.youtube.com/oembed?format=json&url=" + Uri.EscapeDataString("https://www.youtube.com/watch?v=" + videoId);
            try
            {
                using (HttpResponseMessage response = Http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancel).GetAwaiter().GetResult())
                {
                    return (int)response.StatusCode;
                }
            }
            catch (HttpRequestException) { return null; }
            catch (TaskCanceledException) { return null; } // timeout or cancelled
        }

        /// <summary>
        /// What an oEmbed answer settles: 200 is Available; no answer or a server error is Unknown, since yt-dlp would most
        /// likely get nowhere either; any other status settles nothing and returns null, to ask yt-dlp: 400 for a video
        /// that does not exist, 401/403/404 for one that is private, gone or not embeddable, and 429, oEmbed's own
        /// "too many requests", which says nothing about the video.
        /// </summary>
        public static YoutubeCheckResult? FromOEmbed(int? status, bool hasLocalFile)
        {
            if (status == 200) return new YoutubeCheckResult(YoutubeVideoState.Available, hasLocalFile, "");
            if (status == null) return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, "YouTube 沒有回應");
            if (status >= 500) return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, "YouTube 暫時無法回應 (HTTP " + status + ")");
            return null;
        }

        /// <summary>
        /// Checks every song: first <paramref name="probe"/> (oEmbed), <see cref="OEmbedParallel"/> at a time, then
        /// <paramref name="ytDlp"/> one at a time for those it did not settle. <paramref name="done"/> gets each song's
        /// index and result as soon as it is known, from more than one thread at once. Stops when <paramref name="cancel"/>
        /// is set; the songs not checked by then get no result.
        /// </summary>
        public static void CheckAll(IReadOnlyList<(object VideoId, bool HasLocalFile)> songs, OEmbedProbe probe, string ytDlp, TimeSpan ytDlpTimeout,
            CancellationToken cancel, Action<int, YoutubeCheckResult> done)
        {
            var needYtDlp = new ConcurrentBag<int>();
            try
            {
                Parallel.For(0, songs.Count, new ParallelOptions { MaxDegreeOfParallelism = OEmbedParallel, CancellationToken = cancel }, i =>
                {
                    var song = songs[i];
                    string id = VideoIdText(song.VideoId);
                    YoutubeCheckResult? settled = Precheck(id, song.HasLocalFile) ?? FromOEmbed(probe(id, cancel), song.HasLocalFile);
                    if (cancel.IsCancellationRequested) return;
                    if (settled != null) done(i, settled.Value);
                    else needYtDlp.Add(i);
                });
            }
            catch (OperationCanceledException)
            {
                return;
            }

            foreach (int i in needYtDlp.OrderBy(i => i))
            {
                if (cancel.IsCancellationRequested) return;
                YoutubeCheckResult result = Check(ytDlp, songs[i].VideoId, songs[i].HasLocalFile, ytDlpTimeout, cancel);
                if (cancel.IsCancellationRequested) return;
                done(i, result);
            }
        }

        private static string VideoIdText(object videoId) => (videoId == null || videoId == DBNull.Value) ? "" : videoId.ToString().Trim();

        /// <summary>A video ID nobody needs to be asked about: missing or not 11 valid characters (Unknown). Null for a real ID.</summary>
        private static YoutubeCheckResult? Precheck(string id, bool hasLocalFile)
        {
            if (id == "") return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, "沒有影片 ID");
            if (!YoutubeFavorites.IsValidVideoId(id)) return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, "影片 ID 格式不正確");
            return null;
        }

        /// <summary>
        /// yt-dlp errors that mean the video will not come back. UltimateKtv's YtDlpHelper.PermanentErrorMarkers, plus
        /// "This video is unavailable", which is what yt-dlp prints for a video that does not exist.
        /// </summary>
        private static readonly string[] GoneMarkers =
        {
            "This video is unavailable", "Video unavailable", "Private video", "This video has been removed", "has been removed",
            "account associated with this video has been terminated", "Incomplete YouTube ID", "looks truncated", "is not a valid URL",
        };

        /// <summary>Reads yt-dlp's exit code and error output.</summary>
        public static YoutubeCheckResult Classify(int exitCode, string error, bool hasLocalFile)
        {
            if (exitCode == 0) return new YoutubeCheckResult(YoutubeVideoState.Available, hasLocalFile, "");

            string line = FirstErrorLine(error);
            bool gone = GoneMarkers.Any(m => (error ?? "").IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0);
            return new YoutubeCheckResult(gone ? YoutubeVideoState.Unavailable : YoutubeVideoState.Unknown, hasLocalFile,
                (line != "") ? line : "yt-dlp 結束代碼 " + exitCode);
        }

        /// <summary>The first "ERROR:" line, without "ERROR: [youtube] id: ", or the first line.</summary>
        private static string FirstErrorLine(string error)
        {
            string[] lines = (error ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(l => l.Trim()).Where(l => l != "").ToArray();
            string line = lines.FirstOrDefault(l => l.StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase)) ?? lines.FirstOrDefault() ?? "";
            if (line.StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase)) line = line.Substring(6).Trim();
            int bracket = line.IndexOf("] ", StringComparison.Ordinal);
            if (line.StartsWith("[", StringComparison.Ordinal) && bracket > 0)
            {
                line = line.Substring(bracket + 2);
                int colon = line.IndexOf(": ", StringComparison.Ordinal);
                if (colon > 0 && YoutubeFavorites.IsValidVideoId(line.Substring(0, colon))) line = line.Substring(colon + 2);
            }
            return line;
        }

        /// <summary>
        /// The yt-dlp in <paramref name="appDir"/>\Tools, as UltimateKtv puts it there: yt-dlp.exe on 64-bit Windows
        /// (SongMgr itself is 32-bit, but that does not matter for another program), else yt-dlp_x86.exe. Null when neither is there.
        /// </summary>
        public static string FindYtDlp(string appDir)
        {
            string tools = Path.Combine(appDir, "Tools");
            string[] names = Environment.Is64BitOperatingSystem ? new[] { "yt-dlp.exe", "yt-dlp_x86.exe" } : new[] { "yt-dlp_x86.exe" };
            return names.Select(n => Path.Combine(tools, n)).FirstOrDefault(File.Exists);
        }

        /// <summary>
        /// UltimateKtv's YouTube 下載目錄: YoutubeDownloadDir of <paramref name="appDir"/>\settings.json (a JSON file with
        /// // comments), relative to <paramref name="appDir"/> unless rooted. <paramref name="fromSettings"/> is false when
        /// the file or the setting is missing or cannot be read, and the default <see cref="DefaultDownloadDir"/> is used.
        /// </summary>
        public static string DownloadDir(string appDir, out bool fromSettings)
        {
            string dir = DefaultDownloadDir;
            fromSettings = false;
            try
            {
                string file = Path.Combine(appDir, "settings.json");
                if (File.Exists(file))
                {
                    JToken value = JObject.Parse(File.ReadAllText(file, Encoding.UTF8))["YoutubeDownloadDir"];
                    if (value != null && value.Type == JTokenType.String && value.ToString().Trim() != "")
                    {
                        dir = value.ToString().Trim();
                        fromSettings = true;
                    }
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is Newtonsoft.Json.JsonException)
            {
                fromSettings = false;
            }
            return Path.IsPathRooted(dir) ? dir : Path.Combine(appDir, dir);
        }

        /// <summary>Whether the YouTube 歌曲's FileName is in <paramref name="downloadDir"/> (as UltimateKtv joins them).</summary>
        public static bool HasLocalFile(string downloadDir, object fileName)
        {
            string name = (fileName == null || fileName == DBNull.Value) ? "" : fileName.ToString().Trim();
            if (name == "") return false;
            try
            {
                return File.Exists(Path.Combine(downloadDir, name));
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        /// <summary>
        /// Runs <paramref name="ytDlp"/> on the video, reading only its details (no download), and classifies the answer.
        /// A video ID that is not 11 valid characters is Unknown without asking. Gives up after <paramref name="timeout"/>,
        /// or when <paramref name="cancel"/> is set, with Unknown.
        /// </summary>
        public static YoutubeCheckResult Check(string ytDlp, object videoId, bool hasLocalFile, TimeSpan timeout, CancellationToken cancel)
        {
            string id = VideoIdText(videoId);
            YoutubeCheckResult? invalid = Precheck(id, hasLocalFile);
            if (invalid != null) return invalid.Value;
            if (ytDlp == null || !File.Exists(ytDlp)) return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, "找不到 yt-dlp");

            var start = new ProcessStartInfo(ytDlp, "--simulate --no-playlist --no-warnings --encoding utf-8 --print id \"https://www.youtube.com/watch?v=" + id + "\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            try
            {
                using (Process process = Process.Start(start))
                {
                    var output = process.StandardOutput.ReadToEndAsync();
                    var error = process.StandardError.ReadToEndAsync();
                    var watch = Stopwatch.StartNew();
                    while (!process.WaitForExit(200))
                    {
                        if (cancel.IsCancellationRequested || watch.Elapsed > timeout)
                        {
                            try { process.Kill(); } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { }
                            process.WaitForExit();
                            return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, cancel.IsCancellationRequested ? "已取消" : "逾時");
                        }
                    }
                    process.WaitForExit();
                    return Classify(process.ExitCode, error.Result, hasLocalFile);
                }
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                return new YoutubeCheckResult(YoutubeVideoState.Unknown, hasLocalFile, "無法執行 yt-dlp: " + ex.Message);
            }
        }
    }
}

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 程式自動更新: 從 GitHub Releases 取得最新版本.
    /// release 附有 update.json 時, 依清單裡每個元件 (主程式 / Codec_x86 / Tools) 的檔案雜湊值
    /// 比對本機檔案, 只下載有變動的元件; 沒有清單時退回下載整包 zip.
    /// 下載的檔案解壓到暫存資料夾, 程式結束後由批次檔覆蓋回程式資料夾並重新啟動.
    /// update.json 與各元件 zip 由 tools\pack-release.ps1 產生.
    /// </summary>
    public static class AppUpdater
    {
        public const string LatestReleaseApi = "https://api.github.com/repos/JordanChiang/SongMgr/releases/latest";
        public const string ExeName = "CrazyKTV_SongMgr.exe";
        public const string ManifestName = "update.json";

        /// <summary>程式資料夾已有這些檔案時保留不覆蓋 (使用者的參考資料庫、自訂語系及設定).</summary>
        public static readonly string[] KeepExisting =
        {
            @"SongMgr\CrazySongMgr.mdb",
            @"SongMgr\CrazySongMgr.db",
            @"SongMgr\Custom.lang",
            "CrazyKTV_SongMgr.cfg",
        };

        public class ReleaseInfo
        {
            public Version Version;
            public string Tag;
            public string PageUrl;
            public string ManifestUrl;
            public string ZipUrl;
            public string ZipName;
            public long ZipSize;
        }

        /// <summary>更新清單裡的一個元件; Files 為 null 表示沒有清單的整包 zip.</summary>
        public class Component
        {
            public string Name;
            public string Desc;
            public string Url;
            public long Size;
            public string DisplayName => string.IsNullOrEmpty(Desc) ? Name : Desc;
            /// <summary>相對於程式資料夾的路徑 (以 / 分隔) 對應小寫 SHA256.</summary>
            public Dictionary<string, string> Files;
        }

        public class UpdatePlan
        {
            public ReleaseInfo Release;
            public List<Component> Components = new List<Component>();
            public bool FromManifest;
            public long TotalSize => Components.Sum(c => c.Size);
        }

        private static readonly HttpClient Http = CreateHttpClient();

        private static HttpClient CreateHttpClient()
        {
            HttpClient client = new HttpClient(new HttpClientHandler { AutomaticDecompression = DecompressionMethods.All });
            client.Timeout = TimeSpan.FromMinutes(30);
            client.DefaultRequestHeaders.UserAgent.ParseAdd("CrazyKTV_SongMgr-Updater");
            return client;
        }

        #region --- 版本 ---

        /// <summary>把 "v2.11.2"、"2.11.2.0"、"2.11" 之類的字串轉成四段版本 (缺的段補 0); 無法解析時回傳 null.</summary>
        public static Version ParseVersion(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            Match m = Regex.Match(text.Trim(), @"^[vV]?(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?");
            if (!m.Success) return null;
            int Part(int i) => m.Groups[i].Success ? int.Parse(m.Groups[i].Value) : 0;
            try { return new Version(Part(1), Part(2), Part(3), Part(4)); }
            catch { return null; }
        }

        public static string CurrentVersionText()
        {
            try { return FileVersionInfo.GetVersionInfo(Environment.ProcessPath).FileVersion ?? ""; }
            catch { return ""; }
        }

        public static Version CurrentVersion() => ParseVersion(CurrentVersionText());

        public static bool IsNewer(ReleaseInfo release, Version current)
        {
            if (release == null || release.Version == null) return false;
            if (string.IsNullOrEmpty(release.ManifestUrl) && string.IsNullOrEmpty(release.ZipUrl)) return false;
            if (current == null) return false;
            return release.Version > current;
        }

        #endregion

        #region --- 取得最新版本資訊 ---

        /// <summary>解析 GitHub releases/latest API 的回應; 草稿、預覽版或沒有 update.json / zip 附檔時回傳 null.</summary>
        public static ReleaseInfo ParseRelease(string json)
        {
            try
            {
                JObject obj = JObject.Parse(json);
                if ((bool?)obj["draft"] == true || (bool?)obj["prerelease"] == true) return null;

                string tag = (string)obj["tag_name"];
                Version ver = ParseVersion(tag);
                if (ver == null) return null;

                JArray assets = obj["assets"] as JArray ?? new JArray();
                JToken manifest = assets.FirstOrDefault(a => string.Equals((string)a["name"], ManifestName, StringComparison.OrdinalIgnoreCase));

                // 沒有清單時用來整包更新的 zip: 優先取名稱含 SongMgr 的完整包
                List<JToken> zips = assets.Where(a => ((string)a["name"] ?? "").EndsWith(".zip", StringComparison.OrdinalIgnoreCase)).ToList();
                JToken zip = zips.FirstOrDefault(a => Regex.IsMatch((string)a["name"], @"^CrazyKTV_SongMgr[\d.]+\.zip$", RegexOptions.IgnoreCase))
                             ?? zips.FirstOrDefault();

                if (manifest == null && zip == null) return null;

                return new ReleaseInfo
                {
                    Version = ver,
                    Tag = tag,
                    PageUrl = (string)obj["html_url"],
                    ManifestUrl = (string)manifest?["browser_download_url"],
                    ZipUrl = (string)zip?["browser_download_url"],
                    ZipName = (string)zip?["name"],
                    ZipSize = (long?)zip?["size"] ?? 0,
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>向 GitHub 查詢最新版本; 網路錯誤時回傳 null.</summary>
        public static ReleaseInfo GetLatestRelease()
        {
            string json = GetString(LatestReleaseApi, "application/vnd.github+json");
            return json == null ? null : ParseRelease(json);
        }

        private static string GetString(string url, string accept)
        {
            try
            {
                using (HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Get, url))
                using (CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(15)))
                {
                    if (accept != null) req.Headers.Accept.ParseAdd(accept);
                    using (HttpResponseMessage res = Http.Send(req, cts.Token))
                    {
                        if (!res.IsSuccessStatusCode) return null;
                        using (StreamReader sr = new StreamReader(res.Content.ReadAsStream(cts.Token), Encoding.UTF8))
                        {
                            return sr.ReadToEnd();
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region --- 更新清單 ---

        /// <summary>解析 update.json; 格式錯誤時回傳 null.</summary>
        public static List<Component> ParseManifest(string json)
        {
            try
            {
                JObject obj = JObject.Parse(json);
                List<Component> list = new List<Component>();
                foreach (JToken c in obj["components"] as JArray ?? new JArray())
                {
                    Component comp = new Component
                    {
                        Name = (string)c["name"],
                        Desc = (string)c["desc"],
                        Url = (string)c["url"],
                        Size = (long?)c["size"] ?? 0,
                        Files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                    };
                    if (string.IsNullOrEmpty(comp.Name) || string.IsNullOrEmpty(comp.Url)) return null;
                    foreach (JProperty f in (c["files"] as JObject ?? new JObject()).Properties())
                    {
                        comp.Files[f.Name] = ((string)f.Value ?? "").ToLowerInvariant();
                    }
                    if (comp.Files.Count == 0) return null;
                    list.Add(comp);
                }
                return list.Count > 0 ? list : null;
            }
            catch
            {
                return null;
            }
        }

        public static string LocalPath(string appDir, string relPath) =>
            Path.Combine(appDir, relPath.Replace('/', Path.DirectorySeparatorChar));

        public static bool IsKept(string relPath)
        {
            string rel = relPath.Replace('/', '\\');
            return KeepExisting.Any(k => string.Equals(k, rel, StringComparison.OrdinalIgnoreCase));
        }

        public static string FileHash(string file)
        {
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                return Convert.ToHexString(SHA256.HashData(fs)).ToLowerInvariant();
            }
        }

        /// <summary>元件裡有任何檔案不存在或雜湊值不同就需要更新 (使用者已有的保留檔案不算).</summary>
        public static bool NeedsUpdate(Component comp, string appDir)
        {
            foreach (KeyValuePair<string, string> f in comp.Files)
            {
                string local = LocalPath(appDir, f.Key);
                bool exists = File.Exists(local);
                if (IsKept(f.Key) && exists) continue;
                if (!exists) return true;
                try
                {
                    if (!string.Equals(FileHash(local), f.Value, StringComparison.OrdinalIgnoreCase)) return true;
                }
                catch
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>依 release 決定要下載哪些元件; 已是最新或無法取得資訊時回傳 null.</summary>
        public static UpdatePlan GetUpdatePlan(ReleaseInfo release, string appDir)
        {
            UpdatePlan plan = new UpdatePlan { Release = release };

            if (!string.IsNullOrEmpty(release.ManifestUrl))
            {
                string json = GetString(release.ManifestUrl, null);
                List<Component> comps = json == null ? null : ParseManifest(json);
                if (comps != null)
                {
                    plan.FromManifest = true;
                    plan.Components.AddRange(comps.Where(c => NeedsUpdate(c, appDir)));
                    return plan.Components.Count > 0 ? plan : null;
                }
            }

            if (string.IsNullOrEmpty(release.ZipUrl)) return null;
            plan.Components.Add(new Component { Name = release.ZipName, Url = release.ZipUrl, Size = release.ZipSize });
            return plan;
        }

        #endregion

        #region --- 下載與準備更新檔 ---

        public static string StageDir => Path.Combine(Path.GetTempPath(), "CrazyKTV_SongMgr_Update");

        public static bool CanWrite(string dir)
        {
            try
            {
                string probe = Path.Combine(dir, "~update_probe_" + Guid.NewGuid().ToString("N") + ".tmp");
                File.WriteAllBytes(probe, new byte[0]);
                File.Delete(probe);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>下載並解壓計畫中的元件, 回傳要覆蓋到程式資料夾的根目錄 (已移除要保留的檔案).</summary>
        public static string Prepare(UpdatePlan plan, string appDir, Action<long, long> progress, CancellationToken ct)
        {
            string stage = StageDir;
            if (Directory.Exists(stage)) Directory.Delete(stage, true);
            Directory.CreateDirectory(stage);

            string extractDir = Path.Combine(stage, "files");
            long total = plan.TotalSize;
            long done = 0;

            for (int i = 0; i < plan.Components.Count; i++)
            {
                Component comp = plan.Components[i];
                string zipFile = Path.Combine(stage, "package" + i + ".zip");
                long before = done;
                Download(comp.Url, zipFile, (d, t) => { done = before + d; progress?.Invoke(done, Math.Max(total, done)); }, ct);

                ct.ThrowIfCancellationRequested();
                ZipFile.ExtractToDirectory(zipFile, extractDir, Encoding.UTF8, true);
                File.Delete(zipFile);

                if (comp.Files != null) VerifyComponent(comp, extractDir);
            }

            string root = plan.FromManifest ? extractDir : FindPackageRoot(extractDir);
            if (root == null) throw new InvalidDataException("更新檔裡找不到 " + ExeName);

            RemoveKeptFiles(root, appDir);
            return root;
        }

        private static void Download(string url, string file, Action<long, long> progress, CancellationToken ct)
        {
            using (HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Get, url))
            using (HttpResponseMessage res = Http.Send(req, HttpCompletionOption.ResponseHeadersRead, ct))
            {
                res.EnsureSuccessStatusCode();
                long total = res.Content.Headers.ContentLength ?? 0;
                using (Stream src = res.Content.ReadAsStream(ct))
                using (FileStream dst = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    byte[] buffer = new byte[81920];
                    long done = 0;
                    int n;
                    while ((n = src.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ct.ThrowIfCancellationRequested();
                        dst.Write(buffer, 0, n);
                        done += n;
                        progress?.Invoke(done, total);
                    }
                }
            }
        }

        /// <summary>確認解壓出來的檔案與清單一致, 避免下載不完整或附檔與清單不符.</summary>
        public static void VerifyComponent(Component comp, string extractDir)
        {
            foreach (KeyValuePair<string, string> f in comp.Files)
            {
                string file = LocalPath(extractDir, f.Key);
                if (!File.Exists(file)) throw new InvalidDataException(comp.Name + " 更新檔缺少 " + f.Key);
                if (!string.Equals(FileHash(file), f.Value, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(comp.Name + " 更新檔的 " + f.Key + " 與更新清單不符");
            }
        }

        /// <summary>zip 可能直接放檔案, 也可能多包一層資料夾: 找出 CrazyKTV_SongMgr.exe 所在的資料夾.</summary>
        public static string FindPackageRoot(string extractDir)
        {
            if (File.Exists(Path.Combine(extractDir, ExeName))) return extractDir;
            string exe = Directory.GetFiles(extractDir, ExeName, SearchOption.AllDirectories)
                .OrderBy(f => f.Length)
                .FirstOrDefault();
            return exe == null ? null : Path.GetDirectoryName(exe);
        }

        /// <summary>程式資料夾已有的使用者檔案, 從更新檔中移除以免被覆蓋.</summary>
        public static void RemoveKeptFiles(string packageRoot, string appDir)
        {
            foreach (string rel in KeepExisting)
            {
                string staged = Path.Combine(packageRoot, rel);
                if (File.Exists(staged) && File.Exists(Path.Combine(appDir, rel))) File.Delete(staged);
            }
        }

        #endregion

        #region --- 程式結束後覆蓋檔案 ---

        /// <summary>
        /// 參數: %1 = 要等待結束的 PID, %2 = 更新檔資料夾, %3 = 程式資料夾, %4 = 暫存資料夾.
        /// 路徑由命令列傳入 (Unicode), 批次檔本身只有 ASCII, 不受主控台字碼頁影響.
        /// </summary>
        public const string ScriptText =
            "@echo off\r\n" +
            "setlocal\r\n" +
            "set \"PID=%~1\"\r\n" +
            "set \"SRC=%~2\"\r\n" +
            "set \"DST=%~3\"\r\n" +
            "set \"STAGE=%~4\"\r\n" +
            "set /a TRIES=0\r\n" +
            ":wait\r\n" +
            "tasklist /FI \"PID eq %PID%\" /NH 2>nul | find \" %PID% \" >nul\r\n" +
            "if errorlevel 1 goto copy\r\n" +
            "set /a TRIES+=1\r\n" +
            "if %TRIES% GEQ 120 goto copy\r\n" +
            "ping -n 2 127.0.0.1 >nul\r\n" +
            "goto wait\r\n" +
            ":copy\r\n" +
            "robocopy \"%SRC%\" \"%DST%\" /E /R:10 /W:1 /NFL /NDL /NJH /NJS /NP >nul\r\n" +
            "start \"\" \"%DST%\\" + ExeName + "\" " + RestartArg + "\r\n" +
            "rd /s /q \"%STAGE%\" 2>nul\r\n" +
            "(goto) 2>nul & del \"%~f0\"\r\n";

        /// <summary>啟動批次檔; 呼叫後應立即關閉程式, 批次檔會等這個行程結束再覆蓋檔案.</summary>
        public static void LaunchReplace(string packageRoot, string appDir)
        {
            string script = Path.Combine(Path.GetTempPath(), "CrazyKTV_SongMgr_Update.cmd");
            File.WriteAllText(script, ScriptText, Encoding.ASCII);

            string Arg(string path) => "\"" + path.TrimEnd('\\') + "\"";
            ProcessStartInfo psi = new ProcessStartInfo("cmd.exe")
            {
                Arguments = "/c \"" + Arg(script) + " " + Environment.ProcessId + " " + Arg(packageRoot) + " " + Arg(appDir) + " " + Arg(StageDir) + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
            };
            Process.Start(psi);
        }

        public static void OpenPage(string url)
        {
            try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { }
        }

        #endregion

        #region --- 更新後重新啟動 ---

        /// <summary>批次檔重新啟動程式時帶的參數, 讓程式知道要把視窗帶到最上層.</summary>
        public const string RestartArg = "--updated";

        public static bool LaunchedAfterUpdate =>
            Environment.GetCommandLineArgs().Skip(1).Any(a => string.Equals(a, RestartArg, StringComparison.OrdinalIgnoreCase));

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// 由背景批次檔啟動的程式不會自動取得焦點, 會被其他視窗蓋住:
        /// 暫時設為最上層並取得焦點, 再還原成使用者的「視窗置頂」設定.
        /// </summary>
        public static void BringToFront(Form form, bool alwaysOnTop)
        {
            if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
            form.TopMost = true;
            form.BringToFront();
            form.Activate();
            SetForegroundWindow(form.Handle);
            form.TopMost = alwaysOnTop;
        }

        #endregion
    }

    /// <summary>下載更新檔時顯示進度, 可取消. DialogResult.OK 表示更新檔已備妥 (PackageRoot).</summary>
    public class AppUpdateProgressForm : Form
    {
        private readonly AppUpdater.UpdatePlan plan;
        private readonly string appDir;
        private readonly CancellationTokenSource cts = new CancellationTokenSource();
        private readonly Label status = new Label();
        private readonly ProgressBar bar = new ProgressBar();
        private readonly Button cancel = new Button();

        public string PackageRoot { get; private set; }
        public Exception Error { get; private set; }

        public AppUpdateProgressForm(AppUpdater.UpdatePlan plan, string appDir)
        {
            this.plan = plan;
            this.appDir = appDir;

            Text = "下載 CrazyKTV 加歌程式 " + plan.Release.Tag;
            Font = new Font("微軟正黑體", 10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(420, 120);

            status.SetBounds(16, 14, 388, 24);
            status.Text = "正在下載更新檔...";
            bar.SetBounds(16, 42, 388, 22);
            bar.Maximum = 1000;
            cancel.SetBounds(314, 76, 90, 30);
            cancel.Text = "取消";
            cancel.Click += (s, e) => { cancel.Enabled = false; cts.Cancel(); };

            Controls.AddRange(new Control[] { status, bar, cancel });
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Task.Run(() =>
            {
                try
                {
                    PackageRoot = AppUpdater.Prepare(plan, appDir, ReportProgress, cts.Token);
                }
                catch (Exception ex)
                {
                    Error = ex;
                }
                BeginInvoke((Action)(() =>
                {
                    DialogResult = PackageRoot != null ? DialogResult.OK : DialogResult.Cancel;
                    Close();
                }));
            });
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 使用者按 X: 取消下載, 等背景工作收尾後才真正關閉
            if (PackageRoot == null && Error == null && e.CloseReason == CloseReason.UserClosing)
            {
                cts.Cancel();
                cancel.Enabled = false;
                e.Cancel = true;
                return;
            }
            base.OnFormClosing(e);
        }

        private long lastReport;

        private void ReportProgress(long done, long total)
        {
            long now = Environment.TickCount64;
            if (now - lastReport < 100 && done != total) return;
            lastReport = now;
            BeginInvoke((Action)(() =>
            {
                string mb(long b) => (b / 1048576.0).ToString("0.0") + " MB";
                if (total > 0)
                {
                    bar.Value = (int)Math.Min(1000, done * 1000 / total);
                    status.Text = "正在下載更新檔... " + mb(done) + " / " + mb(total);
                }
                else
                {
                    status.Text = "正在下載更新檔... " + mb(done);
                }
                if (total > 0 && done >= total) status.Text = "正在解壓縮並檢查更新檔...";
            }));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) cts.Dispose();
            base.Dispose(disposing);
        }
    }
}

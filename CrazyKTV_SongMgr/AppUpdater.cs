using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 程式自動更新: 從 GitHub Releases 取得最新版本.
    /// release 附有 update.json 時, 依清單裡每個元件的檔案雜湊值比對本機檔案, 只下載有變動的元件;
    /// 沒有清單時退回下載完整包 zip.
    /// 下載的檔案解壓到暫存資料夾, 程式結束後由批次檔覆蓋回程式資料夾並重新啟動, 覆蓋失敗時還原原本的檔案.
    /// 元件的定義、update.json 與各元件 zip 都由 tools\pack-release.ps1 產生.
    /// </summary>
    public static class AppUpdater
    {
        public const string LatestReleaseApi = "https://api.github.com/repos/JordanChiang/SongMgr/releases/latest";
        public const string ReleasesPage = "https://github.com/JordanChiang/SongMgr/releases/latest";
        public const string ExeName = "CrazyKTV_SongMgr.exe";
        public const string ManifestName = "update.json";

        /// <summary>完整包的檔名, 例如 CrazyKtv_SongMgr2.12.0.zip (由 pack-release.ps1 產生).</summary>
        public const string FullPackagePattern = @"^CrazyKTV_SongMgr[\d.]+\.zip$";

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

        /// <summary>位元組數轉成 "12.3 MB".</summary>
        public static string FormatMB(long bytes) => (bytes / 1048576.0).ToString("0.0") + " MB";

        #region --- 版本 ---

        /// <summary>把 "v2.11.2"、"2.11.2.0"、"2.11" 之類的字串轉成四段版本 (缺的段補 0); 無法解析時回傳 null.</summary>
        public static Version ParseVersion(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            Match m = Regex.Match(text.Trim(), @"^[vV]?(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?");
            if (!m.Success) return null;
            if (!TryPart(m, 1, out int major) || !TryPart(m, 2, out int minor) ||
                !TryPart(m, 3, out int build) || !TryPart(m, 4, out int revision))
            {
                return null;
            }
            return new Version(major, minor, build, revision);
        }

        private static bool TryPart(Match m, int group, out int value)
        {
            value = 0;
            return !m.Groups[group].Success || int.TryParse(m.Groups[group].Value, out value);
        }

        public static string CurrentVersionText()
        {
            try
            {
                return FileVersionInfo.GetVersionInfo(Environment.ProcessPath).FileVersion ?? "";
            }
            catch
            {
                return "";
            }
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

        /// <summary>解析 GitHub releases/latest API 的回應; 草稿、預覽版或沒有 update.json / 完整包附檔時回傳 null.</summary>
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

                // 沒有清單時用來整包更新的 zip: 只認完整包, 不能拿元件 zip (例如 Codec_x86_*.zip) 代替
                JToken zip = assets.FirstOrDefault(a => Regex.IsMatch((string)a["name"] ?? "", FullPackagePattern, RegexOptions.IgnoreCase));

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

        /// <summary>
        /// 依 release 決定要下載哪些元件; 已是最新或無法取得資訊時回傳 null.
        /// release 有 update.json 卻讀不到或格式錯誤時也回傳 null (下次啟動再檢查), 不改抓完整包.
        /// </summary>
        public static UpdatePlan GetUpdatePlan(ReleaseInfo release, string appDir)
        {
            UpdatePlan plan = new UpdatePlan { Release = release };

            if (!string.IsNullOrEmpty(release.ManifestUrl))
            {
                string json = GetString(release.ManifestUrl, null);
                List<Component> comps = json == null ? null : ParseManifest(json);
                if (comps == null) return null;

                plan.FromManifest = true;
                plan.Components.AddRange(comps.Where(c => NeedsUpdate(c, appDir)));
                return plan.Components.Count > 0 ? plan : null;
            }

            // 沒有 update.json 的舊 release: 下載完整包
            if (string.IsNullOrEmpty(release.ZipUrl)) return null;
            plan.Components.Add(new Component { Name = release.ZipName, Url = release.ZipUrl, Size = release.ZipSize });
            return plan;
        }

        #endregion

        #region --- 下載與準備更新檔 ---

        public static string StageDir => Path.Combine(Path.GetTempPath(), "CrazyKTV_SongMgr_Update");

        /// <summary>覆蓋前程式資料夾裡會被取代的檔案備份在這裡; 覆蓋與還原都失敗時才會留下.</summary>
        public static string BackupDir => Path.Combine(StageDir, "backup");

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

        /// <summary>
        /// 把程式資料夾裡會被更新檔取代的檔案複製到 backupDir (保持相對路徑), 覆蓋失敗時由批次檔還原.
        /// 更新檔裡的新檔案沒有對應的舊檔, 不備份; 還原後會留在程式資料夾, 不影響舊版執行.
        /// </summary>
        public static void BackupReplacedFiles(string packageRoot, string appDir, string backupDir)
        {
            if (Directory.Exists(backupDir)) Directory.Delete(backupDir, true);
            Directory.CreateDirectory(backupDir);

            foreach (string staged in Directory.GetFiles(packageRoot, "*", SearchOption.AllDirectories))
            {
                string rel = Path.GetRelativePath(packageRoot, staged);
                string current = Path.Combine(appDir, rel);
                if (!File.Exists(current)) continue;

                string backup = Path.Combine(backupDir, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(backup));
                File.Copy(current, backup, true);
            }
        }

        #endregion

        #region --- 程式結束後覆蓋檔案 ---

        /// <summary>批次檔等程式結束的最長秒數; 逾時仍未結束就放棄這次更新, 不動任何檔案.</summary>
        public const int WaitSeconds = 600;

        /// <summary>
        /// 參數: %1 = 要等待結束的 PID, %2 = 更新檔資料夾, %3 = 程式資料夾, %4 = 暫存資料夾, %5 = 備份資料夾.
        /// robocopy 結束碼 8 以上表示有檔案複製失敗: 改把備份還原回程式資料夾, 再以 --update-failed 啟動;
        /// 還原也失敗時保留暫存資料夾 (含備份), 讓使用者手動複製.
        /// 系統工具一律以 %SystemRoot%\System32 完整路徑呼叫, 避免 PATH 上同名的工具 (例如 Git 的 find) 搶先.
        /// robocopy 加 /IS /IT: 大小與時間相同的檔案也照樣覆蓋.
        /// 路徑由命令列傳入 (Unicode), 批次檔本身只有 ASCII, 不受主控台字碼頁影響.
        /// </summary>
        public static readonly string ScriptText =
            "@echo off\r\n" +
            "setlocal\r\n" +
            "set \"PID=%~1\"\r\n" +
            "set \"SRC=%~2\"\r\n" +
            "set \"DST=%~3\"\r\n" +
            "set \"STAGE=%~4\"\r\n" +
            "set \"BAK=%~5\"\r\n" +
            "set \"SYS=%SystemRoot%\\System32\"\r\n" +
            "set /a TRIES=0\r\n" +
            ":wait\r\n" +
            "\"%SYS%\\tasklist.exe\" /FI \"PID eq %PID%\" /NH 2>nul | \"%SYS%\\find.exe\" \" %PID% \" >nul\r\n" +
            "if errorlevel 1 goto copy\r\n" +
            "set /a TRIES+=1\r\n" +
            "if %TRIES% GEQ " + WaitSeconds + " goto abort\r\n" +
            "\"%SYS%\\ping.exe\" -n 2 127.0.0.1 >nul\r\n" +
            "goto wait\r\n" +
            ":abort\r\n" +
            "rd /s /q \"%STAGE%\" 2>nul\r\n" +
            "goto end\r\n" +
            ":copy\r\n" +
            "\"%SYS%\\robocopy.exe\" \"%SRC%\" \"%DST%\" /E /IS /IT /R:10 /W:1 /NFL /NDL /NJH /NJS /NP >nul\r\n" +
            "if errorlevel 8 goto rollback\r\n" +
            "start \"\" \"%DST%\\" + ExeName + "\" " + RestartArg + "\r\n" +
            "rd /s /q \"%STAGE%\" 2>nul\r\n" +
            "goto end\r\n" +
            ":rollback\r\n" +
            "\"%SYS%\\robocopy.exe\" \"%BAK%\" \"%DST%\" /E /IS /IT /R:10 /W:1 /NFL /NDL /NJH /NJS /NP >nul\r\n" +
            "if errorlevel 8 goto failed\r\n" +
            "rd /s /q \"%STAGE%\" 2>nul\r\n" +
            ":failed\r\n" +
            "start \"\" \"%DST%\\" + ExeName + "\" " + FailedArg + "\r\n" +
            ":end\r\n" +
            "(goto) 2>nul & del \"%~f0\"\r\n";

        /// <summary>備份會被取代的檔案並啟動批次檔; 呼叫後應立即關閉程式, 批次檔會等這個行程結束再覆蓋檔案.</summary>
        public static void LaunchReplace(string packageRoot, string appDir)
        {
            BackupReplacedFiles(packageRoot, appDir, BackupDir);

            string script = Path.Combine(Path.GetTempPath(), "CrazyKTV_SongMgr_Update.cmd");
            File.WriteAllText(script, ScriptText, Encoding.ASCII);

            string Arg(string path) => "\"" + path.TrimEnd('\\') + "\"";
            ProcessStartInfo psi = new ProcessStartInfo("cmd.exe")
            {
                Arguments = "/c \"" + Arg(script) + " " + Environment.ProcessId + " " + Arg(packageRoot) + " " + Arg(appDir) + " " + Arg(StageDir) + " " + Arg(BackupDir) + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
            };
            Process.Start(psi);
        }

        /// <summary>以預設瀏覽器開啟網頁; 開不了時顯示網址讓使用者自行前往.</summary>
        public static void OpenPage(string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch
            {
                MessageBox.Show("無法開啟瀏覽器, 請自行前往:" + Environment.NewLine + url, "開啟下載頁面", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion

        #region --- 更新後重新啟動 ---

        /// <summary>批次檔重新啟動程式時帶的參數, 讓程式知道要把視窗帶到最上層.</summary>
        public const string RestartArg = "--updated";

        /// <summary>覆蓋檔案失敗時批次檔帶的參數, 讓程式告知使用者更新失敗.</summary>
        public const string FailedArg = "--update-failed";

        public static bool LaunchedAfterUpdate => HasArg(RestartArg);

        public static bool LaunchedAfterFailedUpdate => HasArg(FailedArg);

        private static bool HasArg(string arg) =>
            Environment.GetCommandLineArgs().Skip(1).Any(a => string.Equals(a, arg, StringComparison.OrdinalIgnoreCase));

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
            NativeMethods.SetForegroundWindow(form.Handle);
            form.TopMost = alwaysOnTop;
        }

        #endregion
    }
}

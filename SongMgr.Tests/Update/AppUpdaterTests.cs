using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Update
{
    /// <summary>程式自動更新: 版本比對、GitHub release 與 update.json 解析、依元件比對下載.</summary>
    public class AppUpdaterTests : IDisposable
    {
        readonly string dir = Path.Combine(Path.GetTempPath(), "AppUpdaterTests_" + Guid.NewGuid().ToString("N"));

        public AppUpdaterTests() => Directory.CreateDirectory(dir);

        public void Dispose()
        {
            try { Directory.Delete(dir, true); } catch { }
        }

        string Write(string root, string rel, string content)
        {
            string path = AppUpdater.LocalPath(root, rel);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content);
            return path;
        }

        static string Hash(string content) =>
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(content))).ToLowerInvariant();

        static AppUpdater.Component Comp(string name, params (string rel, string content)[] files) => new AppUpdater.Component
        {
            Name = name,
            Url = "https://x/" + name + ".zip",
            Files = files.ToDictionary(f => f.rel, f => Hash(f.content), StringComparer.OrdinalIgnoreCase),
        };

        #region 版本

        [Theory]
        [InlineData("2.11.2", "2.11.2.0")]
        [InlineData("v2.11.2", "2.11.2.0")]
        [InlineData("2.11", "2.11.0.0")]
        [InlineData("2.11.2.7", "2.11.2.7")]
        [InlineData("2.11.2-beta", "2.11.2.0")]
        public void ParseVersion_normalizes_to_four_parts(string text, string expected)
        {
            Assert.Equal(Version.Parse(expected), AppUpdater.ParseVersion(text));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("latest")]
        public void ParseVersion_returns_null_for_non_versions(string text)
        {
            Assert.Null(AppUpdater.ParseVersion(text));
        }

        [Theory]
        // 舊做法去掉小數點比數字: 2.10.11 -> 21011 < 2.9.1 -> 29100, 會判斷錯誤
        [InlineData("2.10.11", "2.9.1", true)]
        [InlineData("2.11.3", "2.11.2", true)]
        [InlineData("2.11.2", "2.11.2", false)]
        [InlineData("2.11.1", "2.11.2", false)]
        public void IsNewer_compares_versions_numerically(string remote, string local, bool expected)
        {
            AppUpdater.ReleaseInfo r = new AppUpdater.ReleaseInfo { Version = AppUpdater.ParseVersion(remote), ManifestUrl = "https://x/update.json" };
            Assert.Equal(expected, AppUpdater.IsNewer(r, AppUpdater.ParseVersion(local)));
        }

        [Fact]
        public void IsNewer_needs_a_manifest_or_zip()
        {
            AppUpdater.ReleaseInfo r = new AppUpdater.ReleaseInfo { Version = new Version(9, 0, 0, 0) };
            Assert.False(AppUpdater.IsNewer(r, new Version(1, 0, 0, 0)));
            Assert.False(AppUpdater.IsNewer(null, new Version(1, 0, 0, 0)));
            r.ZipUrl = "https://x/a.zip";
            Assert.True(AppUpdater.IsNewer(r, new Version(1, 0, 0, 0)));
        }

        #endregion

        #region GitHub release

        const string ReleaseJson = @"{
  ""tag_name"": ""2.12.0"",
  ""html_url"": ""https://github.com/JordanChiang/SongMgr/releases/tag/2.12.0"",
  ""draft"": false,
  ""prerelease"": false,
  ""assets"": [
    { ""name"": ""Codec_x86_2.12.0.zip"", ""browser_download_url"": ""https://x/Codec_x86_2.12.0.zip"", ""size"": 12 },
    { ""name"": ""CrazyKtv_SongMgr2.12.0.zip"", ""browser_download_url"": ""https://x/CrazyKtv_SongMgr2.12.0.zip"", ""size"": 123456 },
    { ""name"": ""update.json"", ""browser_download_url"": ""https://x/update.json"", ""size"": 900 }
  ]
}";

        [Fact]
        public void ParseRelease_reads_manifest_and_full_package()
        {
            AppUpdater.ReleaseInfo r = AppUpdater.ParseRelease(ReleaseJson);
            Assert.NotNull(r);
            Assert.Equal(new Version(2, 12, 0, 0), r.Version);
            Assert.Equal("2.12.0", r.Tag);
            Assert.Equal("https://x/update.json", r.ManifestUrl);
            Assert.Equal("https://x/CrazyKtv_SongMgr2.12.0.zip", r.ZipUrl);   // 不是排在前面的 Codec_x86 zip
            Assert.Equal(123456, r.ZipSize);
            Assert.Equal("https://github.com/JordanChiang/SongMgr/releases/tag/2.12.0", r.PageUrl);
        }

        [Fact]
        public void ParseRelease_accepts_old_releases_with_only_a_zip()
        {
            AppUpdater.ReleaseInfo r = AppUpdater.ParseRelease(@"{ ""tag_name"": ""2.11.2"", ""assets"": [
                { ""name"": ""CrazyKtv_SongMgr2.11.2.zip"", ""browser_download_url"": ""https://x/full.zip"", ""size"": 5 } ] }");
            Assert.Null(r.ManifestUrl);
            Assert.Equal("https://x/full.zip", r.ZipUrl);
        }

        [Fact]
        public void ParseRelease_skips_prerelease_and_releases_without_assets()
        {
            Assert.Null(AppUpdater.ParseRelease(ReleaseJson.Replace(@"""prerelease"": false", @"""prerelease"": true")));
            Assert.Null(AppUpdater.ParseRelease(@"{ ""tag_name"": ""2.11.3"", ""assets"": [] }"));
            Assert.Null(AppUpdater.ParseRelease("not json"));
        }

        #endregion

        #region update.json

        [Fact]
        public void ParseManifest_reads_components()
        {
            List<AppUpdater.Component> list = AppUpdater.ParseManifest(@"{ ""version"": ""2.12.0"", ""components"": [
                { ""name"": ""App"", ""desc"": ""加歌程式"", ""url"": ""https://x/App.zip"", ""size"": 10,
                  ""files"": { ""CrazyKTV_SongMgr.exe"": ""AB"", ""SongMgr/Custom.lang"": ""cd"" } },
                { ""name"": ""Tools"", ""url"": ""https://x/Tools.zip"", ""size"": 20, ""files"": { ""Tools/ffmpeg.exe"": ""ef"" } } ] }");

            Assert.Equal(2, list.Count);
            Assert.Equal("加歌程式", list[0].DisplayName);
            Assert.Equal("Tools", list[1].DisplayName);
            Assert.Equal("ab", list[0].Files["CrazyKTV_SongMgr.exe"]);   // 雜湊值統一小寫
            Assert.Equal(20, list[1].Size);
        }

        [Theory]
        [InlineData("not json")]
        [InlineData(@"{ ""components"": [] }")]
        [InlineData(@"{ ""components"": [ { ""name"": ""App"", ""url"": ""https://x"", ""files"": {} } ] }")]
        [InlineData(@"{ ""components"": [ { ""name"": ""App"", ""files"": { ""a"": ""b"" } } ] }")]
        public void ParseManifest_rejects_broken_manifests(string json)
        {
            Assert.Null(AppUpdater.ParseManifest(json));
        }

        [Fact]
        public void NeedsUpdate_only_when_a_file_is_missing_or_different()
        {
            string app = Path.Combine(dir, "app");
            Write(app, "Tools/ffmpeg.exe", "ffmpeg 7.1");
            Write(app, "Tools/ffprobe.exe", "ffprobe 7.1");

            Assert.False(AppUpdater.NeedsUpdate(Comp("Tools", ("Tools/ffmpeg.exe", "ffmpeg 7.1"), ("Tools/ffprobe.exe", "ffprobe 7.1")), app));
            Assert.True(AppUpdater.NeedsUpdate(Comp("Tools", ("Tools/ffmpeg.exe", "ffmpeg 8.0"), ("Tools/ffprobe.exe", "ffprobe 7.1")), app));
            Assert.True(AppUpdater.NeedsUpdate(Comp("Tools", ("Tools/ffmpeg.exe", "ffmpeg 7.1"), ("Tools/ffplay.exe", "ffplay 7.1")), app));
        }

        [Fact]
        public void NeedsUpdate_ignores_user_files_that_already_exist()
        {
            string app = Path.Combine(dir, "app");
            Write(app, "CrazyKTV_SongMgr.exe", "exe");
            Write(app, "SongMgr/CrazySongMgr.mdb", "使用者修改過的歌手資料");

            AppUpdater.Component c = Comp("App", ("CrazyKTV_SongMgr.exe", "exe"), ("SongMgr/CrazySongMgr.mdb", "原始參考資料庫"));
            Assert.False(AppUpdater.NeedsUpdate(c, app));

            File.Delete(AppUpdater.LocalPath(app, "SongMgr/CrazySongMgr.mdb"));
            Assert.True(AppUpdater.NeedsUpdate(c, app));   // 不存在時照裝
        }

        [Fact]
        public void VerifyComponent_rejects_mismatched_files()
        {
            string ext = Path.Combine(dir, "ext");
            Write(ext, "Tools/ffmpeg.exe", "ffmpeg 7.1");
            AppUpdater.VerifyComponent(Comp("Tools", ("Tools/ffmpeg.exe", "ffmpeg 7.1")), ext);
            Assert.Throws<InvalidDataException>(() => AppUpdater.VerifyComponent(Comp("Tools", ("Tools/ffmpeg.exe", "ffmpeg 8.0")), ext));
            Assert.Throws<InvalidDataException>(() => AppUpdater.VerifyComponent(Comp("Tools", ("Tools/ffplay.exe", "x")), ext));
        }

        #endregion

        #region 更新檔整理

        [Fact]
        public void FindPackageRoot_handles_flat_and_nested_zip()
        {
            string flat = Path.Combine(dir, "flat");
            Write(flat, AppUpdater.ExeName, "");
            Assert.Equal(flat, AppUpdater.FindPackageRoot(flat));

            string nested = Path.Combine(dir, "nested");
            string inner = Path.Combine(nested, "CrazyKtv_SongMgr2.11.3");
            Write(inner, AppUpdater.ExeName, "");
            Write(inner, "Tools/ffmpeg.exe", "");
            Assert.Equal(inner, AppUpdater.FindPackageRoot(nested));

            string empty = Path.Combine(dir, "empty");
            Directory.CreateDirectory(empty);
            Assert.Null(AppUpdater.FindPackageRoot(empty));
        }

        [Fact]
        public void RemoveKeptFiles_keeps_user_files_that_already_exist()
        {
            string pkg = Path.Combine(dir, "pkg");
            string app = Path.Combine(dir, "app");
            Write(pkg, "SongMgr/CrazySongMgr.mdb", "new");
            Write(pkg, "SongMgr/Custom.lang", "new");
            Write(pkg, "SongMgr/CrazySongEmpty.mdb", "new");
            Write(app, "SongMgr/CrazySongMgr.mdb", "user");

            AppUpdater.RemoveKeptFiles(pkg, app);

            Assert.False(File.Exists(AppUpdater.LocalPath(pkg, "SongMgr/CrazySongMgr.mdb")));   // 使用者已有, 不覆蓋
            Assert.True(File.Exists(AppUpdater.LocalPath(pkg, "SongMgr/Custom.lang")));         // 使用者沒有, 照裝
            Assert.True(File.Exists(AppUpdater.LocalPath(pkg, "SongMgr/CrazySongEmpty.mdb")));  // 非使用者資料, 照常更新
        }

        [Fact]
        public void Script_restarts_with_the_updated_argument()
        {
            Assert.Contains("start \"\" \"%DST%\\" + AppUpdater.ExeName + "\" " + AppUpdater.RestartArg + "\r\n", AppUpdater.ScriptText);
        }

        #endregion

        #region 下載整段流程 (本機 HTTP)

        [Fact]
        public void Only_changed_components_are_downloaded_and_verified()
        {
            // 新版: 主程式與 Codec 變了, Tools 沒變
            string pub = Path.Combine(dir, "publish");
            Write(pub, "CrazyKTV_SongMgr.exe", "exe 2.12.0");
            Write(pub, "SongMgr/CrazySongMgr.mdb", "原始參考資料庫");
            Write(pub, "Codec_x86/LAVVideo.ax", "lav 0.80");
            Write(pub, "Tools/ffmpeg.exe", "ffmpeg 7.1");

            // 使用者電腦: 舊版主程式、舊 Codec、已是最新的 Tools、自己改過的參考資料庫
            string app = Path.Combine(dir, "app");
            Write(app, "CrazyKTV_SongMgr.exe", "exe 2.11.2");
            Write(app, "SongMgr/CrazySongMgr.mdb", "使用者的歌手資料");
            Write(app, "Codec_x86/LAVVideo.ax", "lav 0.79");
            Write(app, "Tools/ffmpeg.exe", "ffmpeg 7.1");

            using (TinyHttpServer server = new TinyHttpServer())
            {
                string Pack(string name, params string[] rels)
                {
                    string zip = Path.Combine(dir, name + ".zip");
                    using (ZipArchive z = ZipFile.Open(zip, ZipArchiveMode.Create))
                        foreach (string rel in rels) z.CreateEntryFromFile(AppUpdater.LocalPath(pub, rel), rel);
                    server.Files["/" + name + ".zip"] = File.ReadAllBytes(zip);
                    string files = string.Join(",", rels.Select(r => "\"" + r + "\": \"" + AppUpdater.FileHash(AppUpdater.LocalPath(pub, r)) + "\""));
                    return "{ \"name\": \"" + name + "\", \"url\": \"" + server.Url + name + ".zip\", \"size\": " + new FileInfo(zip).Length + ", \"files\": { " + files + " } }";
                }

                string manifest = "{ \"version\": \"2.12.0\", \"components\": [ " +
                    Pack("App", "CrazyKTV_SongMgr.exe", "SongMgr/CrazySongMgr.mdb") + ", " +
                    Pack("Codec_x86", "Codec_x86/LAVVideo.ax") + ", " +
                    Pack("Tools", "Tools/ffmpeg.exe") + " ] }";
                server.Files["/update.json"] = Encoding.UTF8.GetBytes(manifest);

                AppUpdater.ReleaseInfo release = new AppUpdater.ReleaseInfo { Version = new Version(2, 12, 0, 0), Tag = "2.12.0", ManifestUrl = server.Url + "update.json" };
                AppUpdater.UpdatePlan plan = AppUpdater.GetUpdatePlan(release, app);

                Assert.True(plan.FromManifest);
                Assert.Equal(new[] { "App", "Codec_x86" }, plan.Components.Select(c => c.Name));

                string root = AppUpdater.Prepare(plan, app, null, CancellationToken.None);
                try
                {
                    Assert.Equal("exe 2.12.0", File.ReadAllText(AppUpdater.LocalPath(root, "CrazyKTV_SongMgr.exe")));
                    Assert.Equal("lav 0.80", File.ReadAllText(AppUpdater.LocalPath(root, "Codec_x86/LAVVideo.ax")));
                    Assert.False(File.Exists(AppUpdater.LocalPath(root, "Tools/ffmpeg.exe")));            // 沒下載
                    Assert.False(File.Exists(AppUpdater.LocalPath(root, "SongMgr/CrazySongMgr.mdb")));    // 保留使用者的
                    Assert.DoesNotContain("/Tools.zip", server.Requests);
                }
                finally
                {
                    Directory.Delete(AppUpdater.StageDir, true);
                }
            }
        }

        /// <summary>只回應 GET 的極簡 HTTP 服務, 不需要系統管理員權限.</summary>
        sealed class TinyHttpServer : IDisposable
        {
            readonly TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
            public readonly Dictionary<string, byte[]> Files = new Dictionary<string, byte[]>();
            public readonly List<string> Requests = new List<string>();
            public string Url { get; }

            public TinyHttpServer()
            {
                listener.Start();
                Url = "http://127.0.0.1:" + ((IPEndPoint)listener.LocalEndpoint).Port + "/";
                Task.Run(Loop);
            }

            async Task Loop()
            {
                while (true)
                {
                    TcpClient client;
                    try { client = await listener.AcceptTcpClientAsync(); } catch { return; }
                    _ = Task.Run(() => Serve(client));
                }
            }

            void Serve(TcpClient client)
            {
                using (client)
                using (NetworkStream ns = client.GetStream())
                {
                    StreamReader reader = new StreamReader(ns, Encoding.ASCII);
                    string first = reader.ReadLine() ?? "";
                    while (!string.IsNullOrEmpty(reader.ReadLine())) { }
                    string path = first.Split(' ').Length > 1 ? first.Split(' ')[1] : "/";
                    lock (Requests) Requests.Add(path);

                    byte[] body;
                    string status = Files.TryGetValue(path, out body) ? "200 OK" : "404 Not Found";
                    body = body ?? new byte[0];
                    byte[] head = Encoding.ASCII.GetBytes("HTTP/1.1 " + status + "\r\nContent-Length: " + body.Length + "\r\nConnection: close\r\n\r\n");
                    ns.Write(head, 0, head.Length);
                    ns.Write(body, 0, body.Length);
                }
            }

            public void Dispose() => listener.Stop();
        }

        #endregion
    }
}

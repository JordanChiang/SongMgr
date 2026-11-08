using System;
using System.IO;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Update
{
    /// <summary>程式自動更新: 版本比對、GitHub release 解析與更新檔整理.</summary>
    public class AppUpdaterTests : IDisposable
    {
        readonly string dir = Path.Combine(Path.GetTempPath(), "AppUpdaterTests_" + Guid.NewGuid().ToString("N"));

        public AppUpdaterTests() => Directory.CreateDirectory(dir);

        public void Dispose()
        {
            try { Directory.Delete(dir, true); } catch { }
        }

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
            AppUpdater.ReleaseInfo r = new AppUpdater.ReleaseInfo { Version = AppUpdater.ParseVersion(remote), ZipUrl = "https://x/a.zip" };
            Assert.Equal(expected, AppUpdater.IsNewer(r, AppUpdater.ParseVersion(local)));
        }

        [Fact]
        public void IsNewer_needs_a_zip()
        {
            AppUpdater.ReleaseInfo r = new AppUpdater.ReleaseInfo { Version = new Version(9, 0, 0, 0) };
            Assert.False(AppUpdater.IsNewer(r, new Version(1, 0, 0, 0)));
            Assert.False(AppUpdater.IsNewer(null, new Version(1, 0, 0, 0)));
        }

        const string ReleaseJson = @"{
  ""tag_name"": ""2.11.3"",
  ""html_url"": ""https://github.com/JordanChiang/SongMgr/releases/tag/2.11.3"",
  ""draft"": false,
  ""prerelease"": false,
  ""assets"": [
    { ""name"": ""notes.txt"", ""browser_download_url"": ""https://x/notes.txt"", ""size"": 10 },
    { ""name"": ""CrazyKtv_SongMgr2.11.3.zip"", ""browser_download_url"": ""https://x/CrazyKtv_SongMgr2.11.3.zip"", ""size"": 123456 }
  ]
}";

        [Fact]
        public void ParseRelease_reads_tag_and_zip_asset()
        {
            AppUpdater.ReleaseInfo r = AppUpdater.ParseRelease(ReleaseJson);
            Assert.NotNull(r);
            Assert.Equal(new Version(2, 11, 3, 0), r.Version);
            Assert.Equal("2.11.3", r.Tag);
            Assert.Equal("https://x/CrazyKtv_SongMgr2.11.3.zip", r.ZipUrl);
            Assert.Equal(123456, r.ZipSize);
            Assert.Equal("https://github.com/JordanChiang/SongMgr/releases/tag/2.11.3", r.PageUrl);
        }

        [Fact]
        public void ParseRelease_skips_prerelease_and_missing_zip()
        {
            Assert.Null(AppUpdater.ParseRelease(ReleaseJson.Replace(@"""prerelease"": false", @"""prerelease"": true")));
            Assert.Null(AppUpdater.ParseRelease(@"{ ""tag_name"": ""2.11.3"", ""assets"": [] }"));
            Assert.Null(AppUpdater.ParseRelease("not json"));
        }

        [Fact]
        public void FindPackageRoot_handles_flat_and_nested_zip()
        {
            string flat = Path.Combine(dir, "flat");
            Directory.CreateDirectory(flat);
            File.WriteAllText(Path.Combine(flat, AppUpdater.ExeName), "");
            Assert.Equal(flat, AppUpdater.FindPackageRoot(flat));

            string nested = Path.Combine(dir, "nested");
            string inner = Path.Combine(nested, "CrazyKtv_SongMgr2.11.3");
            Directory.CreateDirectory(Path.Combine(inner, "Tools"));
            File.WriteAllText(Path.Combine(inner, AppUpdater.ExeName), "");
            Assert.Equal(inner, AppUpdater.FindPackageRoot(nested));

            string empty = Path.Combine(dir, "empty");
            Directory.CreateDirectory(empty);
            Assert.Null(AppUpdater.FindPackageRoot(empty));
        }

        [Fact]
        public void Script_restarts_with_the_updated_argument()
        {
            Assert.Contains("start \"\" \"%DST%\\" + AppUpdater.ExeName + "\" " + AppUpdater.RestartArg + "\r\n", AppUpdater.ScriptText);
        }

        [Fact]
        public void RemoveKeptFiles_keeps_user_files_that_already_exist()
        {
            string pkg = Path.Combine(dir, "pkg");
            string app = Path.Combine(dir, "app");
            Directory.CreateDirectory(Path.Combine(pkg, "SongMgr"));
            Directory.CreateDirectory(Path.Combine(app, "SongMgr"));

            File.WriteAllText(Path.Combine(pkg, @"SongMgr\CrazySongMgr.mdb"), "new");
            File.WriteAllText(Path.Combine(pkg, @"SongMgr\Custom.lang"), "new");
            File.WriteAllText(Path.Combine(pkg, @"SongMgr\CrazySongEmpty.mdb"), "new");
            File.WriteAllText(Path.Combine(app, @"SongMgr\CrazySongMgr.mdb"), "user");

            AppUpdater.RemoveKeptFiles(pkg, app);

            Assert.False(File.Exists(Path.Combine(pkg, @"SongMgr\CrazySongMgr.mdb")));   // 使用者已有, 不覆蓋
            Assert.True(File.Exists(Path.Combine(pkg, @"SongMgr\Custom.lang")));         // 使用者沒有, 照裝
            Assert.True(File.Exists(Path.Combine(pkg, @"SongMgr\CrazySongEmpty.mdb")));  // 非使用者資料, 照常更新
        }
    }
}

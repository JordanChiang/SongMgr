using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Settings
{
    /// <summary>
    /// 歌庫設定 used to be read as a 56-slot list, one XML parse per key, and written key by key.
    /// The settings file module must read and write exactly what that code did.
    /// </summary>
    public class SettingsFileTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string dir = TestEnv.NewTempDir("settings");
        readonly Action<string> savedReporter = SettingsFile.ReportError;
        readonly List<string> errors = new List<string>();

        public SettingsFileTests()
        {
            SettingsFile.ReportError = errors.Add;
        }

        public void Dispose()
        {
            SettingsFile.ReportError = savedReporter;
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(dir);
        }

        [Fact]
        public void A_broken_file_is_reported_once_and_reads_as_empty()
        {
            string path = Path.Combine(dir, "broken.cfg");
            File.WriteAllText(path, "<Configeruation><setting");

            SettingsFile cfg = SettingsFile.Load(path);

            Assert.Equal("", cfg["SongMgrDestFolder"]);
            Assert.Equal("", cfg["SongMgrSongAddMode"]);
            Assert.Equal(new[] { "【broken.cfg】設定檔內容有錯誤,請刪除後再執行。" }, errors);
        }

        /// <summary>The keys MainForm_Load wrote into a fresh CrazyKTV_SongMgr.cfg (master @ 21fb779), in order.</summary>
        static readonly string[] LegacyDefaultKeys =
        {
            "CrazyktvDatabaseFile", "SongMgrSupportFormat", "SongMgrDestFolder", "SongMgrSongAddMode", "SongMgrChorusMerge",
            "SongMgrMaxDigitCode", "SongMgrLangCode", "SongMgrSongType", "SongMgrSongInfoSeparate", "SongMgrChorusSeparate",
            "SongAddDefaultSongLang", "SongAddDefaultSingerType", "SongAddDefaultSongTrack", "SongAddDefaultSongType",
            "SongAddDefaultSongVolume", "SongAddSpecialStr", "SongMgrFolderStructure", "SongMgrFileStructure", "SongMgrSongTrackMode",
            "SongAddSongIdentificationMode", "SongAddDupSongMode", "CrazyktvSongLangStr", "CrazyktvSongLangKeyWord",
            "SongMgrBackupRemoveSong", "SongAddEngSongNameFormat", "MainCfgAlwaysOnTop",
            "MainCfgHideSongDBConverterTabPage", "MainCfgHideSongAddResultTabPage", "MainCfgHideSongLogTabPage",
            "MainCfgBackupRemoveSongDays", "MainCfgUIScale", "SongMaintenanceReplayGainVolume",
            "SongAddEnableVolumeDetect", "SongAddUseCustomSongID", "MainCfgUIFont", "MainCfgEnableUIScale",
            "SongMgrCustomSingerTypeStructure", "MainCfgEnableAutoUpdate", "SongAddEnableConvToTC", "SongMgrEnableMonitorFolders",
            "SongMgrMonitorFolders", "SingerMgrSyncSongSinger", "SongMgrSingerGroup", "MainCfgHideApplyCashboxIdButton",
            "MainCfgUICustomScale", "MainCfgBackupDB", "MainCfgBackupDBPath", "MainCfgPlayerCore", "MainCfgPlayerOutput",
            "MainCfgPlayerEnableAudioCompressor", "MainCfgPlayerEnableAudioProcessor", "MainCfgPlayerDefaultVolume",
            "MainCfgAutoSongFileCount",
        };

        [Fact]
        public void Defaults_cover_the_same_keys_in_the_same_order()
        {
            Assert.Equal(LegacyDefaultKeys, SongMgrSettings.Defaults().Select(d => d.Name));
        }

        [Fact]
        public void A_fresh_file_is_byte_for_byte_what_the_key_by_key_writes_produced()
        {
            Global.SongMgrDestFolder = @"D:\KTV\歌庫";
            Global.SongMgrSongType = "演唱會,MTV";
            Global.SongMgrMonitorFoldersList = new List<string> { @"E:\監視", "", "", "", "" };

            string legacy = Path.Combine(dir, "legacy.cfg");
            CommonFunc.CreateConfigXmlFile(legacy);
            foreach (var d in SongMgrSettings.Defaults()) CommonFunc.SaveConfigXmlFile(legacy, d.Name, d.Value);

            string module = Path.Combine(dir, "module.cfg");
            SettingsFile.Create(module, SongMgrSettings.Defaults());

            Assert.Equal(File.ReadAllBytes(legacy), File.ReadAllBytes(module));
        }

        [Fact]
        public void Reading_by_name_matches_the_old_per_key_reads()
        {
            string path = Path.Combine(dir, "user.cfg");
            CommonFunc.CreateConfigXmlFile(path);
            CommonFunc.SaveConfigXmlFile(path, "SongMgrDestFolder", @"D:\KTV\歌庫");
            CommonFunc.SaveConfigXmlFile(path, "SongMgrSongAddMode", "1");
            CommonFunc.SaveConfigXmlFile(path, "CrazyktvSongLangKeyWord", "國語,國|台語,台");
            CommonFunc.SaveConfigXmlFile(path, "Empty", "");
            // A file edited by hand can repeat a key; the old reader kept the last one.
            File.WriteAllText(path, File.ReadAllText(path).Replace("</Configeruation>", "  <setting Name=\"SongMgrSongAddMode\"><Value>4</Value></setting>\n</Configeruation>"), System.Text.Encoding.Unicode);

            SettingsFile cfg = SettingsFile.Load(path);
            foreach (string key in LegacyDefaultKeys.Concat(new[] { "SongMgrDestFolder", "SongMgrSongAddMode", "CrazyktvSongLangKeyWord", "Empty", "NotThere" }))
            {
                // the pre-refactor reader, verbatim
                string legacy = "";
                var root = System.Xml.Linq.XElement.Load(path);
                foreach (var node in root.Elements("setting").Where(n => (string)n.Attribute("Name") == key)) legacy = node.Value;

                Assert.Equal(legacy, cfg[key]);
                Assert.Equal(legacy, CommonFunc.LoadConfigXmlFile(path, key, null, false));
            }
            Assert.Equal("4", cfg["SongMgrSongAddMode"]);
        }

        [Theory]
        [InlineData("1", "Move", true)]
        [InlineData("2", "Copy", true)]
        [InlineData("3", "InPlace", false)]
        [InlineData("4", "Monitor", false)]
        public void Add_mode_reads_the_stored_code_and_decides_whether_files_move(string stored, string mode, bool moves)
        {
            SongMgrSettings.SongMgrSongAddMode.Set(stored);

            Assert.Equal(mode, Global.SongMgrSongAddMode.ToString());
            Assert.Equal(stored, SongMgrSettings.SongMgrSongAddMode.Get());
            Assert.Equal(moves, LibraryPlacement.MovesFiles);
        }

        [Theory]
        [InlineData("01")]
        [InlineData("0")]
        [InlineData("5")]
        [InlineData("Move")]
        public void Add_mode_other_than_1_to_4_keeps_its_value(string stored)
        {
            Global.SongMgrSongAddMode = SongAddMode.Copy;

            SongMgrSettings.SongMgrSongAddMode.Set(stored);

            Assert.Equal(SongAddMode.Copy, Global.SongMgrSongAddMode);
        }

        [Fact]
        public void Add_mode_defaults_to_in_place()
        {
            Assert.Equal(SongAddMode.InPlace, SongMgrSettings.SongMgrSongAddMode.Default);
            Assert.Contains(("SongMgrSongAddMode", "3"), SongMgrSettings.Defaults());
        }
    }
}

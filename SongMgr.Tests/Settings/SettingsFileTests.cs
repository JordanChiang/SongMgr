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
            "SongMgrBackupRemoveSong", "SongAddEngSongNameFormat", "MainCfgAlwaysOnTop", "CrazyktvSongDBVer",
            "MainCfgHideSongDBConverterTabPage", "MainCfgHideSongAddResultTabPage", "MainCfgHideSongLogTabPage",
            "MainCfgBackupRemoveSongDays", "DBVerEnableDBVerUpdate", "MainCfgUIScale", "SongMaintenanceReplayGainVolume",
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
        [InlineData("1", 1, true)]
        [InlineData("2", 2, true)]
        [InlineData("3", 3, false)]
        [InlineData("4", 4, false)]
        [InlineData("", 0, true)]
        [InlineData("01", 0, true)]
        public void Add_mode_names_the_stored_value_and_decides_whether_files_move(string stored, int mode, bool moves)
        {
            Global.SongMgrSongAddMode = stored;
            Assert.Equal((SongAddMode)mode, Global.AddMode);
            // the old check, verbatim
            Assert.Equal(Global.SongMgrSongAddMode != "3" && Global.SongMgrSongAddMode != "4", LibraryPlacement.MovesFiles);
            Assert.Equal(moves, LibraryPlacement.MovesFiles);
        }
    }
}

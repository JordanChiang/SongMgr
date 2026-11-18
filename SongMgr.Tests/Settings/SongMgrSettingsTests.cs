using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Settings
{
    /// <summary>
    /// 歌庫設定 used to be loaded by a hand-written if per key, saved one file write per key from lists in each page,
    /// and created from a third list. The registry <see cref="SongMgrSettings.Entries"/> drives all three;
    /// these tests pin it to what the old code read and wrote.
    /// </summary>
    public class SongMgrSettingsTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string dir = TestEnv.NewTempDir("songmgrsettings");
        readonly string path;

        public SongMgrSettingsTests()
        {
            path = Path.Combine(dir, "CrazyKTV_SongMgr.cfg");
            Global.SongMgrCfgFile = path;
        }

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(dir);
        }

        /// <summary>Puts Global back to its initial values, keeping the test's settings file.</summary>
        void ResetGlobals()
        {
            TestEnv.Restore(saved);
            Global.SongMgrCfgFile = path;
        }

        static readonly SettingsGroup[] SavedGroups = { SettingsGroup.MainCfg, SettingsGroup.SongMgrCfg, SettingsGroup.SongAdd, SettingsGroup.Direct };

        static IEnumerable<Setting> In(SettingsGroup group) => SongMgrSettings.Entries.Where(s => s.Group == group);

        /// <summary>What each page's 儲存 button wrote, in its order (master @ 11624ff).</summary>
        static readonly string[] LegacyMainCfgSave =
        {
            "MainCfgAlwaysOnTop", "MainCfgEnableAutoUpdate", "MainCfgAutoSongFileCount", "MainCfgHideSongDBConverterTabPage",
            "MainCfgHideSongAddResultTabPage", "MainCfgHideSongLogTabPage", "MainCfgHideApplyCashboxIdButton", "MainCfgBackupRemoveSongDays",
            "MainCfgUIScale", "MainCfgUIFont", "MainCfgEnableUIScale", "MainCfgUICustomScale", "MainCfgBackupDB", "MainCfgBackupDBPath",
            "MainCfgPlayerCore", "MainCfgPlayerOutput", "MainCfgPlayerEnableAudioCompressor", "MainCfgPlayerEnableAudioProcessor",
            "MainCfgPlayerDefaultVolume",
        };

        static readonly string[] LegacySongMgrCfgSave =
        {
            "CrazyktvDatabaseFile", "SongMgrSupportFormat", "SongMgrDestFolder", "SongMgrSongAddMode", "SongMgrChorusMerge",
            "SongMgrMaxDigitCode", "SongMgrLangCode", "SongMgrSongType", "SongMgrSongInfoSeparate", "SongMgrChorusSeparate",
            "SongMgrFolderStructure", "SongMgrFileStructure", "SongMgrSongTrackMode", "SongMgrBackupRemoveSong",
            "SongMgrCustomSingerTypeStructure", "SongMgrEnableMonitorFolders", "SongMgrMonitorFolders", "SongMgrSingerGroup",
        };

        static readonly string[] LegacySongAddSave =
        {
            "SongAddDefaultSongLang", "SongAddDefaultSingerType", "SongAddDefaultSongTrack", "SongAddDefaultSongType",
            "SongAddDefaultSongVolume", "SongAddSpecialStr", "SongAddSongIdentificationMode", "SongAddDupSongMode",
            "SongAddEngSongNameFormat", "SongAddUseCustomSongID", "SongAddEnableConvToTC", "SongAddEnableVolumeDetect",
        };

        /// <summary>Saved one at a time where they change, by no button.</summary>
        static readonly string[] LegacyDirectSave =
        {
            "CrazyktvSongLangStr", "CrazyktvSongLangKeyWord", "SongMaintenanceReplayGainVolume", "SingerMgrSyncSongSinger",
        };

        [Fact]
        public void Each_group_holds_what_its_button_saved()
        {
            Assert.Equal(LegacyMainCfgSave.OrderBy(n => n), In(SettingsGroup.MainCfg).Select(s => s.Name).OrderBy(n => n));
            Assert.Equal(LegacySongMgrCfgSave.OrderBy(n => n), In(SettingsGroup.SongMgrCfg).Select(s => s.Name).OrderBy(n => n));
            Assert.Equal(LegacySongAddSave.OrderBy(n => n), In(SettingsGroup.SongAdd).Select(s => s.Name).OrderBy(n => n));
            Assert.Equal(LegacyDirectSave.OrderBy(n => n), In(SettingsGroup.Direct).Select(s => s.Name).OrderBy(n => n));
            Assert.Equal(new[] { "SingerMgrDefaultSingerDataTable" }, In(SettingsGroup.LoadOnly).Select(s => s.Name));
        }

        [Fact]
        public void Names_are_unique()
        {
            Assert.Equal(SongMgrSettings.Entries.Count, SongMgrSettings.Entries.Select(s => s.Name).Distinct().Count());
        }

        [Fact]
        public void Entries_are_every_setting_field_in_declaration_order()
        {
            var fields = typeof(SongMgrSettings).GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
                .Where(f => typeof(Setting).IsAssignableFrom(f.FieldType))
                .Select(f => (Setting)f.GetValue(null)).ToList();

            Assert.Equal(fields.OrderBy(s => s.Name), SongMgrSettings.Entries.OrderBy(s => s.Name));
            Assert.All(typeof(SongMgrSettings).GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
                .Where(f => typeof(Setting).IsAssignableFrom(f.FieldType)),
                f => Assert.Equal(f.Name, ((Setting)f.GetValue(null)).Name));
        }

        [Fact]
        public void Global_is_a_view_of_the_registry()
        {
            Global.SongMgrDestFolder = @"D:\KTV\歌庫";
            Assert.Equal(@"D:\KTV\歌庫", SongMgrSettings.SongMgrDestFolder.Value);

            SongMgrSettings.MainCfgUIFont.Value = "Tahoma";
            Assert.Equal("Tahoma", Global.MainCfgUIFont);

            Global.SongMgrMonitorFoldersList[0] = @"E:\監視";
            Assert.Equal(@"E:\監視,,,,", SongMgrSettings.SongMgrMonitorFolders.Get());
        }

        [Fact]
        public void A_snapshot_restores_values_and_list_contents_more_than_once()
        {
            var snapshot = SongMgrSettings.Snapshot();
            for (int i = 0; i < 2; i++)
            {
                Global.SongMgrDestFolder = "改過";
                Global.CrazyktvSongLangList.Add("新語言");

                SongMgrSettings.Restore(snapshot);

                Assert.Equal("", Global.SongMgrDestFolder);
                Assert.Equal(SongMgrSettings.CrazyktvSongLangStr.Default, Global.CrazyktvSongLangList);
            }
        }

        [Theory]
        [InlineData("True", true)]
        [InlineData("False", false)]
        [InlineData("true", true)]
        [InlineData("FALSE", false)]
        public void A_bool_reads_true_or_false_in_any_case(string text, bool expected)
        {
            SongMgrSettings.SongMgrChorusMerge.Value = !expected;

            SongMgrSettings.SongMgrChorusMerge.Set(text);

            Assert.Equal(expected, Global.SongMgrChorusMerge);
        }

        [Theory]
        [InlineData("yes")]
        [InlineData("1")]
        [InlineData("True-ish")]
        public void A_bool_that_does_not_parse_keeps_its_value(string text)
        {
            SettingsFile.Create(path, new[] { ("SongMgrChorusMerge", text), ("MainCfgEnableAutoUpdate", text) });

            SongMgrSettings.Load(SettingsFile.Load(path));

            Assert.False(Global.SongMgrChorusMerge);
            Assert.True(Global.MainCfgEnableAutoUpdate);
        }

        [Fact]
        public void A_bool_is_written_as_True_or_False()
        {
            SettingsFile.Create(path, SongMgrSettings.Defaults());
            Global.SongMgrChorusMerge = true;
            Global.SongMgrSongTrackMode = false;

            SongMgrSettings.Save(SettingsGroup.SongMgrCfg);

            SettingsFile cfg = SettingsFile.Load(path);
            Assert.Equal("True", cfg["SongMgrChorusMerge"]);
            Assert.Equal("False", cfg["SongMgrSongTrackMode"]);
            Assert.Equal("True", cfg["MainCfgEnableAutoUpdate"]);
            Assert.Equal("False", cfg["MainCfgAlwaysOnTop"]);
        }

        [Fact]
        public void A_list_default_is_a_new_copy_each_time()
        {
            SongMgrSettings.SongMgrMonitorFolders.Default.Add("不該留下");

            Assert.Equal(new[] { "", "", "", "", "" }, SongMgrSettings.SongMgrMonitorFolders.Default);
        }

        [Fact]
        public void Defaults_are_every_saved_setting_with_its_current_value()
        {
            Global.SongMgrDestFolder = @"D:\KTV\歌庫";
            Global.CrazyktvSongLangKeyWordList = new List<string> { "國語,國", "台語,台" };

            var expected = SongMgrSettings.Entries.Where(s => s.Group != SettingsGroup.LoadOnly).Select(s => (s.Name, s.Get()));

            Assert.Equal(expected, SongMgrSettings.Defaults());
            Assert.Contains(("SongMgrDestFolder", @"D:\KTV\歌庫"), SongMgrSettings.Defaults());
            Assert.Contains(("CrazyktvSongLangKeyWord", "國語,國|台語,台"), SongMgrSettings.Defaults());
        }

        [Fact]
        public void Every_saved_setting_round_trips_through_the_file()
        {
            SettingsFile.Create(path, SongMgrSettings.Defaults());
            string loadOnlyDefault = Global.SingerMgrDefaultSingerDataTable;
            foreach (Setting s in SongMgrSettings.Entries) s.Set(Changed(s, "-值"));

            foreach (SettingsGroup group in SavedGroups) SongMgrSettings.Save(group);
            ResetGlobals();
            SongMgrSettings.Load(SettingsFile.Load(path));

            foreach (Setting s in SongMgrSettings.Entries.Where(s => s.Group != SettingsGroup.LoadOnly))
            {
                Assert.Equal(Changed(s, "-值"), s.Get());
            }
            Assert.Equal(loadOnlyDefault, Global.SingerMgrDefaultSingerDataTable);
            Assert.Equal("", SettingsFile.Load(path)["SingerMgrDefaultSingerDataTable"]);
        }

        [Fact]
        public void Lists_are_split_and_joined_on_their_separator()
        {
            SettingsFile.Create(path, SongMgrSettings.Defaults());
            Global.CrazyktvSongLangList = new List<string> { "國語", "台語" };
            Global.CrazyktvSongLangKeyWordList = new List<string> { "國語,國", "台語,台" };
            Global.SongMgrMonitorFoldersList = new List<string> { @"E:\監視", "", "", "", "" };

            SongMgrSettings.Save("CrazyktvSongLangStr", "CrazyktvSongLangKeyWord");
            SongMgrSettings.Save(SettingsGroup.SongMgrCfg);
            SettingsFile cfg = SettingsFile.Load(path);
            ResetGlobals();
            SongMgrSettings.Load(cfg);

            Assert.Equal("國語,台語", cfg["CrazyktvSongLangStr"]);
            Assert.Equal("國語,國|台語,台", cfg["CrazyktvSongLangKeyWord"]);
            Assert.Equal(@"E:\監視,,,,", cfg["SongMgrMonitorFolders"]);
            Assert.Equal(new[] { "國語", "台語" }, Global.CrazyktvSongLangList);
            Assert.Equal(new[] { "國語,國", "台語,台" }, Global.CrazyktvSongLangKeyWordList);
            Assert.Equal(new[] { @"E:\監視", "", "", "", "" }, Global.SongMgrMonitorFoldersList);
        }

        [Fact]
        public void An_old_file_loads_what_it_has_and_keeps_the_defaults_for_the_rest()
        {
            var defaults = SongMgrSettings.Entries.ToDictionary(s => s.Name, s => s.Get());
            CommonFunc.CreateConfigXmlFile(path);
            CommonFunc.SaveConfigXmlFile(path, "SongMgrDestFolder", @"D:\KTV\歌庫");
            CommonFunc.SaveConfigXmlFile(path, "SongMgrSongAddMode", "");
            CommonFunc.SaveConfigXmlFile(path, "CrazyktvSongLangIDStr", "國語,國*台語,台");

            SongMgrSettings.Load(SettingsFile.Load(path));

            Assert.Equal(@"D:\KTV\歌庫", Global.SongMgrDestFolder);
            foreach (Setting s in SongMgrSettings.Entries.Where(s => s.Name != "SongMgrDestFolder"))
            {
                Assert.Equal(defaults[s.Name], s.Get());
            }
        }

        static void LegacySave(string file, IEnumerable<(string Name, string Value)> settings)
        {
            foreach (var (name, value) in settings) CommonFunc.SaveConfigXmlFile(file, name, value);
        }

        IEnumerable<(string Name, string Value)> Current(IEnumerable<string> names)
        {
            var byName = SongMgrSettings.Entries.ToDictionary(s => s.Name);
            return names.Select(n => (n, byName[n].Get())).ToList();
        }

        void ChangeEverySetting()
        {
            foreach (Setting s in SongMgrSettings.Entries) s.Set(Changed(s, "-新"));
        }

        /// <summary>File text for <paramref name="s"/> that differs from its default and parses: the other bool, or tagged text.</summary>
        static string Changed(Setting s, string tag)
        {
            if (s is Setting<bool> b) return b.Default ? "False" : "True";
            return s.Name + tag;
        }

        /// <summary>A file with every setting, a hand-added unknown key, the legacy key, and a repeated key.</summary>
        void WriteUserFile(string file)
        {
            SettingsFile.Create(file, SongMgrSettings.Defaults());
            CommonFunc.SaveConfigXmlFile(file, "CrazyktvSongLangIDStr", "國語,國*台語,台");
            CommonFunc.SaveConfigXmlFile(file, "SomeoneElsesKey", "留著");
            File.WriteAllText(file, File.ReadAllText(file).Replace("</Configeruation>", "  <setting Name=\"MainCfgUIFont\"><Value>Tahoma</Value></setting>\n</Configeruation>"), System.Text.Encoding.Unicode);
        }

        [Fact]
        public void A_group_save_is_byte_for_byte_what_the_key_by_key_writes_produced()
        {
            foreach (var (group, legacyOrder) in new[]
            {
                (SettingsGroup.MainCfg, LegacyMainCfgSave),
                (SettingsGroup.SongMgrCfg, LegacySongMgrCfgSave),
                (SettingsGroup.SongAdd, LegacySongAddSave),
            })
            {
                ResetGlobals();
                string legacy = Path.Combine(dir, group + ".legacy.cfg");
                WriteUserFile(legacy);
                File.Copy(legacy, path, true);
                ChangeEverySetting();

                LegacySave(legacy, Current(legacyOrder));
                SongMgrSettings.Save(group);

                Assert.Equal(File.ReadAllBytes(legacy), File.ReadAllBytes(path));
            }
        }

        [Fact]
        public void A_direct_save_is_byte_for_byte_what_the_key_by_key_writes_produced()
        {
            string legacy = Path.Combine(dir, "direct.legacy.cfg");
            WriteUserFile(legacy);
            File.Copy(legacy, path, true);
            ChangeEverySetting();

            LegacySave(legacy, Current(new[] { "SongMgrMaxDigitCode", "SongMgrLangCode" }));
            LegacySave(legacy, new[] { ("SongAddSpecialStr", "過濾後") });
            SongMgrSettings.Save("SongMgrMaxDigitCode", "SongMgrLangCode");
            SongMgrSettings.SaveValue("SongAddSpecialStr", "過濾後");

            Assert.Equal(File.ReadAllBytes(legacy), File.ReadAllBytes(path));
        }

        [Fact]
        public void Saving_keeps_unknown_and_legacy_keys_and_updates_every_copy_of_a_repeated_key()
        {
            WriteUserFile(path);
            Global.MainCfgUIFont = "微軟正黑體";

            SongMgrSettings.Save(SettingsGroup.MainCfg);

            var settings = XElement.Load(path).Elements("setting").ToList();
            Assert.Equal("留著", settings.Single(n => (string)n.Attribute("Name") == "SomeoneElsesKey").Value);
            Assert.Equal("國語,國*台語,台", settings.Single(n => (string)n.Attribute("Name") == "CrazyktvSongLangIDStr").Value);
            Assert.Equal(new[] { "微軟正黑體", "微軟正黑體" }, settings.Where(n => (string)n.Attribute("Name") == "MainCfgUIFont").Select(n => n.Value));
        }

        [Fact]
        public void A_file_missing_settings_gets_them_added_at_the_end_in_registry_order()
        {
            CommonFunc.CreateConfigXmlFile(path);
            CommonFunc.SaveConfigXmlFile(path, "SomeoneElsesKey", "留著");
            CommonFunc.SaveConfigXmlFile(path, "MainCfgUIFont", "Tahoma");

            SongMgrSettings.Save(SettingsGroup.MainCfg);

            var names = XElement.Load(path).Elements("setting").Select(n => (string)n.Attribute("Name")).ToList();
            var expected = new[] { "SomeoneElsesKey", "MainCfgUIFont" }
                .Concat(In(SettingsGroup.MainCfg).Select(s => s.Name).Where(n => n != "MainCfgUIFont"));
            Assert.Equal(expected, names);
            Assert.Equal(Global.MainCfgUIFont, SettingsFile.Load(path)["MainCfgUIFont"]);
        }

        [Fact]
        public void An_override_is_written_without_changing_the_global()
        {
            SettingsFile.Create(path, SongMgrSettings.Defaults());
            Global.SongAddSpecialStr = "ft.|feat.|自訂";

            SongMgrSettings.Save(SettingsGroup.SongAdd, ("SongAddSpecialStr", "自訂"));

            Assert.Equal("自訂", SettingsFile.Load(path)["SongAddSpecialStr"]);
            Assert.Equal("ft.|feat.|自訂", Global.SongAddSpecialStr);
        }

        [Fact]
        public void A_name_outside_the_registry_or_its_group_is_refused()
        {
            SettingsFile.Create(path, SongMgrSettings.Defaults());
            byte[] before = File.ReadAllBytes(path);

            Assert.Throws<ArgumentException>(() => SongMgrSettings.Save("SongMgrDestFoldr"));
            Assert.Throws<ArgumentException>(() => SongMgrSettings.Save("SongMgrDestFolder", "SongMgrDestFoldr"));
            Assert.Throws<ArgumentException>(() => SongMgrSettings.SaveValue("SongMgrDestFoldr", "x"));
            Assert.Throws<ArgumentException>(() => SongMgrSettings.Save("SingerMgrDefaultSingerDataTable"));
            Assert.Throws<ArgumentException>(() => SongMgrSettings.Save(SettingsGroup.LoadOnly));
            Assert.Throws<ArgumentException>(() => SongMgrSettings.Save(SettingsGroup.MainCfg, ("SongAddSpecialStr", "x")));

            Assert.Equal(before, File.ReadAllBytes(path));
        }
    }
}

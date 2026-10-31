using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// A settings file (CrazyKTV_SongMgr.cfg, CrazyKTV.cfg, UpdateDB.xml):
    /// &lt;Configeruation&gt;&lt;setting Name="…"&gt;&lt;Value&gt;…&lt;/Value&gt;&lt;/setting&gt;….
    /// Read once, then looked up by name; a missing setting reads as "".
    /// </summary>
    class SettingsFile
    {
        private readonly Dictionary<string, string> values = new Dictionary<string, string>();

        /// <summary>How a broken file is reported; a message box in the app.</summary>
        public static Action<string> ReportError = message => MessageBox.Show(message);

        private SettingsFile(XElement root)
        {
            if (root == null) return;
            foreach (XElement setting in root.Elements("setting"))
            {
                string name = (string)setting.Attribute("Name");
                if (name != null) values[name] = setting.Value; // a repeated name keeps its last value
            }
        }

        public string this[string name] => values.TryGetValue(name, out string value) ? value : "";

        /// <summary>Reads the file; a broken file is reported once and reads as empty.</summary>
        public static SettingsFile Load(string path)
        {
            try
            {
                return new SettingsFile(XElement.Load(path));
            }
            catch
            {
                ReportError("【" + Path.GetFileName(path) + "】設定檔內容有錯誤,請刪除後再執行。");
                return new SettingsFile(null);
            }
        }

        public static SettingsFile Load(Stream stream, string nameForErrors)
        {
            try
            {
                stream.Position = 0;
                return new SettingsFile(XElement.Load(stream));
            }
            catch
            {
                ReportError("【" + Path.GetFileName(nameForErrors) + "】設定檔內容有錯誤,請刪除後再執行。");
                return new SettingsFile(null);
            }
        }

        /// <summary>Writes a new settings file holding <paramref name="settings"/> in the given order.</summary>
        public static void Create(string path, IEnumerable<(string Name, string Value)> settings)
        {
            XDocument xmldoc = new XDocument
                (
                    new XDeclaration("1.0", "utf-16", null),
                    new XElement("Configeruation", settings.Select(s => new XElement("setting", new XAttribute("Name", s.Name), new XElement("Value", s.Value))))
                );
            xmldoc.Save(path);
        }
    }

    /// <summary>加歌模式: what happens to a song's file when it enters the 歌庫.</summary>
    enum SongAddMode
    {
        /// <summary>A value the settings file should never hold.</summary>
        Unknown = 0,
        /// <summary>1: move the file into the 歌庫 folder.</summary>
        Move = 1,
        /// <summary>2: copy the file into the 歌庫 folder.</summary>
        Copy = 2,
        /// <summary>3: leave the file where it is.</summary>
        InPlace = 3,
        /// <summary>4: 歌庫監視 — watch folders and add files where they are.</summary>
        Monitor = 4,
    }

    /// <summary>歌庫設定 (CrazyKTV_SongMgr.cfg): what a fresh settings file contains.</summary>
    static class SongMgrSettings
    {
        /// <summary>Every setting written when the file does not exist yet, with the current Global values, in file order.</summary>
        public static List<(string Name, string Value)> Defaults()
        {
            return new List<(string Name, string Value)>
            {
                ("CrazyktvDatabaseFile", Global.CrazyktvDatabaseFile),
                ("SongMgrSupportFormat", Global.SongMgrSupportFormat),
                ("SongMgrDestFolder", Global.SongMgrDestFolder),
                ("SongMgrSongAddMode", Global.SongMgrSongAddMode),
                ("SongMgrChorusMerge", Global.SongMgrChorusMerge),
                ("SongMgrMaxDigitCode", Global.SongMgrMaxDigitCode),
                ("SongMgrLangCode", Global.SongMgrLangCode),
                ("SongMgrSongType", Global.SongMgrSongType),
                ("SongMgrSongInfoSeparate", Global.SongMgrSongInfoSeparate),
                ("SongMgrChorusSeparate", Global.SongMgrChorusSeparate),
                ("SongAddDefaultSongLang", Global.SongAddDefaultSongLang),
                ("SongAddDefaultSingerType", Global.SongAddDefaultSingerType),
                ("SongAddDefaultSongTrack", Global.SongAddDefaultSongTrack),
                ("SongAddDefaultSongType", Global.SongAddDefaultSongType),
                ("SongAddDefaultSongVolume", Global.SongAddDefaultSongVolume),
                ("SongAddSpecialStr", Global.SongAddSpecialStr),
                ("SongMgrFolderStructure", Global.SongMgrFolderStructure),
                ("SongMgrFileStructure", Global.SongMgrFileStructure),
                ("SongMgrSongTrackMode", Global.SongMgrSongTrackMode),
                ("SongAddSongIdentificationMode", Global.SongAddSongIdentificationMode),
                ("SongAddDupSongMode", Global.SongAddDupSongMode),
                ("CrazyktvSongLangStr", string.Join(",", Global.CrazyktvSongLangList)),
                ("CrazyktvSongLangKeyWord", string.Join("|", Global.CrazyktvSongLangKeyWordList)),
                ("SongMgrBackupRemoveSong", Global.SongMgrBackupRemoveSong),
                ("SongAddEngSongNameFormat", Global.SongAddEngSongNameFormat),
                ("MainCfgAlwaysOnTop", Global.MainCfgAlwaysOnTop),
                ("CrazyktvSongDBVer", Global.CrazyktvSongDBVer),
                ("MainCfgHideSongDBConverterTabPage", Global.MainCfgHideSongDBConverterTabPage),
                ("MainCfgHideSongAddResultTabPage", Global.MainCfgHideSongAddResultTabPage),
                ("MainCfgHideSongLogTabPage", Global.MainCfgHideSongLogTabPage),
                ("MainCfgBackupRemoveSongDays", Global.MainCfgBackupRemoveSongDays),
                ("DBVerEnableDBVerUpdate", Global.DBVerEnableDBVerUpdate),
                ("MainCfgUIScale", Global.MainCfgUIScale),
                ("SongMaintenanceReplayGainVolume", Global.SongMaintenanceReplayGainVolume),
                ("SongAddEnableVolumeDetect", Global.SongAddEnableVolumeDetect),
                ("SongAddUseCustomSongID", Global.SongAddUseCustomSongID),
                ("MainCfgUIFont", Global.MainCfgUIFont),
                ("MainCfgEnableUIScale", Global.MainCfgEnableUIScale),
                ("SongMgrCustomSingerTypeStructure", Global.SongMgrCustomSingerTypeStructure),
                ("MainCfgEnableAutoUpdate", Global.MainCfgEnableAutoUpdate),
                ("SongAddEnableConvToTC", Global.SongAddEnableConvToTC),
                ("SongMgrEnableMonitorFolders", Global.SongMgrEnableMonitorFolders),
                ("SongMgrMonitorFolders", string.Join(",", Global.SongMgrMonitorFoldersList)),
                ("SingerMgrSyncSongSinger", Global.SingerMgrSyncSongSinger),
                ("SongMgrSingerGroup", Global.SongMgrSingerGroup),
                ("MainCfgHideApplyCashboxIdButton", Global.MainCfgHideApplyCashboxIdButton),
                ("MainCfgUICustomScale", Global.MainCfgUICustomScale),
                ("MainCfgBackupDB", Global.MainCfgBackupDB),
                ("MainCfgBackupDBPath", Global.MainCfgBackupDBPath),
                ("MainCfgPlayerCore", Global.MainCfgPlayerCore),
                ("MainCfgPlayerOutput", Global.MainCfgPlayerOutput),
                ("MainCfgPlayerEnableAudioCompressor", Global.MainCfgPlayerEnableAudioCompressor),
                ("MainCfgPlayerEnableAudioProcessor", Global.MainCfgPlayerEnableAudioProcessor),
                ("MainCfgPlayerDefaultVolume", Global.MainCfgPlayerDefaultVolume)
            };
        }
    }
}

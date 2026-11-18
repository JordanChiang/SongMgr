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

    /// <summary>Which save writes a 歌庫設定 setting.</summary>
    enum SettingsGroup
    {
        /// <summary>主要設定 page's 儲存 button.</summary>
        MainCfg,
        /// <summary>歌庫設定 page's 儲存 button.</summary>
        SongMgrCfg,
        /// <summary>加歌 page's 儲存設定 button.</summary>
        SongAdd,
        /// <summary>No button; written by name the moment it changes.</summary>
        Direct,
        /// <summary>Read from the file but never written, and not part of a fresh file.</summary>
        LoadOnly,
    }

    /// <summary>One 歌庫設定 entry: its name in the file, its save group, and the Global field it reads and writes as text.</summary>
    sealed class Setting
    {
        public string Name { get; }
        public SettingsGroup Group { get; }
        /// <summary>The value as stored in the file.</summary>
        public Func<string> Get { get; }
        /// <summary>Takes a non-empty value from the file.</summary>
        public Action<string> Set { get; }

        public Setting(string name, SettingsGroup group, Func<string> get, Action<string> set)
        {
            Name = name;
            Group = group;
            Get = get;
            Set = set;
        }
    }

    /// <summary>
    /// 歌庫設定 (CrazyKTV_SongMgr.cfg). <see cref="Entries"/> is the one list of settings:
    /// loading, saving and the content of a fresh file all come from it, so adding a setting is one line here.
    /// The values still live in the Global fields; their initial values are the defaults.
    /// </summary>
    static class SongMgrSettings
    {
        const SettingsGroup MainCfg = SettingsGroup.MainCfg, SongMgrCfg = SettingsGroup.SongMgrCfg, SongAdd = SettingsGroup.SongAdd, Direct = SettingsGroup.Direct;

        static Setting Text(string name, SettingsGroup group, Func<string> get, Action<string> set) => new Setting(name, group, get, set);

        static Setting List(string name, SettingsGroup group, char separator, Func<List<string>> get, Action<List<string>> set)
            => new Setting(name, group, () => string.Join(separator.ToString(), get()), v => set(new List<string>(v.Split(separator))));

        /// <summary>Every setting, in the order a fresh file lists them.</summary>
        public static readonly IReadOnlyList<Setting> Entries = new List<Setting>
        {
            Text("CrazyktvDatabaseFile", SongMgrCfg, () => Global.CrazyktvDatabaseFile, v => Global.CrazyktvDatabaseFile = v),
            Text("SongMgrSupportFormat", SongMgrCfg, () => Global.SongMgrSupportFormat, v => Global.SongMgrSupportFormat = v),
            Text("SongMgrDestFolder", SongMgrCfg, () => Global.SongMgrDestFolder, v => Global.SongMgrDestFolder = v),
            Text("SongMgrSongAddMode", SongMgrCfg, () => Global.SongMgrSongAddMode, v => Global.SongMgrSongAddMode = v),
            Text("SongMgrChorusMerge", SongMgrCfg, () => Global.SongMgrChorusMerge, v => Global.SongMgrChorusMerge = v),
            Text("SongMgrMaxDigitCode", SongMgrCfg, () => Global.SongMgrMaxDigitCode, v => Global.SongMgrMaxDigitCode = v),
            Text("SongMgrLangCode", SongMgrCfg, () => Global.SongMgrLangCode, v => Global.SongMgrLangCode = v),
            Text("SongMgrSongType", SongMgrCfg, () => Global.SongMgrSongType, v => Global.SongMgrSongType = v),
            Text("SongMgrSongInfoSeparate", SongMgrCfg, () => Global.SongMgrSongInfoSeparate, v => Global.SongMgrSongInfoSeparate = v),
            Text("SongMgrChorusSeparate", SongMgrCfg, () => Global.SongMgrChorusSeparate, v => Global.SongMgrChorusSeparate = v),
            Text("SongAddDefaultSongLang", SongAdd, () => Global.SongAddDefaultSongLang, v => Global.SongAddDefaultSongLang = v),
            Text("SongAddDefaultSingerType", SongAdd, () => Global.SongAddDefaultSingerType, v => Global.SongAddDefaultSingerType = v),
            Text("SongAddDefaultSongTrack", SongAdd, () => Global.SongAddDefaultSongTrack, v => Global.SongAddDefaultSongTrack = v),
            Text("SongAddDefaultSongType", SongAdd, () => Global.SongAddDefaultSongType, v => Global.SongAddDefaultSongType = v),
            Text("SongAddDefaultSongVolume", SongAdd, () => Global.SongAddDefaultSongVolume, v => Global.SongAddDefaultSongVolume = v),
            Text("SongAddSpecialStr", SongAdd, () => Global.SongAddSpecialStr, v => Global.SongAddSpecialStr = v),
            Text("SongMgrFolderStructure", SongMgrCfg, () => Global.SongMgrFolderStructure, v => Global.SongMgrFolderStructure = v),
            Text("SongMgrFileStructure", SongMgrCfg, () => Global.SongMgrFileStructure, v => Global.SongMgrFileStructure = v),
            Text("SongMgrSongTrackMode", SongMgrCfg, () => Global.SongMgrSongTrackMode, v => Global.SongMgrSongTrackMode = v),
            Text("SongAddSongIdentificationMode", SongAdd, () => Global.SongAddSongIdentificationMode, v => Global.SongAddSongIdentificationMode = v),
            Text("SongAddDupSongMode", SongAdd, () => Global.SongAddDupSongMode, v => Global.SongAddDupSongMode = v),
            List("CrazyktvSongLangStr", Direct, ',', () => Global.CrazyktvSongLangList, v => Global.CrazyktvSongLangList = v),
            List("CrazyktvSongLangKeyWord", Direct, '|', () => Global.CrazyktvSongLangKeyWordList, v => Global.CrazyktvSongLangKeyWordList = v),
            Text("SongMgrBackupRemoveSong", SongMgrCfg, () => Global.SongMgrBackupRemoveSong, v => Global.SongMgrBackupRemoveSong = v),
            Text("SongAddEngSongNameFormat", SongAdd, () => Global.SongAddEngSongNameFormat, v => Global.SongAddEngSongNameFormat = v),
            Text("MainCfgAlwaysOnTop", MainCfg, () => Global.MainCfgAlwaysOnTop, v => Global.MainCfgAlwaysOnTop = v),
            Text("MainCfgHideSongDBConverterTabPage", MainCfg, () => Global.MainCfgHideSongDBConverterTabPage, v => Global.MainCfgHideSongDBConverterTabPage = v),
            Text("MainCfgHideSongAddResultTabPage", MainCfg, () => Global.MainCfgHideSongAddResultTabPage, v => Global.MainCfgHideSongAddResultTabPage = v),
            Text("MainCfgHideSongLogTabPage", MainCfg, () => Global.MainCfgHideSongLogTabPage, v => Global.MainCfgHideSongLogTabPage = v),
            Text("MainCfgBackupRemoveSongDays", MainCfg, () => Global.MainCfgBackupRemoveSongDays, v => Global.MainCfgBackupRemoveSongDays = v),
            Text("MainCfgUIScale", MainCfg, () => Global.MainCfgUIScale, v => Global.MainCfgUIScale = v),
            Text("SongMaintenanceReplayGainVolume", Direct, () => Global.SongMaintenanceReplayGainVolume, v => Global.SongMaintenanceReplayGainVolume = v),
            Text("SongAddEnableVolumeDetect", SongAdd, () => Global.SongAddEnableVolumeDetect, v => Global.SongAddEnableVolumeDetect = v),
            Text("SongAddUseCustomSongID", SongAdd, () => Global.SongAddUseCustomSongID, v => Global.SongAddUseCustomSongID = v),
            Text("MainCfgUIFont", MainCfg, () => Global.MainCfgUIFont, v => Global.MainCfgUIFont = v),
            Text("MainCfgEnableUIScale", MainCfg, () => Global.MainCfgEnableUIScale, v => Global.MainCfgEnableUIScale = v),
            Text("SongMgrCustomSingerTypeStructure", SongMgrCfg, () => Global.SongMgrCustomSingerTypeStructure, v => Global.SongMgrCustomSingerTypeStructure = v),
            Text("MainCfgEnableAutoUpdate", MainCfg, () => Global.MainCfgEnableAutoUpdate, v => Global.MainCfgEnableAutoUpdate = v),
            Text("SongAddEnableConvToTC", SongAdd, () => Global.SongAddEnableConvToTC, v => Global.SongAddEnableConvToTC = v),
            Text("SongMgrEnableMonitorFolders", SongMgrCfg, () => Global.SongMgrEnableMonitorFolders, v => Global.SongMgrEnableMonitorFolders = v),
            List("SongMgrMonitorFolders", SongMgrCfg, ',', () => Global.SongMgrMonitorFoldersList, v => Global.SongMgrMonitorFoldersList = v),
            Text("SingerMgrSyncSongSinger", Direct, () => Global.SingerMgrSyncSongSinger, v => Global.SingerMgrSyncSongSinger = v),
            Text("SongMgrSingerGroup", SongMgrCfg, () => Global.SongMgrSingerGroup, v => Global.SongMgrSingerGroup = v),
            Text("MainCfgHideApplyCashboxIdButton", MainCfg, () => Global.MainCfgHideApplyCashboxIdButton, v => Global.MainCfgHideApplyCashboxIdButton = v),
            Text("MainCfgUICustomScale", MainCfg, () => Global.MainCfgUICustomScale, v => Global.MainCfgUICustomScale = v),
            Text("MainCfgBackupDB", MainCfg, () => Global.MainCfgBackupDB, v => Global.MainCfgBackupDB = v),
            Text("MainCfgBackupDBPath", MainCfg, () => Global.MainCfgBackupDBPath, v => Global.MainCfgBackupDBPath = v),
            Text("MainCfgPlayerCore", MainCfg, () => Global.MainCfgPlayerCore, v => Global.MainCfgPlayerCore = v),
            Text("MainCfgPlayerOutput", MainCfg, () => Global.MainCfgPlayerOutput, v => Global.MainCfgPlayerOutput = v),
            Text("MainCfgPlayerEnableAudioCompressor", MainCfg, () => Global.MainCfgPlayerEnableAudioCompressor, v => Global.MainCfgPlayerEnableAudioCompressor = v),
            Text("MainCfgPlayerEnableAudioProcessor", MainCfg, () => Global.MainCfgPlayerEnableAudioProcessor, v => Global.MainCfgPlayerEnableAudioProcessor = v),
            Text("MainCfgPlayerDefaultVolume", MainCfg, () => Global.MainCfgPlayerDefaultVolume, v => Global.MainCfgPlayerDefaultVolume = v),
            Text("MainCfgAutoSongFileCount", MainCfg, () => Global.MainCfgAutoSongFileCount, v => Global.MainCfgAutoSongFileCount = v),
            Text("SingerMgrDefaultSingerDataTable", SettingsGroup.LoadOnly, () => Global.SingerMgrDefaultSingerDataTable, v => Global.SingerMgrDefaultSingerDataTable = v),
        };

        /// <summary>Takes every non-empty value in <paramref name="cfg"/> into its Global field; an empty or missing value keeps the field as it is.</summary>
        public static void Load(SettingsFile cfg)
        {
            foreach (Setting setting in Entries)
            {
                string value = cfg[setting.Name];
                if (value != "") setting.Set(value);
            }
        }

        /// <summary>
        /// Writes every setting of <paramref name="group"/> with one read and one write of the file.
        /// An override replaces what is written for that setting, without changing its Global field.
        /// </summary>
        public static void Save(SettingsGroup group, params (string Name, string Value)[] overrides)
        {
            if (group == SettingsGroup.LoadOnly) throw new ArgumentException("唯讀的設定不會存檔。", nameof(group));
            var written = new Dictionary<string, string>();
            foreach (var o in overrides)
            {
                if (Find(o.Name).Group != group) throw new ArgumentException("【" + o.Name + "】不屬於 " + group + "。", nameof(overrides));
                written[o.Name] = o.Value;
            }
            Write(Entries.Where(s => s.Group == group)
                .Select(s => (s.Name, written.TryGetValue(s.Name, out string value) ? value : s.Get())).ToList());
        }

        /// <summary>Writes the named settings with one read and one write of the file.</summary>
        public static void Save(params string[] names)
        {
            Write(names.Select(Savable).Select(s => (s.Name, s.Get())).ToList());
        }

        /// <summary>Writes <paramref name="value"/> for the named setting without changing its Global field.</summary>
        public static void SaveValue(string name, string value)
        {
            Write(new[] { (Savable(name).Name, value) });
        }

        /// <summary>Every setting written when the file does not exist yet, with the current Global values, in file order.</summary>
        public static List<(string Name, string Value)> Defaults()
        {
            return Entries.Where(s => s.Group != SettingsGroup.LoadOnly).Select(s => (s.Name, s.Get())).ToList();
        }

        static Setting Find(string name)
        {
            return Entries.FirstOrDefault(s => s.Name == name) ?? throw new ArgumentException("【" + name + "】不是歌庫設定。", nameof(name));
        }

        static Setting Savable(string name)
        {
            Setting setting = Find(name);
            if (setting.Group == SettingsGroup.LoadOnly) throw new ArgumentException("【" + name + "】是唯讀的設定,不會存檔。", nameof(name));
            return setting;
        }

        /// <summary>
        /// The same file the old key-by-key writes produced: a setting already in the file is updated where it is
        /// (every copy, if the name repeats), a new one is added at the end, and anything else is kept.
        /// </summary>
        static void Write(IEnumerable<(string Name, string Value)> settings)
        {
            XDocument xmldoc = XDocument.Load(Global.SongMgrCfgFile);
            XElement rootElement = xmldoc.Element("Configeruation");
            foreach (var (name, value) in settings)
            {
                List<XElement> existing = rootElement.Elements("setting").Where(n => (string)n.Attribute("Name") == name).ToList();
                if (existing.Count > 0)
                {
                    foreach (XElement node in existing) node.Element("Value").Value = value;
                }
                else
                {
                    rootElement.Add(new XElement("setting", new XAttribute("Name", name), new XElement("Value", value)));
                }
            }
            xmldoc.Save(Global.SongMgrCfgFile);
        }
    }
}

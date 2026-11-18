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

    /// <summary>One 歌庫設定 entry: its name in the file, its save group, and its value as the file stores it.</summary>
    abstract class Setting
    {
        public string Name { get; }
        public SettingsGroup Group { get; }

        protected Setting(string name, SettingsGroup group)
        {
            Name = name;
            Group = group;
        }

        /// <summary>The value as stored in the file.</summary>
        public abstract string Get();

        /// <summary>Takes a non-empty value from the file; one that does not parse keeps the current value.</summary>
        public abstract void Set(string text);

        /// <summary>A copy of the current value, for <see cref="SongMgrSettings.Snapshot"/>.</summary>
        internal abstract object Capture();

        internal abstract void Restore(object state);
    }

    /// <summary>A 歌庫設定 entry holding a <typeparamref name="T"/>; it starts at its default.</summary>
    sealed class Setting<T> : Setting
    {
        public delegate bool Parser(string text, out T value);

        readonly T initial;
        readonly Func<T, string> format;
        readonly Parser parse;
        readonly Func<T, T> copy;

        public T Value { get; set; }

        /// <summary>The value a fresh settings file starts with; a new copy each time for a list.</summary>
        public T Default => copy(initial);

        public Setting(string name, SettingsGroup group, T initial, Func<T, string> format, Parser parse, Func<T, T> copy)
            : base(name, group)
        {
            this.initial = initial;
            this.format = format;
            this.parse = parse;
            this.copy = copy;
            Value = copy(initial);
        }

        public override string Get() => format(Value);

        public override void Set(string text)
        {
            if (parse(text, out T value)) Value = value;
        }

        internal override object Capture() => copy(Value);

        internal override void Restore(object state) => Value = copy((T)state);
    }

    /// <summary>
    /// 歌庫設定 (CrazyKTV_SongMgr.cfg). Each setting is one field here, named as in the file, holding its type and default;
    /// <see cref="Entries"/> lists them in declaration order, which is the order of a fresh file.
    /// Loading, saving and the content of a fresh file all come from it, so adding a setting is one line here
    /// (plus its view in <see cref="Global"/> while code still reads it there).
    /// </summary>
    static class SongMgrSettings
    {
        const SettingsGroup MainCfg = SettingsGroup.MainCfg, SongMgrCfg = SettingsGroup.SongMgrCfg, SongAdd = SettingsGroup.SongAdd, Direct = SettingsGroup.Direct, LoadOnly = SettingsGroup.LoadOnly;

        // Declared before the settings: static fields are initialized in text order, and each setting adds itself.
        static readonly List<Setting> entries = new List<Setting>();

        /// <summary>Every setting, in the order a fresh file lists them.</summary>
        public static IReadOnlyList<Setting> Entries => entries;

        static Setting<T> Add<T>(Setting<T> setting)
        {
            entries.Add(setting);
            return setting;
        }

        static Setting<string> Text(string name, SettingsGroup group, string initial)
            => Add(new Setting<string>(name, group, initial, v => v, (string text, out string v) => { v = text; return true; }, v => v));

        /// <summary>Stored as "True"/"False"; read in any case, and a value that is neither keeps the current one.</summary>
        static Setting<bool> Bool(string name, SettingsGroup group, bool initial)
            => Add(new Setting<bool>(name, group, initial, v => v ? "True" : "False", bool.TryParse, v => v));

        static Setting<List<string>> List(string name, SettingsGroup group, char separator, params string[] initial)
            => Add(new Setting<List<string>>(name, group, new List<string>(initial),
                v => string.Join(separator.ToString(), v),
                (string text, out List<string> v) => { v = new List<string>(text.Split(separator)); return true; },
                v => new List<string>(v)));

        public static readonly Setting<string> CrazyktvDatabaseFile = Text("CrazyktvDatabaseFile", SongMgrCfg, Application.StartupPath + @"\CrazySong.mdb");
        public static readonly Setting<string> SongMgrSupportFormat = Text("SongMgrSupportFormat", SongMgrCfg, ".avi;.flv;.dat;.mkv;.mp4;.mov;.mpeg;.mpg;.rmvb;.ts;.vob;.webm;.wmv");
        public static readonly Setting<string> SongMgrDestFolder = Text("SongMgrDestFolder", SongMgrCfg, "");
        public static readonly Setting<string> SongMgrSongAddMode = Text("SongMgrSongAddMode", SongMgrCfg, "3");
        public static readonly Setting<bool> SongMgrChorusMerge = Bool("SongMgrChorusMerge", SongMgrCfg, false);
        public static readonly Setting<string> SongMgrMaxDigitCode = Text("SongMgrMaxDigitCode", SongMgrCfg, "2");
        public static readonly Setting<string> SongMgrLangCode = Text("SongMgrLangCode", SongMgrCfg, "");
        public static readonly Setting<string> SongMgrSongType = Text("SongMgrSongType", SongMgrCfg, "");
        public static readonly Setting<string> SongMgrSongInfoSeparate = Text("SongMgrSongInfoSeparate", SongMgrCfg, "1");
        public static readonly Setting<string> SongMgrChorusSeparate = Text("SongMgrChorusSeparate", SongMgrCfg, "1");
        public static readonly Setting<string> SongAddDefaultSongLang = Text("SongAddDefaultSongLang", SongAdd, "11");
        public static readonly Setting<string> SongAddDefaultSingerType = Text("SongAddDefaultSingerType", SongAdd, "8");
        public static readonly Setting<string> SongAddDefaultSongTrack = Text("SongAddDefaultSongTrack", SongAdd, "1");
        public static readonly Setting<string> SongAddDefaultSongType = Text("SongAddDefaultSongType", SongAdd, "1");
        public static readonly Setting<string> SongAddDefaultSongVolume = Text("SongAddDefaultSongVolume", SongAdd, "100");
        public static readonly Setting<string> SongAddSpecialStr = Text("SongAddSpecialStr", SongAdd, "");
        public static readonly Setting<string> SongMgrFolderStructure = Text("SongMgrFolderStructure", SongMgrCfg, "1");
        public static readonly Setting<string> SongMgrFileStructure = Text("SongMgrFileStructure", SongMgrCfg, "1");
        public static readonly Setting<bool> SongMgrSongTrackMode = Bool("SongMgrSongTrackMode", SongMgrCfg, false);
        public static readonly Setting<string> SongAddSongIdentificationMode = Text("SongAddSongIdentificationMode", SongAdd, "1");
        public static readonly Setting<string> SongAddDupSongMode = Text("SongAddDupSongMode", SongAdd, "1");
        public static readonly Setting<List<string>> CrazyktvSongLangStr = List("CrazyktvSongLangStr", Direct, ',', "國語", "台語", "粵語", "日語", "英語", "客語", "原住民語", "韓語", "兒歌", "其它");
        public static readonly Setting<List<string>> CrazyktvSongLangKeyWord = List("CrazyktvSongLangKeyWord", Direct, '|', "國語,國", "台語,台,閩南,閩,臺語,臺", "粵語,粵,廣東", "日語,日文,日", "英語,英文,英", "客語,客", "原住民語,民謠", "韓語,韓", "兒歌,兒", "其它");
        public static readonly Setting<bool> SongMgrBackupRemoveSong = Bool("SongMgrBackupRemoveSong", SongMgrCfg, false);
        public static readonly Setting<bool> SongAddEngSongNameFormat = Bool("SongAddEngSongNameFormat", SongAdd, false);
        public static readonly Setting<bool> MainCfgAlwaysOnTop = Bool("MainCfgAlwaysOnTop", MainCfg, false);
        public static readonly Setting<bool> MainCfgHideSongDBConverterTabPage = Bool("MainCfgHideSongDBConverterTabPage", MainCfg, false);
        public static readonly Setting<bool> MainCfgHideSongAddResultTabPage = Bool("MainCfgHideSongAddResultTabPage", MainCfg, true);
        public static readonly Setting<bool> MainCfgHideSongLogTabPage = Bool("MainCfgHideSongLogTabPage", MainCfg, true);
        public static readonly Setting<string> MainCfgBackupRemoveSongDays = Text("MainCfgBackupRemoveSongDays", MainCfg, "7");
        public static readonly Setting<string> MainCfgUIScale = Text("MainCfgUIScale", MainCfg, "3");
        public static readonly Setting<string> SongMaintenanceReplayGainVolume = Text("SongMaintenanceReplayGainVolume", Direct, "50");
        public static readonly Setting<bool> SongAddEnableVolumeDetect = Bool("SongAddEnableVolumeDetect", SongAdd, false);
        public static readonly Setting<bool> SongAddUseCustomSongID = Bool("SongAddUseCustomSongID", SongAdd, false);
        public static readonly Setting<string> MainCfgUIFont = Text("MainCfgUIFont", MainCfg, "Arial");
        public static readonly Setting<bool> MainCfgEnableUIScale = Bool("MainCfgEnableUIScale", MainCfg, false);
        public static readonly Setting<string> SongMgrCustomSingerTypeStructure = Text("SongMgrCustomSingerTypeStructure", SongMgrCfg, "1,1,1,1,1,1,1,2");
        public static readonly Setting<bool> MainCfgEnableAutoUpdate = Bool("MainCfgEnableAutoUpdate", MainCfg, true);
        public static readonly Setting<bool> SongAddEnableConvToTC = Bool("SongAddEnableConvToTC", SongAdd, false);
        public static readonly Setting<bool> SongMgrEnableMonitorFolders = Bool("SongMgrEnableMonitorFolders", SongMgrCfg, false);
        public static readonly Setting<List<string>> SongMgrMonitorFolders = List("SongMgrMonitorFolders", SongMgrCfg, ',', "", "", "", "", "");
        public static readonly Setting<bool> SingerMgrSyncSongSinger = Bool("SingerMgrSyncSongSinger", Direct, true);
        public static readonly Setting<string> SongMgrSingerGroup = Text("SongMgrSingerGroup", SongMgrCfg, "");
        public static readonly Setting<bool> MainCfgHideApplyCashboxIdButton = Bool("MainCfgHideApplyCashboxIdButton", MainCfg, true);
        public static readonly Setting<string> MainCfgUICustomScale = Text("MainCfgUICustomScale", MainCfg, "100");
        public static readonly Setting<bool> MainCfgBackupDB = Bool("MainCfgBackupDB", MainCfg, false);
        public static readonly Setting<string> MainCfgBackupDBPath = Text("MainCfgBackupDBPath", MainCfg, "");
        public static readonly Setting<string> MainCfgPlayerCore = Text("MainCfgPlayerCore", MainCfg, "1");
        public static readonly Setting<string> MainCfgPlayerOutput = Text("MainCfgPlayerOutput", MainCfg, "1");
        public static readonly Setting<bool> MainCfgPlayerEnableAudioCompressor = Bool("MainCfgPlayerEnableAudioCompressor", MainCfg, false);
        public static readonly Setting<bool> MainCfgPlayerEnableAudioProcessor = Bool("MainCfgPlayerEnableAudioProcessor", MainCfg, false);
        public static readonly Setting<string> MainCfgPlayerDefaultVolume = Text("MainCfgPlayerDefaultVolume", MainCfg, "100");
        public static readonly Setting<bool> MainCfgAutoSongFileCount = Bool("MainCfgAutoSongFileCount", MainCfg, false);
        public static readonly Setting<string> SingerMgrDefaultSingerDataTable = Text("SingerMgrDefaultSingerDataTable", LoadOnly, "ktv_Singer");

        /// <summary>A copy of every setting's value, for tests to put back with <see cref="Restore"/>.</summary>
        public static IReadOnlyDictionary<Setting, object> Snapshot()
        {
            return entries.ToDictionary(s => s, s => s.Capture());
        }

        /// <summary>Puts back the values of a <see cref="Snapshot"/>; it can be restored more than once.</summary>
        public static void Restore(IReadOnlyDictionary<Setting, object> snapshot)
        {
            foreach (var pair in snapshot) pair.Key.Restore(pair.Value);
        }
        /// <summary>Takes every non-empty value in <paramref name="cfg"/> into its setting; an empty or missing value, or one that does not parse, keeps the current value.</summary>
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

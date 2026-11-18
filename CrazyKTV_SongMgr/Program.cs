using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    static class Program
    {
        static Program()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string assemblyName = new AssemblyName(args.Name).Name + ".dll";
                string[] searchPaths = new[] { "Libs", @"SongMgr\Libs", "Codec_x86" };
                foreach (string path in searchPaths)
                {
                    string fullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path, assemblyName);
                    if (File.Exists(fullPath))
                    {
                        return Assembly.LoadFrom(fullPath);
                    }
                }
                return null;
            };

            try
            {
                AppDomain.CurrentDomain.SetData("PRIVATE_BINPATH", @"Libs;SongMgr\Libs;Codec_x86");
                AppDomain.CurrentDomain.SetData("BINPATH_PROBE_ONLY", @"Libs;SongMgr\Libs;Codec_x86");
                var m = typeof(AppDomainSetup).GetMethod("UpdateContextProperty", BindingFlags.NonPublic | BindingFlags.Static);
                var funsion = typeof(AppDomain).GetMethod("GetFusionContext", BindingFlags.NonPublic | BindingFlags.Instance);
                if (m != null && funsion != null)
                {
                    m.Invoke(null, new object[] { funsion.Invoke(AppDomain.CurrentDomain, null), "PRIVATE_BINPATH", @"Libs;SongMgr\Libs;Codec_x86" });
                }
            }
            catch { }
        }

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // System.Data.SQLite's first SQLiteConnection looks for SEE / Eagle plugin files beside its own dll via
            // Assembly.Location, which is "" once the dll is bundled into the single-file exe, and then fails with
            // "Value cannot be null. (Parameter 'path1')". SongMgr uses no such plugins, so skip that step.
            Environment.SetEnvironmentVariable("SQLite_NoConfigure", "1");

            if (Environment.OSVersion.Version.Major >= 6) NativeMethods.SetProcessDPIAware();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new MainForm());
        }
    }


    /// <summary>
    /// Shared app state. A property that reads and writes <c>SongMgrSettings.…Value</c> is a view of a 歌庫設定:
    /// its type and default live in <see cref="SongMgrSettings"/>.
    /// </summary>
    class Global
    {
        public static float DPIScalingFactor = 1;

        public static string CrazyktvDatabaseFile { get => SongMgrSettings.CrazyktvDatabaseFile.Value; set => SongMgrSettings.CrazyktvDatabaseFile.Value = value; }
        public static string CrazyktvSongMgrDatabaseFile = Application.StartupPath + @"\SongMgr\CrazySongMgr.mdb";
        public static string CrazyktvCfgFile = Application.StartupPath + @"\CrazyKTV.cfg";
        public static bool CrazyktvDatabaseStatus = false;
        public static bool CrazyktvDatabaseIsOld = false;
        public static bool CrazyktvDatabaseError = false;
        public static bool CrazyktvDatabaseMaxDigitCode = true;
        public static bool SongMgrDatabaseError = false;
        public static bool DatabaseUpdateFinished = false;
        public static bool InitializeSongData = false;

        public static string SongMgrVer = FileVersionInfo.GetVersionInfo(Environment.ProcessPath).FileVersion.ToString().Replace(".", "");
        public static string SongMgrCfgFile = Application.StartupPath + @"\CrazyKTV_SongMgr.cfg";
        public static string SongMgrSupportFormat { get => SongMgrSettings.SongMgrSupportFormat.Value; set => SongMgrSettings.SongMgrSupportFormat.Value = value; }
        public static string SongMgrDestFolder { get => SongMgrSettings.SongMgrDestFolder.Value; set => SongMgrSettings.SongMgrDestFolder.Value = value; }
        public static string SongMgrSongAddMode { get => SongMgrSettings.SongMgrSongAddMode.Value; set => SongMgrSettings.SongMgrSongAddMode.Value = value; }

        /// <summary><see cref="SongMgrSongAddMode"/> as a name. The string stays the stored and bound form.</summary>
        public static SongAddMode AddMode
        {
            get
            {
                switch (SongMgrSongAddMode)
                {
                    case "1": return SongAddMode.Move;
                    case "2": return SongAddMode.Copy;
                    case "3": return SongAddMode.InPlace;
                    case "4": return SongAddMode.Monitor;
                    default: return SongAddMode.Unknown;
                }
            }
        }
        public static bool SongMgrChorusMerge { get => SongMgrSettings.SongMgrChorusMerge.Value; set => SongMgrSettings.SongMgrChorusMerge.Value = value; }
        public static string SongMgrMaxDigitCode { get => SongMgrSettings.SongMgrMaxDigitCode.Value; set => SongMgrSettings.SongMgrMaxDigitCode.Value = value; }
        public static string SongMgrLangCode { get => SongMgrSettings.SongMgrLangCode.Value; set => SongMgrSettings.SongMgrLangCode.Value = value; }
        public static string SongMgrSongType { get => SongMgrSettings.SongMgrSongType.Value; set => SongMgrSettings.SongMgrSongType.Value = value; }
        public static string SongMgrSongInfoSeparate { get => SongMgrSettings.SongMgrSongInfoSeparate.Value; set => SongMgrSettings.SongMgrSongInfoSeparate.Value = value; }
        public static string SongMgrChorusSeparate { get => SongMgrSettings.SongMgrChorusSeparate.Value; set => SongMgrSettings.SongMgrChorusSeparate.Value = value; }
        public static string SongMgrFolderStructure { get => SongMgrSettings.SongMgrFolderStructure.Value; set => SongMgrSettings.SongMgrFolderStructure.Value = value; }
        public static string SongMgrFileStructure { get => SongMgrSettings.SongMgrFileStructure.Value; set => SongMgrSettings.SongMgrFileStructure.Value = value; }
        public static bool SongMgrSongTrackMode { get => SongMgrSettings.SongMgrSongTrackMode.Value; set => SongMgrSettings.SongMgrSongTrackMode.Value = value; }
        public static bool SongMgrBackupRemoveSong { get => SongMgrSettings.SongMgrBackupRemoveSong.Value; set => SongMgrSettings.SongMgrBackupRemoveSong.Value = value; }
        public static string SongMgrCustomSingerTypeStructure { get => SongMgrSettings.SongMgrCustomSingerTypeStructure.Value; set => SongMgrSettings.SongMgrCustomSingerTypeStructure.Value = value; }
        public static bool SongMgrEnableMonitorFolders { get => SongMgrSettings.SongMgrEnableMonitorFolders.Value; set => SongMgrSettings.SongMgrEnableMonitorFolders.Value = value; }
        public static string SongMgrSingerGroup { get => SongMgrSettings.SongMgrSingerGroup.Value; set => SongMgrSettings.SongMgrSingerGroup.Value = value; }
        public static bool SongMgrInitializeStatus = false;
        public static bool SongMgrDBVerErrorUIStatus = false;
        public static List<string> SongMgrCustomSingerTypeStructureList = new List<string>() { "男", "女", "團", "合唱", "外男", "外女", "外團", "其他", "歌星姓氏", "全部歌星", "新進" };
        public static List<string> SongMgrMonitorFoldersList { get => SongMgrSettings.SongMgrMonitorFolders.Value; set => SongMgrSettings.SongMgrMonitorFolders.Value = value; }

        public static string SingerMgrDefaultSingerDataTable { get => SongMgrSettings.SingerMgrDefaultSingerDataTable.Value; set => SongMgrSettings.SingerMgrDefaultSingerDataTable.Value = value; }
        public static bool SingerMgrSyncSongSinger { get => SongMgrSettings.SingerMgrSyncSongSinger.Value; set => SongMgrSettings.SingerMgrSyncSongSinger.Value = value; }
        public static string SingerMgrLastNameSortMethod = "1";

        public static string SongAddDefaultSongLang { get => SongMgrSettings.SongAddDefaultSongLang.Value; set => SongMgrSettings.SongAddDefaultSongLang.Value = value; }
        public static string SongAddDefaultSingerType { get => SongMgrSettings.SongAddDefaultSingerType.Value; set => SongMgrSettings.SongAddDefaultSingerType.Value = value; }
        public static string SongAddDefaultSongTrack { get => SongMgrSettings.SongAddDefaultSongTrack.Value; set => SongMgrSettings.SongAddDefaultSongTrack.Value = value; }
        public static string SongAddDefaultSongType { get => SongMgrSettings.SongAddDefaultSongType.Value; set => SongMgrSettings.SongAddDefaultSongType.Value = value; }
        public static string SongAddDefaultSongVolume { get => SongMgrSettings.SongAddDefaultSongVolume.Value; set => SongMgrSettings.SongAddDefaultSongVolume.Value = value; }
        public static string SongAddSpecialStr { get => SongMgrSettings.SongAddSpecialStr.Value; set => SongMgrSettings.SongAddSpecialStr.Value = value; }
        public static string SongAddSongIdentificationMode { get => SongMgrSettings.SongAddSongIdentificationMode.Value; set => SongMgrSettings.SongAddSongIdentificationMode.Value = value; }
        public static string SongAddDupSongMode { get => SongMgrSettings.SongAddDupSongMode.Value; set => SongMgrSettings.SongAddDupSongMode.Value = value; }
        public static bool SongAddEngSongNameFormat { get => SongMgrSettings.SongAddEngSongNameFormat.Value; set => SongMgrSettings.SongAddEngSongNameFormat.Value = value; }
        public static bool SongAddUseCustomSongID { get => SongMgrSettings.SongAddUseCustomSongID.Value; set => SongMgrSettings.SongAddUseCustomSongID.Value = value; }
        public static bool SongAddEnableConvToTC { get => SongMgrSettings.SongAddEnableConvToTC.Value; set => SongMgrSettings.SongAddEnableConvToTC.Value = value; }
        public static bool SongAddEnableVolumeDetect { get => SongMgrSettings.SongAddEnableVolumeDetect.Value; set => SongMgrSettings.SongAddEnableVolumeDetect.Value = value; }

        public static bool MainCfgAlwaysOnTop { get => SongMgrSettings.MainCfgAlwaysOnTop.Value; set => SongMgrSettings.MainCfgAlwaysOnTop.Value = value; }
        public static bool MainCfgEnableAutoUpdate { get => SongMgrSettings.MainCfgEnableAutoUpdate.Value; set => SongMgrSettings.MainCfgEnableAutoUpdate.Value = value; }
        public static bool MainCfgAutoSongFileCount { get => SongMgrSettings.MainCfgAutoSongFileCount.Value; set => SongMgrSettings.MainCfgAutoSongFileCount.Value = value; }
        public static bool MainCfgHideSongDBConverterTabPage { get => SongMgrSettings.MainCfgHideSongDBConverterTabPage.Value; set => SongMgrSettings.MainCfgHideSongDBConverterTabPage.Value = value; }
        public static bool MainCfgHideSongAddResultTabPage { get => SongMgrSettings.MainCfgHideSongAddResultTabPage.Value; set => SongMgrSettings.MainCfgHideSongAddResultTabPage.Value = value; }
        public static bool MainCfgHideSongLogTabPage { get => SongMgrSettings.MainCfgHideSongLogTabPage.Value; set => SongMgrSettings.MainCfgHideSongLogTabPage.Value = value; }
        public static bool MainCfgHideApplyCashboxIdButton { get => SongMgrSettings.MainCfgHideApplyCashboxIdButton.Value; set => SongMgrSettings.MainCfgHideApplyCashboxIdButton.Value = value; }
        public static string MainCfgBackupRemoveSongDays { get => SongMgrSettings.MainCfgBackupRemoveSongDays.Value; set => SongMgrSettings.MainCfgBackupRemoveSongDays.Value = value; }
        public static string MainCfgUIScale { get => SongMgrSettings.MainCfgUIScale.Value; set => SongMgrSettings.MainCfgUIScale.Value = value; }
        public static string MainCfgUICustomScale { get => SongMgrSettings.MainCfgUICustomScale.Value; set => SongMgrSettings.MainCfgUICustomScale.Value = value; }
        public static string MainCfgUIFont { get => SongMgrSettings.MainCfgUIFont.Value; set => SongMgrSettings.MainCfgUIFont.Value = value; }
        public static bool MainCfgEnableUIScale { get => SongMgrSettings.MainCfgEnableUIScale.Value; set => SongMgrSettings.MainCfgEnableUIScale.Value = value; }
        public static bool MainCfgBackupDB { get => SongMgrSettings.MainCfgBackupDB.Value; set => SongMgrSettings.MainCfgBackupDB.Value = value; }
        public static string MainCfgBackupDBPath { get => SongMgrSettings.MainCfgBackupDBPath.Value; set => SongMgrSettings.MainCfgBackupDBPath.Value = value; }
        public static string MainCfgPlayerCore { get => SongMgrSettings.MainCfgPlayerCore.Value; set => SongMgrSettings.MainCfgPlayerCore.Value = value; }
        public static string MainCfgPlayerOutput { get => SongMgrSettings.MainCfgPlayerOutput.Value; set => SongMgrSettings.MainCfgPlayerOutput.Value = value; }
        public static bool MainCfgPlayerEnableAudioCompressor { get => SongMgrSettings.MainCfgPlayerEnableAudioCompressor.Value; set => SongMgrSettings.MainCfgPlayerEnableAudioCompressor.Value = value; }
        public static bool MainCfgPlayerEnableAudioProcessor { get => SongMgrSettings.MainCfgPlayerEnableAudioProcessor.Value; set => SongMgrSettings.MainCfgPlayerEnableAudioProcessor.Value = value; }
        public static string MainCfgPlayerDefaultVolume { get => SongMgrSettings.MainCfgPlayerDefaultVolume.Value; set => SongMgrSettings.MainCfgPlayerDefaultVolume.Value = value; }

        public static string SongMaintenanceEnableMultiSongPath = "False";
        public static string SongMaintenanceReplayGainVolume { get => SongMgrSettings.SongMaintenanceReplayGainVolume.Value; set => SongMgrSettings.SongMaintenanceReplayGainVolume.Value = value; }

        public static List<int> MTotalList = new List<int>();
        public static List<int> TotalList = new List<int>();

        public static List<string> CrazyktvSongLangList { get => SongMgrSettings.CrazyktvSongLangStr.Value; set => SongMgrSettings.CrazyktvSongLangStr.Value = value; }
        public static List<string> CrazyktvSongLangKeyWordList { get => SongMgrSettings.CrazyktvSongLangKeyWord.Value; set => SongMgrSettings.CrazyktvSongLangKeyWord.Value = value; }
        public static List<string> CrazyktvSingerTypeList = new List<string>() { "男歌星", "女歌星", "團體", "合唱", "外國男", "外國女", "外國團體", "其他", "未使用", "未使用", "新進歌星" };
        public static List<string> CrazyktvSingerTypeKeyWordList = new List<string>() { "男,男星,男歌星,男歌手", "女,女星,女歌星,女歌手", "團,團體,樂團", "合唱,對唱", "外男,外國男星,外國男歌星,外國男歌手", "外女,外國女星,外國女歌星,外國女歌手", "外團,外國團體,外國樂團", "未知,其他,其他歌星,其他歌手" };
        public static List<string> CrazyktvSongTrackList = new List<string>() { "V0", "VR", "VL", "V3", "V4", "V5" };
        public static List<string> CrazyktvSongTrackWordList = new List<string>() { "立體聲", "右聲道 / 音軌1", "左聲道 / 音軌2", "音軌3", "音軌4", "音軌5" };
        public static List<string> CrazyktvSongTrackKeyWordList = new List<string>() { "v0,立體", "vr,r,右", "vl,l,左", "v3", "v4", "v5" };

        public static List<string> SingerList = new List<string>();
        public static List<string> SingerLowCaseList = new List<string>();
        public static List<string> SingerTypeList = new List<string>();
        public static List<string> AllSingerList = new List<string>();
        public static List<string> AllSingerLowCaseList = new List<string>();
        public static List<string> AllSingerTypeList = new List<string>();

        public static List<string> SingerGroupList = new List<string>();
        public static List<int> GroupSingerIdList = new List<int>();
        public static List<string> GroupSingerLowCaseList = new List<string>();

        public static List<string> PhoneticsWordList = new List<string>();
        public static List<string> PhoneticsSpellList = new List<string>();
        public static List<string> PhoneticsPenStyleList = new List<string>();
        public static List<string> PhoneticsStrokesList = new List<string>();

        public static List<string> CashboxSongDataFullList = new List<string>();
        public static List<string> CashboxSongDataLangList = new List<string>();
        public static List<string> CashboxSongDataLowCaseList = new List<string>();
        public static List<string> CashboxSongDataFuzzyList = new List<string>();
        public static List<string> CashboxFullMatchSongList = new List<string>();
        public static List<string> CashboxFullAnalysisSongList = new List<string>();
        public static List<string> CashboxHaveSongList = new List<string>();

        public static List<SongRecord> SongDBConvValueList = new List<SongRecord>();
        public static List<string> SongDBConvJetktvLangList = new List<string>();
        public static List<string> SongDBConvJetktvPathList = new List<string>();
        public static List<string> SongDBConvHK2SongPathList = new List<string>();

        public static DataTable SongDT = new DataTable();
        public static DataTable SingerDT = new DataTable();
        public static DataTable AllSingerDT = new DataTable();
        public static OperationLog DuplicateSongLog = new OperationLog();
        public static OperationLog FailureSongLog = new OperationLog();
        public static DataTable SongSrcDT = new DataTable();
        public static DataTable SingerSrcDT = new DataTable();
        public static DataTable SongPathSrcDT = new DataTable();
        public static DataTable ExtensionSrcDT = new DataTable();
        public static OperationLog SongLog = new OperationLog();
        public static DataTable FavoriteUserDT = new DataTable();
        public static DataTable CashboxNewSongDateDT = new DataTable();

        public static DateTime TimerStartTime = new DateTime();
        public static DateTime TimerEndTime = new DateTime();

        public static string SongQueryQueryType = "SongQuery";
        public static string SongQueryFuzzyQuery = "True";
        public static string SongQueryFilter = "全部";
        public static bool SongQuerySynonymousQuery = true;

        public static bool SongQueryMultiEdit;
        public static List<bool> SongQueryMultiEditUpdateList = new List<bool>() { false, false, false, false, false, false, false, false };

        public static List<SongRecord> SongQueryDataGridViewSelectList = new List<SongRecord>();
        public static string SongQueryDataGridViewRestoreCurrentRow = "";
        public static List<string> SongQueryDataGridViewRestoreSelectList = new List<string>();

        public static bool SongAddMultiEdit;
        public static List<bool> SongAddMultiEditUpdateList = new List<bool>() { false, false, false, false, false, false, false, false };

        public static List<SongRecord> SongAddDataGridViewSelectList = new List<SongRecord>();
        public static string SongAddDataGridViewRestoreCurrentRow = "";
        public static List<string> SongAddDataGridViewRestoreSelectList = new List<string>();

        public static bool SingerMgrMultiEdit;
        public static List<bool> SingerMgrMultiEditUpdateList = new List<bool>() { false };
        public static List<string> SingerMgrDataGridViewSelectList = new List<string>();
        public static string SingerMgrDataGridViewRestoreCurrentRow = "";
        public static List<string> SingerMgrDataGridViewRestoreSelectList = new List<string>();

        #if DEBUG
        public static bool CashboxMultiEdit;
        public static List<bool> CashboxMultiEditUpdateList = new List<bool>() { false, false, false };
        #endif


        /// <summary>
        /// <para>0: SongId | 1: SongLang | 2: SongSinger | 3: SongSongName | 4: SongCreatDate</para>
        /// </summary>
        public static List<string> CashboxDataGridViewSelectList = new List<string>();
        public static string CashboxDataGridViewRestoreCurrentRow = "";
        public static List<string> CashboxDataGridViewRestoreSelectList = new List<string>();

        public static List<string> PlayerRandomVideoList = new List<string>();
        public static bool PlayerInitialized = false;

        public static List<string> CrtchorusSeparateList = new List<string>() { "&", "+" };

        public static List<string> SynonymousWordList = new List<string>()
        {
            "阿|啊", "遊|游", "痴|癡", "姍|姗|珊", "袂|抹|袜|抺|祙", "你|妳", "姊|姐", "未|末", "他|她", "秘|祕", "周|週", "已|己", "雙|双",
            "兩|二", "作|做", "密|蜜", "愈|越", "裡|裏", "腳|脚", "嘆|歎", "煙|菸|烟", "叉|义", "叨|叼", "妝|粧", "鉤|鈎", "沈|沉", "灑|洒",
            "無|嘸|毋|唔", "恒|恆", "那|哪", "甘|咁", "庄|莊", "漂|飄", "什|甚|啥", "決|絕", "麥|嘜", "奈|耐", "果|菓", "的|得|の", "一|ㄧ",
            "逃|鼗", "喲|呦", "藥|葯", "只|衹|祇|袛|祗", "群|羣", "拼|拚", "麼|麽", "支|隻", "鬥|斗", "占,佔", "侯,候"
        };

        public static List<string> SingerTypeStructureList = new List<string>()
        {
            "男,男星,男歌星,男歌手",
            "女,女星,女歌星,女歌手",
            "團,團體,樂團",
            "合唱,對唱",
            "外男,外國男星,外國男歌星,外國男歌手",
            "外女,外國女星,外國女歌星,外國女歌手",
            "外團,外國團體,外國樂團",
            "未知,其他,其他歌星,其他歌手"
        };

        public static List<string> SongAnalysisExceptionSingerList = new List<string>() { "台灣女孩", "梁靜茹", "張懸", "莫文蔚" };
        public static List<string> SongAnalysisExceptionSongList = new List<string>()
        {
            "羅百吉&寶貝|台灣女孩",
            "梁靜茹|兒歌",
            "張懸|兒歌",
            "莫文蔚|台"
        };

        public static DataTable SongMonitorDT = new DataTable();

        public static FileSystemWatcher SongMonitorPath1 = new FileSystemWatcher();
        public static FileSystemWatcher SongMonitorPath2 = new FileSystemWatcher();
        public static FileSystemWatcher SongMonitorPath3 = new FileSystemWatcher();
        public static FileSystemWatcher SongMonitorPath4 = new FileSystemWatcher();
        public static FileSystemWatcher SongMonitorPath5 = new FileSystemWatcher();
        public static FileSystemWatcher[] SongMonitorWatcher =
        {
            SongMonitorPath1,
            SongMonitorPath2,
            SongMonitorPath3,
            SongMonitorPath4,
            SongMonitorPath5
        };

        public static Timer SongMonitorTimer = new Timer();
        public static DateTime SongMonitorSTime = new DateTime();
        public static DateTime SongMonitorETime = new DateTime();

        public static List<string> SongMonitorCreatedList = new List<string>();
        public static List<string> SongMonitorDeletedList = new List<string>();

        public static List<string> CashboxSongLangList = new List<string>() { "國語", "台語", "粵語", "日語", "英語", "客語", "韓語", "兒歌", "其它" };
        public static DateTime CashboxUpdDate = DateTime.Parse("2016/01/01");
        public static DataTable CashboxDT = new DataTable();
        public static bool CashboxFuzzyQuery = true;
        public static string CashboxQueryFilter = "全部";
        public static string CashboxNonSymbolList = "、|。|！|，|．|？";
        public static bool CashboxSynonymousQuery = true;

        public static string RegexChorusSeparate = "[&+、]";

        public static bool FFDShowAudioProcessorRegistered = false;

        public static List<string> DetectUnusedFilePathList = new List<string>();
    }
}

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CrazyKTV_SongMgr.NativeMethods;

namespace CrazyKTV_SongMgr
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            SongQuery_DataGridView.MakeDoubleBuffered(true);
            SongAdd_DataGridView.MakeDoubleBuffered(true);
            SingerMgr_DataGridView.MakeDoubleBuffered(true);
            Cashbox_DataGridView.MakeDoubleBuffered(true);


        }

        // 標題列的程式名稱及版本,其後接目前使用的資料庫 (Common_SetDatabaseFile)
        private string MainForm_Title;

        /// <summary>
        /// Makes <paramref name="songDatabase"/> the 歌庫 and pairs CrazySongMgr with it (<see cref="Db.SongMgrDatabaseFor"/>),
        /// then shows the type and both file names in the title bar.
        /// </summary>
        private void Common_SetDatabaseFile(string songDatabase)
        {
            Global.CrazyktvDatabaseFile = songDatabase;
            Global.CrazyktvSongMgrDatabaseFile = Db.SongMgrDatabaseFor(songDatabase, Global.CrazyktvSongMgrDatabaseFile);
            SongMgrCfg_DBFile_TextBox.Text = songDatabase;
            MainCfg_ExportSQLite_UpdateEnabled();
            this.Text = MainForm_Title + " - 【" + (Db.IsSQLite(songDatabase) ? "SQLite" : "Access") + "】" + Path.GetFileName(songDatabase) + " + " + Path.GetFileName(Global.CrazyktvSongMgrDatabaseFile);
        }

        private string Common_MissingSongMgrDatabaseText(string songMgrDatabase)
        {
            if (Db.IsSQLite(songMgrDatabase))
            {
                return "SQLite 歌庫須搭配 SongMgr\\CrazySongMgr.db 參考資料庫。" + Environment.NewLine + "請先載入 CrazySong.mdb 歌庫,再使用「匯出 SQLite 資料庫」產生。";
            }
            return "請將 CrazySongMgr.mdb 移至 SongMgr 資料夾裡。" + Environment.NewLine + "或是重新下載並執行 CrazyKTV_WebUpdater.exe 來重建所有檔案。";
        }

        
        private void MainForm_Load(object sender, EventArgs e)
        {
            string exePath = Environment.ProcessPath;
            var verInfo = FileVersionInfo.GetVersionInfo(exePath);
            string CurVer = " v" + verInfo.FileMajorPart + "." +
                                   verInfo.FileMinorPart + "." +
                                   verInfo.FileBuildPart;

            if (verInfo.FilePrivatePart > 0) CurVer += "." + verInfo.FilePrivatePart;

            this.Text += CurVer;

            if (CommonFunc.IsAdministrator())
            {
                this.Text += " (系統管理員)";
            }
            MainForm_Title = this.Text;

            if(CommonFunc.CheckFor45PlusVersion()) ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

            ElevatedDragDropManager.Instance.EnableDragDrop(SongAdd_DataGridView.Handle);
            ElevatedDragDropManager.Instance.EnableDragDrop(SongAdd_DragDrop_Label.Handle);
            ElevatedDragDropManager.Instance.ElevatedDragDrop += SongAdd_ElevatedDragDrop;

            ElevatedDragDropManager.Instance.EnableDragDrop(SingerMgr_EditSingerImg_Panel.Handle);
            ElevatedDragDropManager.Instance.ElevatedDragDrop += SingerMgr_EditSingerImg_Panel_ElevatedDragDrop;

            #if DEBUG
            Cashbox_EditMode_CheckBox.Visible = true;
            #else
                Debug_TabPage.Hide();
                MainTabControl.TabPages.Remove(Debug_TabPage);
            #endif

            Global.DPIScalingFactor = Common_GetDPIScalingFactor();
            SongAddResult_SplitContainer.SplitterDistance = Convert.ToInt32(SongAddResult_SplitContainer.SplitterDistance * Global.DPIScalingFactor);

            // 歌庫版本資訊
            if (!File.Exists(Global.CrazyktvSongDBUpdateFile))
            {
                if (!Directory.Exists(Application.StartupPath + @"\SongMgr\Update")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Update");
                CommonFunc.CreateConfigXmlFile(Global.CrazyktvSongDBUpdateFile);
                CommonFunc.SaveConfigXmlFile(Global.CrazyktvSongDBUpdateFile, "SongDBVer", Global.CrazyktvSongDBVer);
            }

            // CrazyKTV 設定
            if (!File.Exists(Global.CrazyktvCfgFile))
            {
                CommonFunc.CreateConfigXmlFile(Global.CrazyktvCfgFile);
            }

            // 載入歌庫設定
            if (!File.Exists(Global.SongMgrCfgFile))
            {
                SettingsFile.Create(Global.SongMgrCfgFile, SongMgrSettings.Defaults());
            }

            SettingsFile cfg = SettingsFile.Load(Global.SongMgrCfgFile);

            // 音量正規化 top-level tab (programmatic – added to MainTabControl)
            InitSongNormVolTab();

            foreach (TabPage MainTabPage in MainTabControl.TabPages)
            {
                MainTabPage.Show();
            }

            foreach (TabPage SongQueryTabPage in SongQuery_TabControl.TabPages)
            {
                SongQueryTabPage.Show();
            }

            foreach (TabPage SongMgrCfgTabPage in SongMgrCfg_TabControl.TabPages)
            {
                SongMgrCfgTabPage.Show();
            }

            foreach (TabPage SongMaintenanceTabPage in SongMaintenance_TabControl.TabPages)
            {
                SongMaintenanceTabPage.Show();
            }

            foreach (TabPage CashboxTabPage in Cashbox_TabControl.TabPages)
            {
                CashboxTabPage.Show();
            }

            if (cfg["CrazyktvSongLangStr"] != "") Global.CrazyktvSongLangList = new List<string>(cfg["CrazyktvSongLangStr"].Split(','));
            if (Global.CrazyktvSongLangList.Count < 10) Global.CrazyktvSongLangList = new List<string>() { "國語", "台語", "粵語", "日語", "英語", "客語", "原住民語", "韓語", "兒歌", "其它" };

            if (cfg["CrazyktvDatabaseFile"] != "") Global.CrazyktvDatabaseFile = cfg["CrazyktvDatabaseFile"];
            Common_SetDatabaseFile(Global.CrazyktvDatabaseFile);

            // 檢查加歌程式參考資料庫檔案
            if (!File.Exists(Global.CrazyktvSongMgrDatabaseFile))
            {
                if (Db.IsSQLite(Global.CrazyktvSongMgrDatabaseFile))
                {
                    // SQLite 歌庫沒有 CrazySongMgr.db 時不改用 .mdb,各頁面維持鎖定,可在歌庫設定改載入其它歌庫
                    MessageBox.Show(Common_MissingSongMgrDatabaseText(Global.CrazyktvSongMgrDatabaseFile), @"SongMgr\CrazySongMgr.db 參考資料庫檔案不存在!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(Application.StartupPath + @"\CrazySongMgr.mdb"))
                {
                    try
                    {
                        File.Move(Application.StartupPath + @"\CrazySongMgr.mdb", Global.CrazyktvSongMgrDatabaseFile);
                    }
                    catch
                    {
                        if (MessageBox.Show(Common_MissingSongMgrDatabaseText(Global.CrazyktvSongMgrDatabaseFile), @"SongMgr\CrazySongMgr.mdb 參考資料庫檔案不存在!", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                        {
                            Application.Exit();
                        }
                    }
                }
                else
                {
                    if (MessageBox.Show(Common_MissingSongMgrDatabaseText(Global.CrazyktvSongMgrDatabaseFile), @"SongMgr\CrazySongMgr.mdb 參考資料庫檔案不存在!", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                    {
                        Application.Exit();
                    }
                }
            }

            if (cfg["SongMgrSupportFormat"] != "") Global.SongMgrSupportFormat = cfg["SongMgrSupportFormat"];
            SongMgrCfg_SupportFormat_TextBox.Text = Global.SongMgrSupportFormat;

            if (cfg["SongMgrDestFolder"] != "") Global.SongMgrDestFolder = cfg["SongMgrDestFolder"];
            SongMgrCfg_DestFolder_TextBox.Text = Global.SongMgrDestFolder;
            
            if (cfg["SongMgrChorusMerge"] != "") Global.SongMgrChorusMerge = cfg["SongMgrChorusMerge"];
            SongMgrCfg_CrtchorusMerge_CheckBox.Checked = bool.Parse(Global.SongMgrChorusMerge);

            if (cfg["SongMgrMaxDigitCode"] != "") Global.SongMgrMaxDigitCode = cfg["SongMgrMaxDigitCode"];
            SongMgrCfg_MaxDigitCode_ComboBox.DataSource = SongMgrCfg.GetMaxDigitCodeList();
            SongMgrCfg_MaxDigitCode_ComboBox.DisplayMember = "Display";
            SongMgrCfg_MaxDigitCode_ComboBox.ValueMember = "Value";
            SongMgrCfg_MaxDigitCode_ComboBox.SelectedValue = int.Parse(Global.SongMgrMaxDigitCode);

            if (cfg["SongMgrLangCode"] != "") { Global.SongMgrLangCode = cfg["SongMgrLangCode"]; SongMgrCfg_LoadSongMgrLangCode(); } else { SongMgrCfg_RefreshSongMgrLangCode(); }

            if (cfg["SongMgrSongType"] != "") Global.SongMgrSongType = cfg["SongMgrSongType"];
            SongMgrCfg_SongType_ListBox.DataSource = SongMgrCfg.GetSongTypeList();
            SongMgrCfg_SongType_ListBox.DisplayMember = "Display";
            SongMgrCfg_SongType_ListBox.ValueMember = "Value";

            if (cfg["SongMgrSongInfoSeparate"] != "") Global.SongMgrSongInfoSeparate = cfg["SongMgrSongInfoSeparate"];
            SongMgrCfg_SongInfoSeparate_ComboBox.DataSource = SongMgrCfg.GetSongInfoSeparateList();
            SongMgrCfg_SongInfoSeparate_ComboBox.DisplayMember = "Display";
            SongMgrCfg_SongInfoSeparate_ComboBox.ValueMember = "Value";
            SongMgrCfg_SongInfoSeparate_ComboBox.SelectedValue = int.Parse(Global.SongMgrSongInfoSeparate);

            if (cfg["SongMgrChorusSeparate"] != "") Global.SongMgrChorusSeparate = cfg["SongMgrChorusSeparate"];
            SongMgrCfg_CrtchorusSeparate_ComboBox.DataSource = SongMgrCfg.GetCrtchorusSeparateList();
            SongMgrCfg_CrtchorusSeparate_ComboBox.DisplayMember = "Display";
            SongMgrCfg_CrtchorusSeparate_ComboBox.ValueMember = "Value";
            SongMgrCfg_CrtchorusSeparate_ComboBox.SelectedValue = int.Parse(Global.SongMgrChorusSeparate);

            if (cfg["SongAddDefaultSongLang"] != "") Global.SongAddDefaultSongLang = cfg["SongAddDefaultSongLang"];
            SongAdd_DefaultSongLang_ComboBox.DataSource = SongAdd.GetDefaultSongInfo("DefaultSongLang", false, false);
            SongAdd_DefaultSongLang_ComboBox.DisplayMember = "Display";
            SongAdd_DefaultSongLang_ComboBox.ValueMember = "Value";
            SongAdd_DefaultSongLang_ComboBox.SelectedValue = int.Parse(Global.SongAddDefaultSongLang);

            if (cfg["SongAddDefaultSingerType"] != "") Global.SongAddDefaultSingerType = cfg["SongAddDefaultSingerType"];
            SongAdd_DefaultSingerType_ComboBox.DataSource = SongAdd.GetDefaultSongInfo("DefaultSingerType", false, false);
            SongAdd_DefaultSingerType_ComboBox.DisplayMember = "Display";
            SongAdd_DefaultSingerType_ComboBox.ValueMember = "Value";
            SongAdd_DefaultSingerType_ComboBox.SelectedValue = int.Parse(Global.SongAddDefaultSingerType);

            if (cfg["SongAddDefaultSongTrack"] != "") Global.SongAddDefaultSongTrack = cfg["SongAddDefaultSongTrack"];
            SongAdd_DefaultSongTrack_ComboBox.DataSource = SongAdd.GetDefaultSongInfo("DefaultSongTrack", false, false);
            SongAdd_DefaultSongTrack_ComboBox.DisplayMember = "Display";
            SongAdd_DefaultSongTrack_ComboBox.ValueMember = "Value";
            SongAdd_DefaultSongTrack_ComboBox.SelectedValue = int.Parse(Global.SongAddDefaultSongTrack);

            if (cfg["SongAddDefaultSongType"] != "") Global.SongAddDefaultSongType = cfg["SongAddDefaultSongType"];
            SongAdd_DefaultSongType_ComboBox.DataSource = SongAdd.GetDefaultSongInfo("DefaultSongType", false, false);
            SongAdd_DefaultSongType_ComboBox.DisplayMember = "Display";
            SongAdd_DefaultSongType_ComboBox.ValueMember = "Value";
            SongAdd_DefaultSongType_ComboBox.SelectedValue = int.Parse(Global.SongAddDefaultSongType);

            if (cfg["SongAddDefaultSongVolume"] != "") Global.SongAddDefaultSongVolume = cfg["SongAddDefaultSongVolume"];
            SongAdd_DefaultSongVolume_TextBox.Text = Global.SongAddDefaultSongVolume;

            if (cfg["SongAddSpecialStr"] != "") Global.SongAddSpecialStr = cfg["SongAddSpecialStr"];
            SongAdd_SpecialStr_ListBox.DataSource = SongAdd.GetDefaultSongInfo("SpecialStr", false, false);
            SongAdd_SpecialStr_ListBox.DisplayMember = "Display";
            SongAdd_SpecialStr_ListBox.ValueMember = "Value";

            if (cfg["SongMgrFolderStructure"] != "") Global.SongMgrFolderStructure = cfg["SongMgrFolderStructure"];
            SongMgrCfg_FolderStructure_ComboBox.DataSource = SongMgrCfg.GetFolderStructureList();
            SongMgrCfg_FolderStructure_ComboBox.DisplayMember = "Display";
            SongMgrCfg_FolderStructure_ComboBox.ValueMember = "Value";
            SongMgrCfg_FolderStructure_ComboBox.SelectedValue = Global.SongMgrFolderStructure;

            if (cfg["SongMgrFileStructure"] != "") Global.SongMgrFileStructure = cfg["SongMgrFileStructure"];
            SongMgrCfg_FileStructure_ComboBox.DataSource = SongMgrCfg.GetFileStructureList();
            SongMgrCfg_FileStructure_ComboBox.DisplayMember = "Display";
            SongMgrCfg_FileStructure_ComboBox.ValueMember = "Value";
            SongMgrCfg_FileStructure_ComboBox.SelectedValue = Global.SongMgrFileStructure;

            if (cfg["SongMgrSongTrackMode"] != "") Global.SongMgrSongTrackMode = cfg["SongMgrSongTrackMode"];
            SongMgrCfg_SongTrackMode_CheckBox.Checked = bool.Parse(Global.SongMgrSongTrackMode);

            if (cfg["SongAddSongIdentificationMode"] != "") Global.SongAddSongIdentificationMode = cfg["SongAddSongIdentificationMode"];
            SongAdd_SongIdentificationMode_ComboBox.DataSource = SongAdd.GetSongIdentificationModeList();
            SongAdd_SongIdentificationMode_ComboBox.DisplayMember = "Display";
            SongAdd_SongIdentificationMode_ComboBox.ValueMember = "Value";
            SongAdd_SongIdentificationMode_ComboBox.SelectedValue = Global.SongAddSongIdentificationMode;

            if (cfg["SongAddDupSongMode"] != "") Global.SongAddDupSongMode = cfg["SongAddDupSongMode"];
            SongAdd_DupSongMode_ComboBox.DataSource = SongAdd.GetDupSongModeList();
            SongAdd_DupSongMode_ComboBox.DisplayMember = "Display";
            SongAdd_DupSongMode_ComboBox.ValueMember = "Value";
            SongAdd_DupSongMode_ComboBox.SelectedValue = Global.SongAddDupSongMode;

            if (cfg["CrazyktvSongLangKeyWord"] != "")
            {
                Global.CrazyktvSongLangKeyWordList = new List<string>(cfg["CrazyktvSongLangKeyWord"].Split('|'));
            }
            else
            {
                string str = CommonFunc.LoadConfigXmlFile(Global.SongMgrCfgFile, "CrazyktvSongLangIDStr", null, false);
                Global.CrazyktvSongLangKeyWordList = new List<string>(str.Split('*'));
                if (Global.CrazyktvSongLangKeyWordList.Count < 10)
                {
                    Global.CrazyktvSongLangKeyWordList = new List<string>() { "國語,國", "台語,台,閩南,閩", "粵語,粵,廣東", "日語,日文,日", "英語,英文,英", "客語,客", "原住民語,民謠", "韓語,韓", "兒歌,兒", "其它" };
                }
                else
                {
                    CommonFunc.SaveConfigXmlFile(Global.SongMgrCfgFile, "CrazyktvSongLangKeyWord", string.Join("|", Global.CrazyktvSongLangKeyWordList));
                    CommonFunc.RemoveConfigXmlFile(Global.SongMgrCfgFile, "CrazyktvSongLangIDStr");
                }
            }

            if (cfg["SongMgrBackupRemoveSong"] != "") Global.SongMgrBackupRemoveSong = cfg["SongMgrBackupRemoveSong"];
            SongMgrCfg_BackupRemoveSong_CheckBox.Checked = bool.Parse(Global.SongMgrBackupRemoveSong);

            if (cfg["SongAddEngSongNameFormat"] != "") Global.SongAddEngSongNameFormat = cfg["SongAddEngSongNameFormat"];
            SongAdd_EngSongNameFormat_CheckBox.Checked = bool.Parse(Global.SongAddEngSongNameFormat);

            if (cfg["MainCfgAlwaysOnTop"] != "") Global.MainCfgAlwaysOnTop = cfg["MainCfgAlwaysOnTop"];
            MainCfg_AlwaysOnTop_CheckBox.Checked = bool.Parse(Global.MainCfgAlwaysOnTop);

            if (!string.IsNullOrEmpty(cfg["CrazyktvSongDBVer"])) Global.CrazyktvSongDBVer = cfg["CrazyktvSongDBVer"];

            if (cfg["MainCfgHideSongDBConverterTabPage"] != "") Global.MainCfgHideSongDBConverterTabPage = cfg["MainCfgHideSongDBConverterTabPage"];
            MainCfg_HideSongDBConvTab_CheckBox.Checked = bool.Parse(Global.MainCfgHideSongDBConverterTabPage);

            if (cfg["MainCfgHideSongAddResultTabPage"] != "") Global.MainCfgHideSongAddResultTabPage = cfg["MainCfgHideSongAddResultTabPage"];
            MainCfg_HideSongAddResultTab_CheckBox.Checked = bool.Parse(Global.MainCfgHideSongAddResultTabPage);

            if (cfg["MainCfgHideSongLogTabPage"] != "") Global.MainCfgHideSongLogTabPage = cfg["MainCfgHideSongLogTabPage"];
            MainCfg_HideSongLogTab_CheckBox.Checked = bool.Parse(Global.MainCfgHideSongLogTabPage);

            if (cfg["MainCfgBackupRemoveSongDays"] != "") Global.MainCfgBackupRemoveSongDays = cfg["MainCfgBackupRemoveSongDays"];
            MainCfg_BackupRemoveSongDays_ComboBox.DataSource = MainCfg.GetBackupRemoveSongDaysList();
            MainCfg_BackupRemoveSongDays_ComboBox.DisplayMember = "Display";
            MainCfg_BackupRemoveSongDays_ComboBox.ValueMember = "Value";
            MainCfg_BackupRemoveSongDays_ComboBox.SelectedValue = Global.MainCfgBackupRemoveSongDays;

            if (cfg["SingerMgrDefaultSingerDataTable"] != "") Global.SingerMgrDefaultSingerDataTable = cfg["SingerMgrDefaultSingerDataTable"];
            SingerMgr_DefaultSingerDataTable_ComboBox.DataSource = SingerMgr.GetDefaultSingerDataTableList();
            SingerMgr_DefaultSingerDataTable_ComboBox.DisplayMember = "Display";
            SingerMgr_DefaultSingerDataTable_ComboBox.ValueMember = "Value";

            #if DEBUG
                SingerMgr_DefaultSingerDataTable_ComboBox.SelectedValue = (SingerTable.Default.IsLibrary) ? 1 : 2;
            #else
                Global.SingerMgrDefaultSingerDataTable = "ktv_Singer";
                SingerMgr_DefaultSingerDataTable_ComboBox.SelectedValue = 1;
                SingerMgr_DefaultSingerDataTable_ComboBox.Enabled = false;
            #endif

            if (cfg["DBVerEnableDBVerUpdate"] != "") Global.DBVerEnableDBVerUpdate = cfg["DBVerEnableDBVerUpdate"];
            SongMaintenance_EnableDBVerUpdate_CheckBox.Checked = bool.Parse(Global.DBVerEnableDBVerUpdate);

            if (cfg["MainCfgUIScale"] != "") Global.MainCfgUIScale = cfg["MainCfgUIScale"];
            MainCfg_UIScale_ComboBox.DataSource = MainCfg.GetUIScaleList();
            MainCfg_UIScale_ComboBox.DisplayMember = "Display";
            MainCfg_UIScale_ComboBox.ValueMember = "Value";
            MainCfg_UIScale_ComboBox.SelectedValue = Global.MainCfgUIScale;

            if (cfg["SongMaintenanceReplayGainVolume"] != "") Global.SongMaintenanceReplayGainVolume = cfg["SongMaintenanceReplayGainVolume"];
            SongMaintenance_ReplayGain_TextBox.Text = Global.SongMaintenanceReplayGainVolume;

            if (cfg["SongAddEnableVolumeDetect"] != "") Global.SongAddEnableVolumeDetect = cfg["SongAddEnableVolumeDetect"];
            SongAdd_EnableVolumeDetect_CheckBox.Checked = bool.Parse(Global.SongAddEnableVolumeDetect);

            if (cfg["SongAddUseCustomSongID"] != "") Global.SongAddUseCustomSongID = cfg["SongAddUseCustomSongID"];
            SongAdd_UseCustomSongID_CheckBox.Checked = bool.Parse(Global.SongAddUseCustomSongID);

            if (cfg["MainCfgUIFont"] != "") Global.MainCfgUIFont = cfg["MainCfgUIFont"];
            MainCfg_UIFont_ComboBox.DataSource = MainCfg.GetUIFontList();
            MainCfg_UIFont_ComboBox.DisplayMember = "Display";
            MainCfg_UIFont_ComboBox.ValueMember = "Value";
            MainCfg_UIFont_ComboBox.Text = Global.MainCfgUIFont;

            if (cfg["MainCfgEnableUIScale"] != "") Global.MainCfgEnableUIScale = cfg["MainCfgEnableUIScale"];
            MainCfg_EnableUIScale_CheckBox.Checked = bool.Parse(Global.MainCfgEnableUIScale);

            if (cfg["SongMgrCustomSingerTypeStructure"] != "") Global.SongMgrCustomSingerTypeStructure = cfg["SongMgrCustomSingerTypeStructure"];
            SongMgrCfg_SetCustomSingerTypeStructureCbox();

            if (cfg["MainCfgEnableAutoUpdate"] != "") Global.MainCfgEnableAutoUpdate = cfg["MainCfgEnableAutoUpdate"];
            MainCfg_EnableAutoUpdate_CheckBox.Checked = bool.Parse(Global.MainCfgEnableAutoUpdate);

            if (cfg["SongAddEnableConvToTC"] != "") Global.SongAddEnableConvToTC = cfg["SongAddEnableConvToTC"];
            SongAdd_EnableConvToTC_CheckBox.Checked = bool.Parse(Global.SongAddEnableConvToTC);

            if (cfg["SongMgrEnableMonitorFolders"] != "") Global.SongMgrEnableMonitorFolders = cfg["SongMgrEnableMonitorFolders"];
            SongMgrCfg_MonitorFolders_CheckBox.Checked = bool.Parse(Global.SongMgrEnableMonitorFolders);

            if (cfg["SongMgrMonitorFolders"] != "") Global.SongMgrMonitorFoldersList = new List<string>(cfg["SongMgrMonitorFolders"].Split(','));

            if (cfg["SingerMgrSyncSongSinger"] != "") Global.SingerMgrSyncSongSinger = cfg["SingerMgrSyncSongSinger"];
            SingerMgr_EditSyncSongSinger_CheckBox.Checked = bool.Parse(Global.SingerMgrSyncSongSinger);

            if (cfg["SongMgrSingerGroup"] != "") Global.SongMgrSingerGroup = cfg["SongMgrSingerGroup"];
            SongMgrCfg_SingerGroup_ListBox.DataSource = SongMgrCfg.GetSingerGroupList();
            SongMgrCfg_SingerGroup_ListBox.DisplayMember = "Display";
            SongMgrCfg_SingerGroup_ListBox.ValueMember = "Value";

            if (cfg["MainCfgHideApplyCashboxIdButton"] != "") Global.MainCfgHideApplyCashboxIdButton = cfg["MainCfgHideApplyCashboxIdButton"];
            MainCfg_HideApplyCashboxIdButton_CheckBox.Checked = bool.Parse(Global.MainCfgHideApplyCashboxIdButton);

            if (cfg["MainCfgUICustomScale"] != "") Global.MainCfgUICustomScale = cfg["MainCfgUICustomScale"];
            MainCfg_UIScale_TextBox.Text = Global.MainCfgUICustomScale;

            if (cfg["MainCfgBackupDB"] != "") Global.MainCfgBackupDB = cfg["MainCfgBackupDB"];
            MainCfg_BackupDB_CheckBox.Checked = bool.Parse(Global.MainCfgBackupDB);

            if (cfg["MainCfgBackupDBPath"] != "") Global.MainCfgBackupDBPath = cfg["MainCfgBackupDBPath"];
            MainCfg_BackupDB_TextBox.Text = Global.MainCfgBackupDBPath;

            if (cfg["MainCfgPlayerCore"] != "") Global.MainCfgPlayerCore = "1";
            MainCfg_PlayerCore_RadioButton1.Checked = true;

            if (cfg["MainCfgPlayerOutput"] != "") Global.MainCfgPlayerOutput = cfg["MainCfgPlayerOutput"];
            if (Global.MainCfgPlayerOutput == "1")
            {
                MainCfg_PlayerOutput_RadioButton1.Checked = true;
            }
            else
            {
                MainCfg_PlayerOutput_RadioButton2.Checked = true;
            }

            if (cfg["MainCfgPlayerEnableAudioCompressor"] != "") Global.MainCfgPlayerEnableAudioCompressor = cfg["MainCfgPlayerEnableAudioCompressor"];
            MainCfg_PlayerEnableAudioCompressor_CheckBox.Checked = bool.Parse(Global.MainCfgPlayerEnableAudioCompressor);

            if (cfg["MainCfgPlayerEnableAudioProcessor"] != "") Global.MainCfgPlayerEnableAudioProcessor = cfg["MainCfgPlayerEnableAudioProcessor"];
            MainCfg_PlayerEnableAudioProcessor_CheckBox.Checked = bool.Parse(Global.MainCfgPlayerEnableAudioProcessor);

            if (cfg["MainCfgPlayerDefaultVolume"] != "") Global.MainCfgPlayerDefaultVolume = cfg["MainCfgPlayerDefaultVolume"];
            MainCfg_PlayerDefaultVolume_TextBox.Text = Global.MainCfgPlayerDefaultVolume;

            if (cfg["SongMgrSongAddMode"] != "") Global.SongMgrSongAddMode = cfg["SongMgrSongAddMode"];
            SongMgrCfg_SongAddMode_ComboBox.DataSource = SongMgrCfg.GetSongAddModeList();
            SongMgrCfg_SongAddMode_ComboBox.DisplayMember = "Display";
            SongMgrCfg_SongAddMode_ComboBox.ValueMember = "Value";
            SongMgrCfg_SongAddMode_ComboBox.SelectedValue = int.Parse(Global.SongMgrSongAddMode);

            // 建立歌曲操作記錄資料表
            Global.SongLog.Reset();

            // 檢查解碼器登錄狀態
            RegistryKey regKey = Registry.ClassesRoot.OpenSubKey("CLSID\\{083863F1-70DE-11D0-BD40-00A0C911CE86}\\Instance\\{B86F6BEE-E7C0-4D03-8D52-5B4430CF6C88}");
            Global.FFDShowAudioProcessorRegistered = (regKey == null) ? false : true;
            MainCfg_PlayerEnableAudioProcessor_CheckBox.Enabled = (Global.MainCfgPlayerCore == "1" && Global.FFDShowAudioProcessorRegistered && File.Exists(Application.StartupPath + @"\Codec_x86\ffdshow\ffdshow.ax")) ? true : false;
            MainCfg_PlayerSetAudioProcessor_Button.Enabled = (Global.MainCfgPlayerCore == "1" && Global.MainCfgPlayerEnableAudioProcessor == "True" && Global.FFDShowAudioProcessorRegistered && File.Exists(Application.StartupPath + @"\Codec_x86\ffdshow\ffdshow.ax")) ? true : false;
            MainCfg_PlayerRegAudioProcessor_Button.Enabled = (Global.MainCfgPlayerCore == "1" && !Global.FFDShowAudioProcessorRegistered && File.Exists(Application.StartupPath + @"\Codec_x86\ffdshow\ffdshow.ax")) ? true : false;
            if (!Global.FFDShowAudioProcessorRegistered)
            {
                if (Global.MainCfgPlayerEnableAudioProcessor == "True")
                {
                    MainCfg_PlayerEnableAudioProcessor_CheckBox.Checked = false;
                }
            }

            // 檢查程式更新
            Common_CheckSongMgrVer();

            // 檢查資料庫檔案是否為舊版資料庫
            SongDBUpdate_CheckDatabaseFile();

            // 初始化所需資料
            SongMgrDB_CheckDatabaseFile();
            Common_SwitchDBVerErrorUI(Global.SongMgrDBVerErrorUIStatus);
            Task.Factory.StartNew(() => Common_InitializeSongData(true, true, true, true, true, true));

            // 歌曲查詢 - 載入下拉選單清單及設定
            SongQuery_QueryType_ComboBox.DataSource = SongQuery.GetSongQueryTypeList();
            SongQuery_QueryType_ComboBox.DisplayMember = "Display";
            SongQuery_QueryType_ComboBox.ValueMember = "Value";
            SongQuery_QueryType_ComboBox.SelectedValue = 1;
            SongQuery_QueryValue_TextBox.ImeMode = ImeMode.OnHalf;
            SongQuery_InitializeYoutubeEvents();

            SongQuery_QueryFilter_ComboBox.DataSource = SongQuery.GetSongQueryFilterList(Global.CrazyktvSongLangList);
            SongQuery_QueryFilter_ComboBox.DisplayMember = "Display";
            SongQuery_QueryFilter_ComboBox.ValueMember = "Value";
            SongQuery_QueryFilter_ComboBox.SelectedValue = 1;

            SongQuery_FuzzyQuery_CheckBox.Checked = bool.Parse(Global.SongQueryFuzzyQuery);
            SongQuery_SynonymousQuery_CheckBox.Checked = Global.SongQuerySynonymousQuery;

            SongQuery_ExceptionalQuery_ComboBox.DataSource = SongQuery.GetSongQueryExceptionalList();
            SongQuery_ExceptionalQuery_ComboBox.DisplayMember = "Display";
            SongQuery_ExceptionalQuery_ComboBox.ValueMember = "Value";
            SongQuery_ExceptionalQuery_ComboBox.SelectedValue = 1;

            // 歌手管理 - 載入下拉選單清單及設定
            SingerMgr_QueryType_ComboBox.DataSource = SingerMgr.GetSingerTypeList(false);
            SingerMgr_QueryType_ComboBox.DisplayMember = "Display";
            SingerMgr_QueryType_ComboBox.ValueMember = "Value";
            SingerMgr_QueryType_ComboBox.SelectedValue = 1;
            SingerMgr_QueryValue_TextBox.ImeMode = ImeMode.OnHalf;

            SingerMgr_SingerAddType_ComboBox.DataSource = SingerMgr.GetSingerTypeList(false);
            SingerMgr_SingerAddType_ComboBox.DisplayMember = "Display";
            SingerMgr_SingerAddType_ComboBox.ValueMember = "Value";
            SingerMgr_SingerAddType_ComboBox.SelectedValue = 1;
            SingerMgr_SingerAddName_TextBox.ImeMode = ImeMode.OnHalf;

            SingerMgr_SingerLastName_ComboBox.DataSource = SingerMgr.GetSingerLastNameList();
            SingerMgr_SingerLastName_ComboBox.DisplayMember = "Display";
            SingerMgr_SingerLastName_ComboBox.ValueMember = "Value";
            SingerMgr_SingerLastName_ComboBox.SelectedValue = 1;

            // 歌庫維護 - 載入下拉選單清單及設定
            SongMaintenance_Favorite_UpdateNewsong_ComboBox.DataSource = SongMaintenance.GetNewsongDateList();
            SongMaintenance_Favorite_UpdateNewsong_ComboBox.DisplayMember = "Display";
            SongMaintenance_Favorite_UpdateNewsong_ComboBox.ValueMember = "Value";
            SongMaintenance_Favorite_UpdateNewsong_ComboBox.SelectedValue = 1;

            // 錢櫃資料 - 載入下拉選單清單及設定
            Cashbox_QueryType_ComboBox.DataSource = Cashbox.GetQueryTypeList();
            Cashbox_QueryType_ComboBox.DisplayMember = "Display";
            Cashbox_QueryType_ComboBox.ValueMember = "Value";
            Cashbox_QueryType_ComboBox.SelectedValue = 1;
            Cashbox_QueryValue_TextBox.ImeMode = ImeMode.OnHalf;

            Cashbox_QueryFilter_ComboBox.DataSource = Cashbox.GetQueryFilterList(Global.CashboxSongLangList);
            Cashbox_QueryFilter_ComboBox.DisplayMember = "Display";
            Cashbox_QueryFilter_ComboBox.ValueMember = "Value";
            Cashbox_QueryFilter_ComboBox.SelectedValue = 1;

            Cashbox_OtherQuery_ComboBox.DataSource = Cashbox.GetOtherQueryList();
            Cashbox_OtherQuery_ComboBox.DisplayMember = "Display";
            Cashbox_OtherQuery_ComboBox.ValueMember = "Value";
            Cashbox_OtherQuery_ComboBox.SelectedValue = 1;

            Cashbox_DateQuery_ComboBox.DataSource = Cashbox.GetDateQueryList();
            Cashbox_DateQuery_ComboBox.DisplayMember = "Display";
            Cashbox_DateQuery_ComboBox.ValueMember = "Value";
            Cashbox_DateQuery_ComboBox.SelectedValue = 1;

            Cashbox_FuzzyQuery_CheckBox.Checked = Global.CashboxFuzzyQuery;
            Cashbox_SynonymousQuery_CheckBox.Checked = Global.CashboxSynonymousQuery;

            // 歌庫轉換 - 載入下拉選單清單
            SongDBConverter_SrcDBType_ComboBox.DataSource = SongDBConverter.GetSrcDBTypeList();
            SongDBConverter_SrcDBType_ComboBox.DisplayMember = "Display";
            SongDBConverter_SrcDBType_ComboBox.ValueMember = "Value";
            SongDBConverter_SrcDBType_ComboBox.SelectedValue = 1;

            // 歌庫轉換 - 載入說明
            SongDBConverter_SetRtfText("來源資料庫: ", "請選擇你要轉換的點歌軟體歌曲資料庫檔案。" + Environment.NewLine);
            SongDBConverter_SetRtfText("資料庫類型: ", "請選擇你要轉換的點歌軟體資料庫類型。" + Environment.NewLine);
            SongDBConverter_SetRtfText("目的資料庫: ", "請選擇本工具所附的 CrazySongEmpty.mdb 空白資料庫檔案。" + Environment.NewLine);

            // Debug - 載入下拉選單清單
            #if DEBUG
                Debug_ConfigData_ComboBox.DataSource = Debug.GetConfigDataList();
                Debug_ConfigData_ComboBox.DisplayMember = "Display";
                Debug_ConfigData_ComboBox.ValueMember = "Value";
                Debug_ConfigData_ComboBox.SelectedValue = 1;
            #endif

            // 縮放界面
            Common_ScalingUI();

            SpinWait.SpinUntil(() => Global.DatabaseUpdateFinished == true);

            // 歌庫監視
            SongMgrCfg_MonitorFolders_CheckBox.Enabled = false;
            foreach (string MonitorFolder in Global.SongMgrMonitorFoldersList)
            {
                if (MonitorFolder != "") SongMgrCfg_MonitorFolders_CheckBox.Enabled = true;
            }
            if (!SongMgrCfg_MonitorFolders_CheckBox.Enabled) SongMgrCfg_MonitorFolders_CheckBox.Checked = false;
            SongMonitor_CheckCurSong();

            if (!Global.CrazyktvDatabaseStatus)
            {
                if (MainTabControl.SelectedIndex != 4) MainTabControl.SelectedIndex = 3;
            }
            else
            {
                MainTabControl_SelectedIndexChanged(new TabControl(), new EventArgs());
            }
            Global.SongMgrInitializeStatus = true;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File.Exists(Global.CrazyktvDatabaseFile))
            {
                try
                {
                    Db.Compact(Global.CrazyktvDatabaseFile);
                    Db.Compact(Global.CrazyktvSongMgrDatabaseFile);
                    if (Global.MainCfgBackupDB == "True" && Directory.Exists(Global.MainCfgBackupDBPath))
                    {
                        try
                        {
                            File.Copy(Global.CrazyktvDatabaseFile, Global.MainCfgBackupDBPath + @"\" + Path.GetFileName(Global.CrazyktvDatabaseFile), true);
                        } catch { }
                    }
                }
                catch
                {
                    if (Global.MainCfgBackupDB == "True" && Directory.Exists(Global.MainCfgBackupDBPath))
                    {
                        try
                        {
                            File.Copy(Global.CrazyktvDatabaseFile, Global.MainCfgBackupDBPath + @"\" + Path.GetFileName(Global.CrazyktvDatabaseFile), true);
                        }
                        catch { }
                    }
                }
            }
        }

        // 頁籤切換處理
        private void MainTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label[] Tooltip_Label =
            {
                SongQuery_QueryStatus_Label,
                SongAdd_Tooltip_Label,
                SingerMgr_Tooltip_Label,
                SongMgrCfg_Tooltip_Label,
                SongMaintenance_Tooltip_Label,
                Cashbox_QueryStatus_Label
            };

            int i = -1;

            switch (MainTabControl.SelectedTab.Name)
            {
                case "SongQuery_TabPage":
                    i = 0;
                    SongQuery_QueryValue_TextBox.Focus();
                    break;
                case "SongAdd_TabPage":
                    i = 1;
                    break;
                case "SingerMgr_TabPage":
                    i = 2;
                    break;
                case "SongMgrCfg_TabPage":
                    i = 3;
                    break;
                case "SongMaintenance_TabPage":
                    i = 4;
                    switch (SongMaintenance_TabControl.SelectedTab.Name)
                    {
                        case "SongMaintenance_CustomLang_TabPage":
                        case "SongMaintenance_FFmpeg_TabPage":
                        case "SongMaintenance_DBVer_TabPage":
                            if (SongMaintenance_TabControl.Enabled == true) SongMaintenance_Save_Button.Enabled = true;
                            break;
                        default:
                            SongMaintenance_Save_Button.Enabled = false;
                            break;
                    }
                    break;
                case "Cashbox_TabPage":
                    i = 5;
                    Cashbox_QueryValue_TextBox.Focus();
                    break;
                case "NormVol_TabPage":
                    // Refresh 基準音量 label in case the user changed it in SongMaintenance
                    NormVol_BaseVol_Label.Text = Global.SongMaintenanceReplayGainVolume;
                    NormVol_Singer_TextBox.Focus();
                    break;
            }

            if (i >= 0)
            {
                if (!Global.CrazyktvDatabaseStatus)
                {
                    if (!File.Exists(Global.CrazyktvDatabaseFile)) Tooltip_Label[i].Text = "CrazyKTV 資料庫檔案不存在!";
                    else if (!File.Exists(Global.CrazyktvSongMgrDatabaseFile)) Tooltip_Label[i].Text = "加歌程式參考資料庫檔案不存在!";
                    else if (Global.CrazyktvDatabaseError) Tooltip_Label[i].Text = "資料庫檔案不是 CrazyKTV 資料庫!";
                    else if (Global.SongMgrDatabaseError) Tooltip_Label[i].Text = "加歌程式參考資料庫不是最新版本!";
                    else if (Global.CrazyktvDatabaseIsOld) Tooltip_Label[i].Text = "資料庫檔案為舊版本!";
                    else if (!Global.CrazyktvDatabaseMaxDigitCode) Tooltip_Label[i].Text = "歌庫編碼混雜 5 及 6 位數編碼!";
                    else if (!Directory.Exists(Global.SongMgrDestFolder)) Tooltip_Label[i].Text = "請先設定歌庫資料夾!";
                }
                else
                {
                    List<string> ErrorTextList = new List<string>()
                    {
                        "CrazyKTV 資料庫檔案不存在!",
                        "加歌程式參考資料庫檔案不存在!",
                        "資料庫檔案不是 CrazyKTV 資料庫!",
                        "加歌程式參考資料庫不是最新版本!",
                        "資料庫檔案為舊版本!",
                        "歌庫編碼混雜 5 及 6 位數編碼!",
                        "請先設定歌庫資料夾!",
                        "資料初始化已完成。"
                    };
                    if (ErrorTextList.IndexOf(Tooltip_Label[i].Text) >= 0) Tooltip_Label[i].Text = "";
                    ErrorTextList.Clear();
                }
            }
        }

        private void MainForm_VisibleChanged(object sender, EventArgs e)
        {
            if (Global.PlayerTrackUpdate != null)
            {
                PlayerTrackUpdate update = Global.PlayerTrackUpdate;
                int i = Convert.ToInt32(update.RowIndex);
                switch (update.Source)
                {
                    case "SongQuery":
                        SongRecord song = SongRecord.FromQueryGridRow(SongQuery_DataGridView.Rows[i]);
                        string OldSongId = song.Id;
                        song.OldId = song.Id;
                        song.SingerType = SongQuery_DataGridView.Rows[i].Cells["Song_SingerType"].Value.ToString();
                        song.Track = update.Track;

                        List<SongRecord> UpdateList = new List<SongRecord>() { song };

                        SongQuery.CreateSongDataTable();
                        Common_SwitchSetUI(false);

                        Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                        Global.SongQueryDataGridViewRestoreSelectList = new List<string>();
                        Global.SongQueryDataGridViewRestoreCurrentRow = OldSongId;
                        SongQuery_DataGridView.Sorted -= new EventHandler(SongQuery_DataGridView_Sorted);
                        SongQuery_QueryStatus_Label.Text = "正在更新音軌資料,請稍待...";

                        using (DataTable UpdateDT = (DataTable)SongQuery_DataGridView.DataSource)
                        {
                            var tasks = new List<Task>()
                            {
                                Task.Factory.StartNew(() => SongQuery_SongUpdate(UpdateList, UpdateDT, false))
                            };

                            Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                            {
                                this.BeginInvoke((Action)delegate()
                                {
                                    SongQuery_DataGridView.Sorted += new EventHandler(SongQuery_DataGridView_Sorted);
                                    SongQuery_DataGridView_Sorted(new object(), new EventArgs());
                                    SongQuery_QueryStatus_Label.Text = "已成功更新 " + Global.TotalList[0] + " 首歌曲的音軌資料。";
                                    Common_SwitchSetUI(true);
                                });
                                UpdateList.Clear();
                                SongQuery.DisposeSongDataTable();
                            });
                        }
                        break;
                    case "SongQueryEdit":
                        SongQuery_EditSongTrack_ComboBox.SelectedValue = update.Track;
                        break;
                    case "SongAdd":
                        SongAdd_EditSongTrack_ComboBox.SelectedValue = update.Track;
                        break;
                }
                Global.PlayerTrackUpdate = null;
            }
        }

        private void SongMgrCfg_TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (SongMgrCfg_TabControl.SelectedTab.Name)
            {
                case "SongMgrCfg_SongStructure_TabPage":
                case "SongMgrCfg_CustomStructure_TabPage":
                    SongMgrCfg_UpdateStructure_Button.Visible = true;
                    break;
                default:
                    SongMgrCfg_UpdateStructure_Button.Visible = false;
                    break;
            }
        }


































































































































































































































        // End
    }
}

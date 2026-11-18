using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    public partial class MainForm : Form
    {
        private void SongMgrDB_CheckDatabaseFile()
        {
            Global.SongMgrDatabaseError = false;

            if (File.Exists(Global.CrazyktvSongMgrDatabaseFile))
            {
                List<string> TableList = new List<string>() { "ktv_AllSinger", "ktv_Cashbox", "ktv_Phonetics", "ktv_SongMgr", "ktv_Version" };
                List<string> SongMgrDBTableList = new List<string>(Db.GetTableList(Global.CrazyktvSongMgrDatabaseFile, ""));
                foreach (string TableName in TableList)
                {
                    if (SongMgrDBTableList.IndexOf(TableName) < 0)
                    {
                        Global.SongMgrDatabaseError = true;
                        break;
                    }
                }
                TableList.Clear();
                SongMgrDBTableList.Clear();
            }
        }

        /// <summary>
        /// Checks both database files and the 歌庫資料夾. Sets Global.CrazyktvDatabaseError / SongMgrDatabaseError (tables
        /// missing), CrazyktvDatabaseIsOld (資料庫結構更新 needed, see <see cref="SchemaUpgrade.Check"/>) and
        /// CrazyktvDatabaseStatus (ready to use).
        /// </summary>
        private (bool DatabaseFiles, bool DestFolder) GetCrazyktvDatabaseStatus()
        {
            Global.CrazyktvDatabaseStatus = false;
            Global.CrazyktvDatabaseIsOld = false;
            Global.CrazyktvDatabaseError = false;
            Global.SongMgrDatabaseError = false;
            Global.CrazyktvDatabaseMaxDigitCode = true;

            bool databaseFiles = File.Exists(Global.CrazyktvDatabaseFile) && File.Exists(Global.CrazyktvSongMgrDatabaseFile);
            if (databaseFiles)
            {
                List<string> songTables = Db.GetTableList(Global.CrazyktvDatabaseFile, "");
                string[] requiredSongTables = { "ktv_Favorite", "ktv_Langauage", "ktv_Phonetics", "ktv_Remote", "ktv_Singer", "ktv_SingerName", "ktv_Song", "ktv_Swan", "ktv_User", "ktv_Words" };
                Global.CrazyktvDatabaseError = requiredSongTables.Any(t => !songTables.Contains(t));

                List<string> songMgrTables = Db.GetTableList(Global.CrazyktvSongMgrDatabaseFile, "");
                string[] requiredSongMgrTables = { "ktv_AllSinger", "ktv_Cashbox", "ktv_Phonetics", "ktv_SongMgr", "ktv_Version" };
                Global.SongMgrDatabaseError = requiredSongMgrTables.Any(t => !songMgrTables.Contains(t));

                if (!Global.CrazyktvDatabaseError) Global.CrazyktvDatabaseIsOld = SchemaUpgrade.Check(Global.CrazyktvDatabaseFile).Count > 0;
            }

            // 歌庫資料夾
            bool destFolder = Directory.Exists(Global.SongMgrDestFolder);
            bool needsDestFolder = Global.AddMode != SongAddMode.InPlace && Global.AddMode != SongAddMode.Monitor;

            if (!databaseFiles || Global.CrazyktvDatabaseError || Global.SongMgrDatabaseError || Global.CrazyktvDatabaseIsOld || (needsDestFolder && !destFolder))
            { Global.SongMgrDBVerErrorUIStatus = false; } else { Global.CrazyktvDatabaseStatus = true; }
            return (databaseFiles, destFolder);
        }

        private void SongDBUpdate_CheckDatabaseFile()
        {
            bool databaseFiles = GetCrazyktvDatabaseStatus().DatabaseFiles;

            if (Global.CrazyktvDatabaseStatus)
            {
                var DatabaseReadyTask = Task.Factory.StartNew(() => SongDBUpdate_DatabaseReady());
            }
            else if (databaseFiles && !Global.CrazyktvDatabaseError && !Global.SongMgrDatabaseError && Global.CrazyktvDatabaseIsOld)
            {
                MainTabControl.SelectedIndex = MainTabControl.TabPages.IndexOf(SongMaintenance_TabPage);
                SongMaintenance_TabControl.SelectedIndex = SongMaintenance_TabControl.TabPages.IndexOf(SongMaintenance_DBVer_TabPage);
                SongMaintenance_DBVer1Value_Label.Text = "正在更新...";
                SongMaintenance_DBVerTooltip_Label.Text = "偵測到資料庫結構更動,開始進行更新...";
                var UpdateDBTask = Task.Factory.StartNew(() => SongDBUpdate_UpdateDatabaseFile());
            }
            else
            {
                Global.DatabaseUpdateFinished = true;
            }
        }

        /// <summary>The 歌庫 needs no 資料庫結構更新: show that and the 錢櫃 date, then finish loading.</summary>
        private void SongDBUpdate_DatabaseReady()
        {
            SongDBUpdate_ShowCashboxUpdDate();
            this.BeginInvoke((Action)delegate ()
            {
                SongMaintenance_DBVer1Value_Label.Text = "最新";
            });
            SongDBUpdate_UpdateFinish();
        }

        private void SongDBUpdate_ShowCashboxUpdDate()
        {
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, "select * from ktv_Version", ""))
            {
                if (dt.Rows.Count == 0) return;
                Global.CashboxUpdDate = DateTime.Parse(dt.Rows[dt.Rows.Count - 1]["CashboxUpdDate"].ToString());
            }
            this.BeginInvoke((Action)delegate ()
            {
                Cashbox_UpdDateValue_Label.Text = (CultureInfo.CurrentCulture.Name == "zh-TW") ? Global.CashboxUpdDate.ToLongDateString() : Global.CashboxUpdDate.ToShortDateString();
            });
        }

        /// <summary>資料庫結構更新 of the loaded 歌庫, on a background task; a SQLite 歌庫 can only be told how to get there.</summary>
        private void SongDBUpdate_UpdateDatabaseFile()
        {
            Global.TimerStartTime = DateTime.Now;

            // 結構更新用的是 Access 專用的 alter column,SQLite 歌庫無法更新
            if (Db.IsSQLite(Global.CrazyktvDatabaseFile))
            {
                this.BeginInvoke((Action)delegate ()
                {
                    SongMaintenance_DBVer1Value_Label.Text = "需要更新";
                    SongMaintenance_DBVerTooltip_Label.Text = "SQLite 歌庫無法更新資料庫結構,請載入原本的 .mdb 歌庫更新後,再重新匯出 SQLite 資料庫。";
                });
                Global.DatabaseUpdateFinished = true;
                return;
            }

            try
            {
                SchemaUpgrade.Run(Global.CrazyktvDatabaseFile, Application.StartupPath + @"\SongMgr\Backup");
            }
            catch (Exception ex)
            {
                // A SchemaUpgradeException has put the backup back; anything else failed before the file was touched.
                string message = (ex is SchemaUpgradeException) ? ex.Message : "資料庫結構更新失敗: " + ex.Message;
                this.BeginInvoke((Action)delegate ()
                {
                    SongMaintenance_DBVer1Value_Label.Text = "需要更新";
                    SongMaintenance_DBVerTooltip_Label.Text = message;
                });
                Global.DatabaseUpdateFinished = true;
                return;
            }

            SongDBUpdate_ShowCashboxUpdDate();
            this.BeginInvoke((Action)delegate ()
            {
                Global.TimerEndTime = DateTime.Now;
                SongMaintenance_DBVer1Value_Label.Text = "最新";
                SongMaintenance_DBVerTooltip_Label.Text = "";
                SongMaintenance_Tooltip_Label.Text = "已完成資料庫結構更新,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
            });
            SongDBUpdate_UpdateFinish();
        }


        private void SongDBUpdate_UpdateFinish()
        {
            bool SongMgrDestFolder = GetCrazyktvDatabaseStatus().DestFolder;

            if (Global.CrazyktvDatabaseStatus)
            {
                DataTable dt = new DataTable();
                string SongQuerySqlStr = "select Song_Id from ktv_Song";
                dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");
                if (dt.Rows.Count > 0)
                {
                    var d5code = from row in dt.AsEnumerable()
                                 where row.Field<string>("Song_Id").Length == 5
                                 select row;

                    var d6code = from row in dt.AsEnumerable()
                                 where row.Field<string>("Song_Id").Length == 6
                                 select row;

                    int MaxDigitCode;
                    if (d5code.Count<DataRow>() > d6code.Count<DataRow>()) { MaxDigitCode = 5; } else { MaxDigitCode = 6; }

                    switch (MaxDigitCode)
                    {
                        case 5:
                            ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb => cb.Enabled = false);
                            if (Global.SongMgrMaxDigitCode != "1")
                            {
                                ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb =>
                                {
                                    cb.SelectedValue = 1;
                                    SongMgrCfg_SaveMaxDigitCode();
                                });
                            }
                            break;
                        case 6:
                            ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb => cb.Enabled = false);
                            if (Global.SongMgrMaxDigitCode != "2")
                            {
                                ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb =>
                                {
                                    cb.SelectedValue = 2;
                                    SongMgrCfg_SaveMaxDigitCode();
                                });
                            }
                            break;
                    }

                    var query = from row in dt.AsEnumerable()
                                where row.Field<string>("Song_Id").Length != MaxDigitCode
                                select row;

                    if (query.Count<DataRow>() > 0)
                    {
                        Global.SongMgrDBVerErrorUIStatus = false;
                        ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = false);
                        ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeCorrect_Button, btn => btn.Enabled = true);
                        Global.CrazyktvDatabaseMaxDigitCode = false;
                        Global.CrazyktvDatabaseStatus = false;
                    }
                    else
                    {
                        if (Global.AddMode == SongAddMode.InPlace || Global.AddMode == SongAddMode.Monitor)
                        {
                            Global.SongMgrDBVerErrorUIStatus = true;

                        }
                        else
                        {
                            if (SongMgrDestFolder) { Global.SongMgrDBVerErrorUIStatus = true; } else { Global.SongMgrDBVerErrorUIStatus = false; }
                        }

                        switch (Global.SongMgrMaxDigitCode)
                        {
                            case "1":
                                ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = true);
                                break;
                            case "2":
                                ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = false);
                                break;
                        }
                        ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeCorrect_Button, btn => btn.Enabled = false);
                        Global.CrazyktvDatabaseMaxDigitCode = true;
                    }
                }
                else // 空白資料庫
                {
                    if (Global.AddMode == SongAddMode.InPlace || Global.AddMode == SongAddMode.Monitor)
                    {
                        Global.SongMgrDBVerErrorUIStatus = true;
                    }
                    else
                    {
                        if (SongMgrDestFolder) { Global.SongMgrDBVerErrorUIStatus = true; } else { Global.SongMgrDBVerErrorUIStatus = false; }
                    }

                    ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = false);
                    ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeCorrect_Button, btn => btn.Enabled = false);
                    Global.CrazyktvDatabaseMaxDigitCode = true;
                }
                dt.Dispose();
                dt = null;
            }

            if (Global.CrazyktvDatabaseStatus)
            {
                // 檢查是否有自訂語系
                Common_CheckSongLang();

                // 統計歌曲數量
                Task.Factory.StartNew(() => Common_GetSongStatisticsTask());

                // 統計歌手數量
                Task.Factory.StartNew(() => Common_GetSingerStatisticsTask());

                // 檢查備份移除歌曲是否要刪除
                Task.Factory.StartNew(() => Common_CheckBackupRemoveSongTask());

                this.BeginInvoke((Action)delegate()
                {
                    // 載入我的最愛清單
                    SongQuery_GetFavoriteUserList();

                    // 歌庫設定 - 載入下拉選單清單及設定
                    SongMgrCfg_SetLangLB();

                    // 歌庫維護 - 載入下拉選單清單及設定
                    SongMaintenance_GetFavoriteUserList();
                    SongMaintenance_SetCustomLangControl();
                });
            }
            Global.DatabaseUpdateFinished = true;
        }

    }
}

using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    public partial class MainForm : Form
    {
        private void SongMaintenance_Save_Button_Click(object sender, EventArgs e)
        {
            switch (SongMaintenance_Save_Button.Text)
            {
                case "儲存設定":
                    SongMgrSettings.Save("SongMaintenanceReplayGainVolume");
                    break;
                case "更新語系":
                    Global.CrazyktvSongLangList = new List<string>();
                    Global.CrazyktvSongLangKeyWordList = new List<string>();
                    TextBox[] SongMaintenance_Lang_TextBox =
                    {
                        SongMaintenance_Lang1_TextBox,
                        SongMaintenance_Lang2_TextBox,
                        SongMaintenance_Lang3_TextBox,
                        SongMaintenance_Lang4_TextBox,
                        SongMaintenance_Lang5_TextBox,
                        SongMaintenance_Lang6_TextBox,
                        SongMaintenance_Lang7_TextBox,
                        SongMaintenance_Lang8_TextBox,
                        SongMaintenance_Lang9_TextBox,
                        SongMaintenance_Lang10_TextBox
                    };

                    for (int i = 0; i < SongMaintenance_Lang_TextBox.Count<TextBox>(); i++)
                    {
                        Global.CrazyktvSongLangList.Add(SongMaintenance_Lang_TextBox[i].Text);
                    }
                    SongMgrSettings.Save("CrazyktvSongLangStr");

                    TextBox[] SongMaintenance_LangIDStr_TextBox =
                    {
                        SongMaintenance_Lang1IDStr_TextBox,
                        SongMaintenance_Lang2IDStr_TextBox,
                        SongMaintenance_Lang3IDStr_TextBox,
                        SongMaintenance_Lang4IDStr_TextBox,
                        SongMaintenance_Lang5IDStr_TextBox,
                        SongMaintenance_Lang6IDStr_TextBox,
                        SongMaintenance_Lang7IDStr_TextBox,
                        SongMaintenance_Lang8IDStr_TextBox,
                        SongMaintenance_Lang9IDStr_TextBox,
                        SongMaintenance_Lang10IDStr_TextBox
                    };

                    for (int i = 0; i < SongMaintenance_LangIDStr_TextBox.Count<TextBox>(); i++)
                    {
                        Global.CrazyktvSongLangKeyWordList.Add(SongMaintenance_LangIDStr_TextBox[i].Text);
                    }
                    SongMgrSettings.Save("CrazyktvSongLangKeyWord");

                    Global.TimerStartTime = DateTime.Now;
                    Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                    Common_SwitchSetUI(false);

                    var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_SongLangUpdateTask())
                    };

                    Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                    {
                        Global.TimerEndTime = DateTime.Now;
                        this.BeginInvoke((Action)delegate()
                        {
                            SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 筆自訂語系資料,失敗 " + Global.TotalList[1] + " 筆,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                            Common_RefreshSongLang();
                            Task.Factory.StartNew(() => Common_GetSongStatisticsTask());
                            Common_SwitchSetUI(true);
                        });
                    });
                    break;
            }
        }


        #region --- 拼音校正 ---

        private void SongMaintenance_SingerSpellCorrect_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要校正歌手拼音嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在解析歌手的拼音資料,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_SpellCorrectTask("ktv_Singer")),
                    Task.Factory.StartNew(() => SongMaintenance_SpellCorrectTask("ktv_AllSinger"))
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), SpellCorrectEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 位歌庫歌手及 " + Global.TotalList[1] + " 位預設歌手的拼音資料,失敗 " + Global.TotalList[2] + " 位,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }

        private void SongMaintenance_SongSpellCorrect_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要校正歌曲拼音嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在解析歌曲的拼音資料,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_SpellCorrectTask("ktv_Song"))
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), SpellCorrectEndTask =>
                {
                    this.BeginInvoke((Action)delegate()
                    {
                        Global.TimerEndTime = DateTime.Now;
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 首歌曲的拼音資料,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                        SongMaintenance.DisposeSongDataTable();
                    });
                });
            }
        }

        /// <summary>
        /// 拼音校正 for one 歌手資料表: rewrites each singer's 拼音 / 筆劃 / 筆形 from the name. Counts the singers
        /// tried in Global.TotalList[<paramref name="countIndex"/>] and the failures in Global.TotalList[2].
        /// </summary>
        private void SongMaintenance_SpellCorrectSingers(SingerTable table, DataTable singers, int countIndex)
        {
            if (singers.Rows.Count <= 0) return;

            var updates = new List<(string Id, string Name, NameSpell Spell)>();
            foreach (DataRow row in singers.AsEnumerable())
            {
                string name = row["Singer_Name"].ToString();
                NameSpell spell = NameSpell.Of(name);
                if (spell.Strokes == "") spell = new NameSpell(spell.Spell, spell.SpellNum, "0", spell.PenStyle);
                updates.Add((row["Singer_Id"].ToString(), name, spell));
            }

            using (SongLibraryStore store = SongLibraryStore.Open(table))
            {
                foreach (var u in updates)
                {
                    try
                    {
                        lock (LockThis)
                        {
                            Global.TotalList[countIndex]++;
                        }
                        store.UpdateSingerSpell(table, u.Id, u.Spell);
                    }
                    catch
                    {
                        lock (LockThis)
                        {
                            Global.TotalList[2]++;
                        }
                        Global.SongLog.Add("【拼音校正】更新資料庫時發生錯誤: " + u.Id + "|" + u.Name + "|" + u.Spell.Spell + "|" + u.Spell.Strokes + "|" + u.Spell.SpellNum + "|" + u.Spell.PenStyle);
                    }

                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "正在更新第 " + Global.TotalList[0] + " 位歌庫歌手及 " + Global.TotalList[1] + " 位預設歌手的拼音資料,請稍待...";
                    });
                }
            }
        }

        private void SongMaintenance_SpellCorrectTask(object TableName)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();
            List<string> SpellList = new List<string>();

            switch ((string)TableName)
            {
                case "ktv_Singer":
                    SongMaintenance_SpellCorrectSingers(SingerTable.Library, Global.SingerDT, 0);
                    break;
                case "ktv_AllSinger":
                    SongMaintenance_SpellCorrectSingers(SingerTable.AllSingers, Global.AllSingerDT, 1);
                    break;
                case "ktv_Song":
                    if (Global.SongDT.Rows.Count > 0)
                    {
                        foreach (DataRow row in Global.SongDT.AsEnumerable())
                        {
                            string str = "";
                            SpellList = new List<string>();

                            str = row["Song_Id"].ToString() + "|";
                            str += row["Song_SongName"].ToString() + "|";
                            SpellList = CommonFunc.GetSongNameSpell(row["Song_SongName"].ToString());
                            if (SpellList[2] == "") SpellList[2] = "0";
                            str += SpellList[0] + "|" + SpellList[1] + "|" + SpellList[2] + "|" + SpellList[3];
                            list.Add(str);
                            this.BeginInvoke((Action)delegate()
                            {
                                SongMaintenance_Tooltip_Label.Text = "正在解析第 " + list.Count + " 首歌曲的拼音資料,請稍待...";
                            });
                        }

                        SongLibraryStore store = SongLibraryStore.Open();
                        List<string> valuelist = new List<string>();

                        foreach (string str in list)
                        {
                            valuelist = new List<string>(str.Split('|'));

                            try
                            {
                                store.UpdateSong(new SongRecord { Id = valuelist[0], Spell = valuelist[2], SpellNum = valuelist[3], SongStroke = valuelist[4], PenStyle = valuelist[5] }, SongColumns.Spell | SongColumns.SpellNum | SongColumns.SongStroke | SongColumns.PenStyle);
                                lock (LockThis)
                                {
                                    Global.TotalList[0]++;
                                }
                            }
                            catch
                            {
                                lock (LockThis)
                                {
                                    Global.TotalList[1]++;
                                }
                                Global.SongLog.Add("【拼音校正】更新資料庫時發生錯誤: " + str);
                            }

                            this.BeginInvoke((Action)delegate()
                            {
                                SongMaintenance_Tooltip_Label.Text = "正在更新第 " + Global.TotalList[0] + " 首歌曲的拼音資料,請稍待...";
                            });
                        }
                        store.Dispose();
                        list.Clear();
                    }
                    break;
            }
        }


        #endregion


        #region --- 雙音軌聲道對映互換 ---


        private void SongMaintenance_LRTrackExchange_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要互換雙音軌聲道對映嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_LRTrackExchangeTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), LRTrackExchangeEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        if (Global.TotalList[3] == 0)
                        {
                            SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 首雙音軌歌曲的聲道資料,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        }
                        else
                        {
                            SongMaintenance_Tooltip_Label.Text = "因部份歌曲檔案不存在,已取消互換雙音軌聲道對映作業,請參操作記錄。";
                        }
                        Task.Factory.StartNew(() => Common_GetSongStatisticsTask());
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }


        private void SongMaintenance_LRTrackExchangeTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();

            var query = from row in Global.SongDT.AsEnumerable()
                         where Convert.ToInt32(row["Song_Track"]) == 1 ||
                               Convert.ToInt32(row["Song_Track"]) == 2
                         select row;

            if (query.Count<DataRow>() > 0)
            {
                foreach (DataRow row in query)
                {
                    string SongFilePath = Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"));
                    if (File.Exists(SongFilePath))
                    {
                        int AudioCount = MediaProbe.AudioStreamCount(SongFilePath);

                        if (AudioCount == 2)
                        {
                            switch (row["Song_Track"].ToString())
                            {
                                case "1":
                                    list.Add(string.Format("{0}|{1}|{2}|{3}", "2", row.Field<string>("Song_Id"), row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")));
                                    break;
                                case "2":
                                    list.Add(string.Format("{0}|{1}|{2}|{3}", "1", row.Field<string>("Song_Id"), row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")));
                                    break;
                            }
                        }
                    }
                    else
                    {
                        Global.TotalList[3]++;
                        Global.SongLog.Add("【互換雙音軌聲道對映】歌曲檔案不存在: " + SongFilePath);
                    }
                    
                    Global.TotalList[2]++;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "正在解析第 " + Global.TotalList[2] + " 首歌曲的聲道資料,請稍待...";
                    });
                }
            }
            
            if (list.Count > 0)
            {
                using (SongLibraryStore store = SongLibraryStore.Open())
                {
                    List<string> valuelist = new List<string>();
                    foreach (string str in list)
                    {
                        lock (LockThis) { Global.TotalList[0]++; }
                        this.BeginInvoke((Action)delegate ()
                        {
                            SongMaintenance_Tooltip_Label.Text = "正在轉換第 " + Global.TotalList[0] + " 首雙音軌歌曲的聲道資料,失敗 " + Global.TotalList[1] + " 首,請稍待...";
                        });

                        valuelist = new List<string>(str.Split('|'));

                        string SongFilePath = Path.Combine(valuelist[2], valuelist[3]);
                        string SongFileName = Path.GetFileName(SongFilePath);
                        string DestFileName = SongFileName;

                        Regex vrRegex = new Regex(@"_vr[\.|_]", RegexOptions.IgnoreCase);
                        Regex vlRegex = new Regex(@"_vl[\.|_]", RegexOptions.IgnoreCase);

                        if (vrRegex.IsMatch(SongFileName.ToLower()))
                        {
                            DestFileName = Regex.Replace(DestFileName, @"_vr[\.|_]", delegate (Match match)
                            {
                                return Regex.Replace(match.Value, "_vr", "_VL", RegexOptions.IgnoreCase);
                            }, RegexOptions.IgnoreCase);
                        }
                        else if (vlRegex.IsMatch(SongFileName.ToLower()))
                        {
                            DestFileName = Regex.Replace(DestFileName, @"_vl[\.|_]", delegate (Match match)
                            {
                                return Regex.Replace(match.Value, "_vl", "_VR", RegexOptions.IgnoreCase);
                            }, RegexOptions.IgnoreCase);
                        }
                        
                        bool MoveError = false;
                        if ( DestFileName != SongFileName)
                        {
                            string SongDir = Path.GetDirectoryName(SongFilePath);
                            string DestFilePath = Path.Combine(valuelist[2], DestFileName);

                            if (File.Exists(DestFilePath))
                            {
                                MoveError = true;
                                Global.SongLog.Add("【互換雙音軌聲道對映】目的檔案已存在: " + DestFilePath );
                                lock (LockThis) { Global.TotalList[1]++; }
                            }
                            else
                            {
                                try
                                {
                                    File.Move(SongFilePath, DestFilePath);
                                    valuelist[3] = DestFileName;
                                }
                                catch
                                {
                                    MoveError = true;
                                    Global.SongLog.Add("【互換雙音軌聲道對映】異動檔案時發生錯誤: " + SongFilePath + " (檔案唯讀或正在使用)");
                                    lock (LockThis) { Global.TotalList[1]++; }
                                }
                            }
                        }

                        if (!MoveError)
                        {
                            try
                            {
                                store.UpdateSong(new SongRecord { Id = valuelist[1], Track = valuelist[0], FileName = valuelist[3] }, SongColumns.Track | SongColumns.FileName);
                            }
                            catch
                            {
                                lock (LockThis) { Global.TotalList[1]++; }
                                Global.SongLog.Add("【互換雙音軌聲道對映】更新資料庫時發生錯誤: " + str);
                            }

                        }
                    }
                }
            }
            list.Clear();
        }


        #endregion


        #region --- 音量變更 ---

        private void SongMaintenance_VolumeChange_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要變更音量嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_VolumeChangeTask(SongMaintenance_VolumeChange_TextBox.Text))
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), VolumeChangeEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 首歌曲的音量資料,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }


        private void SongMaintenance_VolumeChangeTask(object VolumeValue)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();
            
            foreach (DataRow row in Global.SongDT.AsEnumerable())
            {
                list.Add((string)VolumeValue + "|" + row["Song_Id"].ToString());
                
                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在解析第 " + list.Count + " 首歌曲的音量資料,請稍待...";
                });
            }

            SongLibraryStore store = SongLibraryStore.Open();
            List<string> valuelist = new List<string>();

            foreach (string str in list)
            {
                valuelist = new List<string>(str.Split('|'));

                try
                {
                    store.UpdateSong(new SongRecord { Id = valuelist[1], Volume = valuelist[0] }, SongColumns.Volume);
                    lock (LockThis)
                    {
                        Global.TotalList[0]++;
                    }
                }
                catch
                {
                    lock (LockThis)
                    {
                        Global.TotalList[1]++;
                    }
                    Global.SongLog.Add("【變更歌曲音量】更新資料庫時發生錯誤: " + str);
                }

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在轉換第 " + Global.TotalList[0] + " 首歌曲的音量資料,請稍待...";
                });
            }
            store.Dispose();
            list.Clear();
        }


        #endregion


        #region --- 編碼位數轉換 ---


        private void SongMaintenance_CodeConvTo6_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要轉換為 6 位數編碼嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                if (Global.SongMgrMaxDigitCode != "2") SongMgrCfg_MaxDigitCode_ComboBox.SelectedValue = 2;
                SongMgrCfg_GetSongMgrLangCode();

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_CodeConvTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), CodeConvEndTask =>
                {
                    this.BeginInvoke((Action)delegate()
                    {
                        // 我的最愛 were renumbered with their songs; a 我的最愛 list on screen still shows the old numbers
                        if (Global.SongQueryQueryType == "FavoriteQuery")
                        {
                            Global.SongQueryQueryType = "SongQuery";
                            SongQuery_DataGridView.DataSource = null;
                            if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                            SongQuery_QueryStatus_Label.Text = "";
                        }

                        Global.TimerEndTime = DateTime.Now;
                        SongMaintenance_Tooltip_Label.Text = "總共轉換 " + Global.TotalList[2] + " 首歌曲的歌曲編號,失敗 " + Global.TotalList[3] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        SongMgrCfg_SaveMaxDigitCode();
                        SongDBUpdate_CheckDatabaseFile();
                        Common_SwitchSetUI(true);
                        SongMaintenance.DisposeSongDataTable();
                    });
                });
            }
        }

        private void SongMaintenance_CodeConvTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            List<string> list = new List<string>();
            SongIdAllocator songIds = SongIdAllocator.FromSettings(Enumerable.Empty<string>());

            if (Global.SongDT.Rows.Count > 0)
            {
                foreach (DataRow row in Global.SongDT.AsEnumerable())
                {
                    string str = "";
                    if (songIds.HasLang(row["Song_Lang"].ToString()))
                    {
                        string NewSongId = songIds.Take(row["Song_Lang"].ToString());

                        str = row["Song_Id"].ToString() + "|";
                        str += NewSongId + "|";
                        str += row["Song_SongName"].ToString() + "|";
                        str += row["Song_Lang"].ToString();
                        list.Add(str);
                    }
                    else
                    {
                        lock (LockThis)
                        {
                            Global.TotalList[3]++;
                        }
                        str = row["Song_Id"].ToString() + "|";
                        str += row["Song_SongName"].ToString() + "|";
                        str += row["Song_Lang"].ToString();

                        Global.SongLog.Add("【編碼位數轉換】此首歌曲的語系資料不正確: " + str);
                    }

                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "正在配發第 " + Global.SongDT.Rows.IndexOf(row) + " 首歌曲的歌曲編號,請稍待...";
                    });
                }
            }

            SongLibraryStore store = SongLibraryStore.Open();
            List<string> valuelist = new List<string>();

            foreach (string str in list)
            {
                valuelist = new List<string>(str.Split('|'));

                try
                {
                    store.RenumberSong(valuelist[0], valuelist[1], valuelist[2], valuelist[3]);
                    lock (LockThis)
                    {
                        Global.TotalList[2]++;
                    }
                }
                catch
                {
                    lock (LockThis)
                    {
                        Global.TotalList[3]++;
                    }
                    Global.SongLog.Add("【編碼位數轉換】更新資料庫時發生錯誤: " + str);
                }

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在轉換第 " + Global.TotalList[2] + " 首歌曲的歌曲編號,請稍待...";
                });
            }
            store.Dispose();
            list.Clear();
        }


        private void SongMaintenance_CodeCorrect_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要校正編碼位數嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                var d5code = from row in Global.SongDT.AsEnumerable()
                             where row.Field<string>("Song_Id").Length == 5
                             select row;

                var d6code = from row in Global.SongDT.AsEnumerable()
                             where row.Field<string>("Song_Id").Length == 6
                             select row;

                if (d5code.Count<DataRow>() > d6code.Count<DataRow>())
                {
                    if (Global.SongMgrMaxDigitCode != "1") SongMgrCfg_MaxDigitCode_ComboBox.SelectedValue = 1;
                }
                else
                {
                    if (Global.SongMgrMaxDigitCode != "2") SongMgrCfg_MaxDigitCode_ComboBox.SelectedValue = 2;
                }
                SongIdAllocator songIds = SongIdAllocator.FromLibrary();
                SongMgrCfg_GetSongMgrLangCode();

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_CodeCorrectTask(songIds))
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), CodeCorrectEndTask =>
                {
                    this.BeginInvoke((Action)delegate()
                    {
                        Global.TimerEndTime = DateTime.Now;
                        SongMaintenance_Tooltip_Label.Text = "總共校正 " + Global.TotalList[0] + " 首歌曲的歌曲編號,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        SongMgrCfg_SaveMaxDigitCode();
                        SongDBUpdate_CheckDatabaseFile();
                        Common_SwitchSetUI(true);
                        SongMaintenance.DisposeSongDataTable();
                    });
                });
            }
        }


        private void SongMaintenance_CodeCorrectTask(SongIdAllocator songIds)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            int CorrectDigitCode = 0;

            if (Global.SongMgrMaxDigitCode == "1") { CorrectDigitCode = 5; } else { CorrectDigitCode = 6; }

            List<string> list = new List<string>();

            if (Global.SongDT.Rows.Count > 0)
            {
                var query = from row in Global.SongDT.AsEnumerable()
                            where row.Field<string>("Song_Id").Length != CorrectDigitCode
                            select row;

                foreach (DataRow row in query)
                {
                    string str = "";
                    if (songIds.HasLang(row["Song_Lang"].ToString()))
                    {
                        string LangStr = row["Song_Lang"].ToString();
                        string NewSongId = songIds.Take(LangStr);

                        str = row["Song_Id"].ToString() + "|";
                        str += NewSongId + "|";
                        str += row["Song_SongName"].ToString() + "|";
                        str += row["Song_Lang"].ToString();
                        list.Add(str);
                    }
                    else
                    {
                        lock (LockThis)
                        {
                            Global.TotalList[1]++;
                        }
                        str = row["Song_Id"].ToString() + "|";
                        str += row["Song_SongName"].ToString() + "|";
                        str += row["Song_Lang"].ToString();

                        Global.SongLog.Add("【編碼位數轉換】此首歌曲的語系資料不正確: " + str);
                    }

                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "正在配發第 " + Global.SongDT.Rows.IndexOf(row) + " 首歌曲的歌曲編號,請稍待...";
                    });
                }
            }

            SongLibraryStore store = SongLibraryStore.Open();
            List<string> valuelist = new List<string>();

            foreach (string str in list)
            {
                valuelist = new List<string>(str.Split('|'));

                try
                {
                    store.RenumberSong(valuelist[0], valuelist[1], valuelist[2], valuelist[3]);
                    lock (LockThis)
                    {
                        Global.TotalList[0]++;
                    }
                }
                catch
                {
                    lock (LockThis)
                    {
                        Global.TotalList[1]++;
                    }
                    Global.SongLog.Add("【編碼位數轉換】更新資料庫時發生錯誤: " + str);
                }

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在轉換第 " + Global.TotalList[0] + " 首歌曲的歌曲編號,請稍待...";
                });
            }
            store.Dispose();
            list.Clear();
        }


        #endregion


        #region --- 歌曲路徑變更 ---


        private void SongMaintenance_SongPathChange_Button_Click(object sender, EventArgs e)
        {
            if (SongMaintenance_SongPathChange_Button.Text == "瀏覽")
            {
                FolderBrowserDialog opd = new FolderBrowserDialog();
                if (SongMaintenance_DestSongPath_TextBox.Text != "") opd.SelectedPath = SongMaintenance_DestSongPath_TextBox.Text;

                if (opd.ShowDialog() == DialogResult.OK && opd.SelectedPath.Length > 0)
                {
                    SongMaintenance_DestSongPath_TextBox.Text = SongPathChange.NormalizeDirectory(opd.SelectedPath);
                    SongMaintenance_SongPathChange_Button.Text = "變更";
                }
            }
            else
            {
                string rawSrcPath = SongMaintenance_SrcSongPath_TextBox.Text.Trim();
                string rawDestPath = SongMaintenance_DestSongPath_TextBox.Text.Trim();

                if (string.IsNullOrEmpty(rawSrcPath))
                {
                    SongMaintenance_Tooltip_Label.Text = "你尚未輸入【原始路徑】!";
                }
                else if (string.IsNullOrEmpty(rawDestPath))
                {
                    SongMaintenance_Tooltip_Label.Text = "你尚未選擇【目標路徑】!";
                }
                else
                {
                    string SrcSongPath = SongPathChange.NormalizeDirectory(rawSrcPath);
                    string DestSongPath = SongPathChange.NormalizeDirectory(rawDestPath);

                    SongMaintenance_SrcSongPath_TextBox.Text = SrcSongPath;
                    SongMaintenance_DestSongPath_TextBox.Text = DestSongPath;

                    if (MessageBox.Show("你確定要變更歌曲路徑嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Jobs.TryStart(new BackgroundJob
                        {
                            Name = "變更歌曲路徑",
                            Status = text => SongMaintenance_Tooltip_Label.Text = text,
                            Work = ctx =>
                            {
                                ThreadPriority priority = Thread.CurrentThread.Priority;
                                Thread.CurrentThread.Priority = ThreadPriority.Lowest;
                                try
                                {
                                    SongMaintenance.CreateSongDataTable();
                                    var changes = SongPathChange.Plan(Global.SongDT, SrcSongPath, DestSongPath, ctx);
                                    using (SongLibraryStore store = SongLibraryStore.Open())
                                    {
                                        SongPathChange.Apply(store, changes, ctx);
                                    }
                                }
                                finally
                                {
                                    SongMaintenance.DisposeSongDataTable();
                                    Thread.CurrentThread.Priority = priority;
                                }
                            },
                            Summary = (c, secs) => "總共更新 " + c[0] + " 首歌曲的歌曲路徑資料,失敗 " + c[1] + " 首,共花費 " + secs + " 秒完成。",
                            Completed = () =>
                            {
                                SongMaintenance_SrcSongPath_TextBox.Text = "";
                                SongMaintenance_DestSongPath_TextBox.Text = "";
                                SongMaintenance_SongPathChange_Button.Text = "瀏覽";
                                Task.Factory.StartNew(() => Common_GetSongStatisticsTask());
                            },
                        });
                    }
                }
            }
        }


        #endregion


        #region --- 播放次數 ---


        private void SongMaintenance_PlayCountExport_Button_Click(object sender, EventArgs e)
        {
            SongMaintenance.CreateSongDataTable();
            List<string> list = new List<string>();

            if (Global.SongDT.Rows.Count > 0)
            {
                foreach (DataRow row in Global.SongDT.AsEnumerable())
                {
                    if (row["Song_PlayCount"].ToString() != "")
                    {
                        if (Convert.ToInt32(row["Song_PlayCount"].ToString()) > 0)
                        {
                            list.Add("ktv_Song|" + row["Song_Lang"].ToString() + "|" + row["Song_Singer"].ToString() + "|" + row["Song_SongName"].ToString() + "|" + row["Song_PlayCount"].ToString());
                        }
                    }
                }
            }

            if (!Directory.Exists(Application.StartupPath + @"\SongMgr\Backup")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Backup");
            StreamWriter sw = new StreamWriter(Application.StartupPath + @"\SongMgr\Backup\PlayCount.txt");
            foreach (string str in list)
            {
                sw.WriteLine(str);
            }

            SongMaintenance_Tooltip_Label.Text = @"已將播放次數資料匯出至【SongMgr\Backup\PlayCount.txt】檔案。";
            sw.Close();
            SongMaintenance.DisposeSongDataTable();
        }


        private void SongMaintenance_PlayCountImport_Button_Click(object sender, EventArgs e)
        {
            if (File.Exists(Application.StartupPath + @"\SongMgr\Backup\PlayCount.txt"))
            {
                if (SongMaintenance_Tooltip_Label.Text == @"【SongMgr\Backup\PlayCount.txt】播放次數備份檔案不存在!") SongMaintenance_Tooltip_Label.Text = "";
                if (MessageBox.Show("你確定要匯入播放次數資料嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Global.TimerStartTime = DateTime.Now;
                    Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                    SongMaintenance.CreateSongDataTable();
                    Common_SwitchSetUI(false);

                    SongMaintenance_Tooltip_Label.Text = "正在匯入播放次數資料,請稍待...";

                    var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_PlayCountImportTask())
                    };

                    Task.Factory.ContinueWhenAll(tasks.ToArray(), PlayCountImportEndTask =>
                    {
                        Global.TimerEndTime = DateTime.Now;
                        this.BeginInvoke((Action)delegate()
                        {
                            SongMaintenance_Tooltip_Label.Text = "總共匯入 " + Global.TotalList[0] + " 筆播放次數資料,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                            Common_SwitchSetUI(true);
                            SongMaintenance.DisposeSongDataTable();
                        });
                    });
                }
            }
            else
            {
                SongMaintenance_Tooltip_Label.Text = @"【SongMgr\Backup\PlayCount.txt】播放次數備份檔案不存在!";
            }
        }


        private void SongMaintenance_PlayCountImportTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            List<string> list = new List<string>();
            List<string> Addlist = new List<string>();

            SongLibraryStore store = SongLibraryStore.Open();

            StreamReader sr = new StreamReader(Application.StartupPath + @"\SongMgr\Backup\PlayCount.txt", Encoding.UTF8);
            while (!sr.EndOfStream)
            {
                Addlist.Add(sr.ReadLine());
            }
            sr.Close();

            foreach (string AddStr in Addlist)
            {
                list = new List<string>(Regex.Split(AddStr, @"\|", RegexOptions.None));
                switch (list[0])
                {
                    case "ktv_Song":
                        var query = from row in Global.SongDT.AsEnumerable()
                                    where row.Field<string>("Song_Lang").Equals(list[1]) &&
                                          row.Field<string>("Song_Singer").Equals(list[2]) &&
                                          row.Field<string>("Song_SongName").Equals(list[3])
                                    select row;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow row in query)
                            {
                                string SongId = row["Song_Id"].ToString();
                                store.UpdateSong(new SongRecord { Id = SongId, PlayCount = list[4] }, SongColumns.PlayCount);
                                lock (LockThis)
                                {
                                    Global.TotalList[0]++;
                                }
                                break;
                            }
                        }
                        break;
                }
                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在匯入第 " + Global.TotalList[0] + " 首歌曲的播放次數資料,請稍待...";
                });
            }
            store.Dispose();
        }


        private void SongMaintenance_PlayCountReset_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要重置播放次數嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_PlayCountResetTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), PlayCountResetEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共重置 " + Global.TotalList[0] + " 首歌曲的播放次數,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }


        private void SongMaintenance_PlayCountResetTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();

            foreach (DataRow row in Global.SongDT.AsEnumerable())
            {
                list.Add("0" + "|" + row["Song_Id"].ToString());

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在解析第 " + list.Count + " 首歌曲的播放次數資料,請稍待...";
                });
            }

            SongLibraryStore store = SongLibraryStore.Open();
            List<string> valuelist = new List<string>();

            foreach (string str in list)
            {
                valuelist = new List<string>(str.Split('|'));

                try
                {
                    store.UpdateSong(new SongRecord { Id = valuelist[1], PlayCount = valuelist[0] }, SongColumns.PlayCount);
                    lock (LockThis)
                    {
                        Global.TotalList[0]++;
                    }
                }
                catch
                {
                    lock (LockThis)
                    {
                        Global.TotalList[1]++;
                    }
                    Global.SongLog.Add("【重置播放次數】更新資料庫時發生錯誤: " + str);
                }

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在重置第 " + Global.TotalList[0] + " 首歌曲的播放次數資料,請稍待...";
                });
            }
            store.Dispose();
            list.Clear();
        }


        #endregion


        #region --- 雜項 ---


        private void SongMaintenance_SongWordCountCorrect_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要校正歌曲字數嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_WordCountCorrectTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 首歌曲的歌曲字數,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }


        private void SongMaintenance_WordCountCorrectTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();
            List<string> SongWordCountList = new List<string>();

            foreach (DataRow row in Global.SongDT.AsEnumerable())
            {
                SongWordCountList = CommonFunc.GetSongWordCount(row["Song_SongName"].ToString());
                string SongWordCount = SongWordCountList[0];

                list.Add(SongWordCount + "|" + row["Song_Id"].ToString());

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在解析第 " + list.Count + " 首歌曲的字數資料,請稍待...";
                });
            }

            SongLibraryStore store = SongLibraryStore.Open();
            List<string> valuelist = new List<string>();

            foreach (string str in list)
            {
                valuelist = new List<string>(str.Split('|'));

                try
                {
                    store.UpdateSong(new SongRecord { Id = valuelist[1], WordCount = valuelist[0] }, SongColumns.WordCount);
                    lock (LockThis)
                    {
                        Global.TotalList[0]++;
                    }
                }
                catch
                {
                    lock (LockThis)
                    {
                        Global.TotalList[1]++;
                    }
                    Global.SongLog.Add("【字數校正】更新資料庫時發生錯誤: " + str);
                }

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在轉換第 " + Global.TotalList[0] + " 首歌曲的字數資料,請稍待...";
                });
            }
            store.Dispose();
            list.Clear();
        }


        private void SongMaintenance_RemoveEmptyDirs_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要移除空資料夾嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在解析空資料夾,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_RemoveEmptyDirsTask(Global.SongMgrDestFolder, false))
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        if (Global.TotalList[0] == 0)
                        {
                            SongMaintenance_Tooltip_Label.Text = "恭喜！在你的歌庫資料夾裡沒有發現任何空白資料夾。";
                        }
                        else
                        {
                            SongMaintenance_Tooltip_Label.Text = "總共移除 " + Global.TotalList[0] + " 個空資料夾,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        }
                        Common_SwitchSetUI(true);
                    });
                });
            }
        }


        private void SongMaintenance_RemoveEmptyDirsTask(string dir, bool stepBack)
        {
            string RootDir = Global.SongMgrDestFolder;
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Directory.GetFileSystemEntries(dir).Length > 0)
            {
                if (!stepBack)
                {
                    foreach (string subdir in Directory.GetDirectories(dir))
                    {
                        SongMaintenance_RemoveEmptyDirsTask(subdir, false);
                    }
                }
            }
            else
            {
                DirectoryInfo dirinfo = new DirectoryInfo(dir);
                if ((dirinfo.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    dirinfo.Attributes &= ~FileAttributes.ReadOnly;
                }

                try
                {
                    Directory.Delete(dir);
                    Global.SongLog.Add("【移除空資料夾】以下為已移除的空資料夾: " + dir);
                    lock (LockThis) { Global.TotalList[0]++; }
                }
                catch
                {
                    Global.SongLog.Add("【移除空資料夾】無法移除因資料夾已被占用: " + dir);
                }

                string prevDir = dir.Substring(0, dir.LastIndexOf("\\"));
                if (RootDir.Length <= prevDir.Length) SongMaintenance_RemoveEmptyDirsTask(prevDir, true);

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "已移除掉 " + Global.TotalList[0] + " 個空資料夾,請稍待...";
                });
            }
        }


        private void SongMaintenance_DetectUnusedFiles_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要偵測歌庫資料夾內的多餘檔案嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Common_SwitchSetUI(false);

                Global.DetectUnusedFilePathList.Clear();
                Global.DetectUnusedFilePathList = null;

                DetectUnusedFileForm newDetectForm = new DetectUnusedFileForm(this);
                this.Hide();
                var DetectFormResult = newDetectForm.ShowDialog();

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_DetectUnusedFilesTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共在歌庫資料夾內偵測到 " + Global.TotalList[0] + " 個的多餘檔案,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                });
            }
        }

        private void SongMaintenance_DetectUnusedFilesTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            if (Global.DetectUnusedFilePathList.Count == 0) return;

            List<string> SongList = new List<string>();
            List<string> FileList = new List<string>();
            List<string> UnusedFileList = new List<string>();

            string SongQuerySqlStr = "select Song_Id, Song_Lang, Song_FileName, Song_Path from ktv_Song order by Song_Id";
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, ""))
            {
                Parallel.ForEach(Global.CrazyktvSongLangList, (langstr, loopState) =>
                {
                    var query = from row in dt.AsEnumerable()
                                where row.Field<string>("Song_Lang").Equals(langstr)
                                select row;

                    if (query.Count<DataRow>() > 0)
                    {
                        foreach (DataRow row in query)
                        {
                            lock (LockThis)
                            {
                                SongList.Add(Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")).ToLower());
                            }
                        }
                    }
                });
            }

            foreach (string dir in Global.DetectUnusedFilePathList)
            {
                FileList.AddRange(Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).ToList<string>().ConvertAll(file => file.ToLower()));
            }

            Parallel.ForEach(FileList, (file, loopState) =>
            {
                if (SongList.IndexOf(file) < 0)
                {
                    lock(LockThis)
                    {
                        Global.TotalList[0]++;
                        UnusedFileList.Add(file);
                    }
                }
            });
            SongList.Clear();
            SongList = null;
            FileList.Clear();
            FileList = null;

            if (UnusedFileList.Count > 0)
            {
                foreach (string file in UnusedFileList)
                {
                    Global.SongLog.Add("【偵測歌庫多餘檔案】以下為偵測到的多餘檔案: " + file);
                }
            }
            UnusedFileList.Clear();
            UnusedFileList = null;
        }

        #endregion


        #region --- 遙控設定 ---


        private void SongMaintenance_RemoteCfgExport_Button_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            string sqlColumnStr = "Remote_Id, Remote_Subject, Remote_Controler, Remote_Controler2, Remote_Name";
            string RemoteQuerySqlStr = "select " + sqlColumnStr + " from ktv_Remote";
            DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, RemoteQuerySqlStr, "");

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.AsEnumerable())
                {
                    list.Add("ktv_Remote," + row["Remote_Id"].ToString() + "," + row["Remote_Subject"].ToString() + "," + row["Remote_Controler"].ToString() + "," + row["Remote_Controler2"].ToString() + "," + row["Remote_Name"].ToString());
                }
            }

            if (!Directory.Exists(Application.StartupPath + @"\SongMgr\Backup")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Backup");
            StreamWriter sw = new StreamWriter(Application.StartupPath + @"\SongMgr\Backup\Remote.txt");
            foreach (string str in list)
            {
                sw.WriteLine(str);
            }

            SongMaintenance_Tooltip_Label.Text = @"已將遙控設定匯出至【SongMgr\Backup\Remote.txt】檔案。";
            sw.Close();
            list.Clear();
            dt.Dispose();
        }


        private void SongMaintenance_RemoteCfgImport_Button_Click(object sender, EventArgs e)
        {
            if (File.Exists(Application.StartupPath + @"\SongMgr\Backup\Remote.txt"))
            {
                if (SongMaintenance_Tooltip_Label.Text == @"【SongMgr\Backup\Remote.txt】遙控設定備份檔案不存在!") SongMaintenance_Tooltip_Label.Text = "";
                if (MessageBox.Show("你確定要重置並匯入遙控設定嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Global.TimerStartTime = DateTime.Now;
                    Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                    Common_SwitchSetUI(false);

                    SongMaintenance_Tooltip_Label.Text = "正在匯入遙控設定,請稍待...";

                    var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_RemoteCfgImportTask())
                    };

                    Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                    {
                        Global.TimerEndTime = DateTime.Now;
                        this.BeginInvoke((Action)delegate()
                        {
                            SongMaintenance_Tooltip_Label.Text = "總共匯入 " + Global.TotalList[0] + " 項遙控設定,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                            Common_SwitchSetUI(true);
                        });
                    });
                }
            }
            else
            {
                SongMaintenance_Tooltip_Label.Text = @"【SongMgr\Backup\Remote.txt】遙控設定備份檔案不存在!";
            }
        }


        private void SongMaintenance_RemoteCfgImportTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();
            List<string> Addlist = new List<string>();

            DbConnection conn = null;
            DbCommand cmd = null;

            conn = Db.OpenConn(Global.CrazyktvDatabaseFile, "");
            string TruncateSqlStr = "delete * from ktv_Remote";
            cmd = Db.Command(TruncateSqlStr, conn);
            cmd.ExecuteNonQuery();

            StreamReader sr = new StreamReader(Application.StartupPath + @"\SongMgr\Backup\Remote.txt", Encoding.UTF8); ;

            while (!sr.EndOfStream)
            {
                Addlist.Add(sr.ReadLine());
            }
            sr.Close();

            string sqlColumnStr = "Remote_Id, Remote_Subject, Remote_Controler, Remote_Controler2, Remote_Name";
            string sqlValuesStr = "@RemoteId, @RemoteSubject, @RemoteControler, @RemoteControler2, @RemoteName";
            string RemoteAddSqlStr = "insert into ktv_Remote ( " + sqlColumnStr + " ) values ( " + sqlValuesStr + " )";
            cmd = Db.Command(RemoteAddSqlStr, conn);

            foreach (string AddStr in Addlist)
            {
                list = new List<string>(Regex.Split(AddStr, ",", RegexOptions.None));
                switch (list[0])
                {
                    case "ktv_Remote":
                        cmd.Parameters.AddWithValue("@RemoteId", list[1]);
                        cmd.Parameters.AddWithValue("@RemoteSubject", list[2]);
                        cmd.Parameters.AddWithValue("@RemoteControler", list[3]);
                        cmd.Parameters.AddWithValue("@RemoteControler2", list[4]);
                        cmd.Parameters.AddWithValue("@RemoteName", list[5]);

                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        lock (LockThis) { Global.TotalList[0]++; }
                        break;
                }
                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在匯入第 " + Global.TotalList[0] + " 項遙控設定資料,請稍待...";
                });
            }
            conn.Close();
        }


        #endregion


        #region --- 拼音資料 ---


        private void SongMaintenance_PhoneticsExport_Button_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            string sqlColumnStr = "Word, Code, Spell, PenStyle, SortIdx, Strokes";
            string PhoneticsQuerySqlStr = "select " + sqlColumnStr + " from ktv_Phonetics order by Code, SortIdx";
            DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, PhoneticsQuerySqlStr, "");

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.AsEnumerable())
                {
                    //list.Add("ktv_Phonetics," + row["Word"].ToString() + "," + CommonFunc.GetWordUnicode(row["Word"].ToString()) + "," + row["Spell"].ToString() + "," + row["PenStyle"].ToString() + "," + row["SortIdx"].ToString() + "," + row["Strokes"].ToString());
                    list.Add("ktv_Phonetics," + row["Word"].ToString() + "," + row["Code"].ToString() + "," + row["Spell"].ToString() + "," + row["PenStyle"].ToString() + "," + row["SortIdx"].ToString() + "," + row["Strokes"].ToString());
                }
            }

            if (!Directory.Exists(Application.StartupPath + @"\SongMgr\Backup")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Backup");
            StreamWriter sw = new StreamWriter(Application.StartupPath + @"\SongMgr\Backup\Phonetics.txt");
            foreach (string str in list)
            {
                sw.WriteLine(str);
            }

            SongMaintenance_Tooltip_Label.Text = @"已將拼音資料匯出至【SongMgr\Backup\Phonetics.txt】檔案。";
            sw.Close();
            list.Clear();
            dt.Dispose();
        }


        private void SongMaintenance_PhoneticsImport_Button_Click(object sender, EventArgs e)
        {
            if (File.Exists(Application.StartupPath + @"\SongMgr\Backup\Phonetics.txt"))
            {
                if (SongMaintenance_Tooltip_Label.Text == @"【SongMgr\Backup\Phonetics.txt】拼音資料備份檔案不存在!") SongMaintenance_Tooltip_Label.Text = "";
                if (MessageBox.Show("你確定要重置並匯入拼音資料嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Global.TimerStartTime = DateTime.Now;
                    Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                    Common_SwitchSetUI(false);

                    SongMaintenance_Tooltip_Label.Text = "正在匯入拼音資料,請稍待...";

                    var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_PhoneticsImportTask(false))
                    };

                    Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                    {
                        Global.TimerEndTime = DateTime.Now;
                        this.BeginInvoke((Action)delegate()
                        {
                            SongMaintenance_Tooltip_Label.Text = "總共匯入 " + Global.TotalList[0] + " 項拼音資料,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                            Common_SwitchSetUI(true);
                        });
                    });
                }
            }
            else
            {
                SongMaintenance_Tooltip_Label.Text = @"【SongMgr\Backup\Phonetics.txt】拼音資料備份檔案不存在!";
            }
        }


        private void SongMaintenance_PhoneticsImportTask(bool Update)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();
            List<string> Addlist = new List<string>();

            DbConnection conn = null;
            DbCommand cmd = null;
            DbCommand Versioncmd = null;

            conn = Db.OpenConn(Global.CrazyktvSongMgrDatabaseFile, "");
            string TruncateSqlStr = "delete * from ktv_Phonetics";
            cmd = Db.Command(TruncateSqlStr, conn);
            cmd.ExecuteNonQuery();

            StreamReader sr;
            if (Update)
            {
                sr = new StreamReader(Application.StartupPath + @"\SongMgr\Update\UpdatePhoneticsDB.txt", Encoding.UTF8);
            }
            else
            {
                sr = new StreamReader(Application.StartupPath + @"\SongMgr\Backup\Phonetics.txt", Encoding.UTF8);
            }

            while (!sr.EndOfStream)
            {
                Addlist.Add(sr.ReadLine());
            }
            sr.Close();

            string sqlColumnStr = "Word, Code, Spell, PenStyle, SortIdx, Strokes";
            string sqlValuesStr = "@Word, @Code, @Spell, @PenStyle, @SortIdx, @Strokes";
            string RemoteAddSqlStr = "insert into ktv_Phonetics ( " + sqlColumnStr + " ) values ( " + sqlValuesStr + " )";
            cmd = Db.Command(RemoteAddSqlStr, conn);

            foreach (string AddStr in Addlist)
            {
                list = new List<string>(Regex.Split(AddStr, ",", RegexOptions.None));
                switch (list[0])
                {
                    case "ktv_Version":
                        string VersionSqlStr = "PhoneticsDB = @PhoneticsDB";
                        string VersionUpdateSqlStr = "update ktv_Version set " + VersionSqlStr + " where Id = @Id";
                        Versioncmd = Db.Command(VersionUpdateSqlStr, conn);
                        Versioncmd.Parameters.AddWithValue("@PhoneticsDB", list[1]);
                        Versioncmd.Parameters.AddWithValue("@Id", "1");
                        Versioncmd.ExecuteNonQuery();
                        Versioncmd.Parameters.Clear();
                        break;
                    case "ktv_Phonetics":
                        cmd.Parameters.AddWithValue("@Word", list[1]);
                        cmd.Parameters.AddWithValue("@Code", list[2]);
                        cmd.Parameters.AddWithValue("@Spell", list[3]);
                        cmd.Parameters.AddWithValue("@PenStyle", list[4]);
                        cmd.Parameters.AddWithValue("@SortIdx", list[5]);
                        cmd.Parameters.AddWithValue("@Strokes", list[6]);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        lock (LockThis) { Global.TotalList[0]++; }
                        break;
                }

                this.BeginInvoke((Action)delegate()
                {

                    SongMaintenance_Tooltip_Label.Text = "正在匯入第 " + Global.TotalList[0] + " 項拼音資料,請稍待...";
                });
            }
            conn.Close();
        }
        
        #endregion

        
        #region --- 歌庫結構重建 ---

        private void SongMaintenance_RebuildSongStructure_Button_Click(object sender, EventArgs e)
        {
            if (SongMaintenance_RebuildSongStructure_Button.Text == "瀏覽")
            {
                FolderBrowserDialog opd = new FolderBrowserDialog();
                if (SongMaintenance_RebuildSongStructure_TextBox.Text != "") opd.SelectedPath = SongMaintenance_RebuildSongStructure_TextBox.Text;

                if (opd.ShowDialog() == DialogResult.OK && opd.SelectedPath.Length > 0)
                {
                    if (Directory.GetFiles(opd.SelectedPath, "*", SearchOption.AllDirectories).Count() == 0)
                    {
                        if (SongMaintenance_Tooltip_Label.Text == "請選擇一個空白的資料夾!") SongMaintenance_Tooltip_Label.Text = "";
                        SongMaintenance_RebuildSongStructure_TextBox.Text = opd.SelectedPath;
                        SongMaintenance_RebuildSongStructure_Button.Text = "重建";
                    }
                    else
                    {
                        SongMaintenance_Tooltip_Label.Text = "請選擇一個空白的資料夾!";
                    }
                }
            }
            else
            {
                if (SongMaintenance_RebuildSongStructure_TextBox.Text == "")
                {
                    SongMaintenance_Tooltip_Label.Text = "你尚未選擇【重建資料夾】!";
                }
                else
                {
                    if (MessageBox.Show("你確定要重建歌庫結構嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Global.TimerStartTime = DateTime.Now;
                        Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                        SongMaintenance.CreateSongDataTable();
                        Common_SwitchSetUI(false);

                        string RebuildSongPath = SongMaintenance_RebuildSongStructure_TextBox.Text;

                        var tasks = new List<Task>()
                        {
                            Task.Factory.StartNew(() => SongMaintenance_RebuildSongStructureTask(RebuildSongPath, Global.SongMgrSongAddMode))
                        };                        ;

                        Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                        {
                            Global.TimerEndTime = DateTime.Now;
                            this.BeginInvoke((Action)delegate()
                            {
                                SongMaintenance_Tooltip_Label.Text = "總共重建 " + Global.TotalList[0] + " 首歌曲,忽略 " + Global.TotalList[1] + " 首,失敗 " + Global.TotalList[2] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成重建。";
                                SongMaintenance_RebuildSongStructure_TextBox.Text = "";
                                SongMaintenance_RebuildSongStructure_Button.Text = "瀏覽";

                                Global.SongMgrDestFolder = RebuildSongPath;
                                SongMgrCfg_DestFolder_TextBox.Text = RebuildSongPath;
                                SongMgrSettings.Save("SongMgrDestFolder");

                                // 統計歌曲數量
                                Task.Factory.StartNew(() => Common_GetSongStatisticsTask());
                                Common_SwitchSetUI(true);
                            });
                            SongMaintenance.DisposeSongDataTable();
                        });
                    }
                }
            }
        }


        private void SongMaintenance_RebuildSongStructureTask(string RebuildSongPath, SongAddMode RebuildMode)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            Global.TotalList = new List<int>() { 0, 0, 0, 0 };
            List<string> RebuildSongFileValueList = new List<string>();

            foreach (DataRow row in Global.SongDT.Rows)
            {
                string SongId = row["Song_Id"].ToString();
                string SongLang = row["Song_Lang"].ToString();
                int SongSingerType = Convert.ToInt32(row["Song_SingerType"]);
                string SongSinger = row["Song_Singer"].ToString();
                string SongSongName = row["Song_SongName"].ToString();
                int SongTrack = Convert.ToInt32(row["Song_Track"]);
                string SongSongType = row["Song_SongType"].ToString();
                string SongFileName = row["Song_FileName"].ToString();
                string SongPath = row["Song_Path"].ToString();

                bool HasInvalidChar = false;
                Regex r = new Regex(@"[\\/:*?<>|" + '"' + "]");
                if (r.IsMatch(SongLang)) HasInvalidChar = true;
                if (r.IsMatch(SongSinger)) HasInvalidChar = true;
                if (r.IsMatch(SongSongName)) HasInvalidChar = true;
                if (r.IsMatch(SongSongType)) HasInvalidChar = true;

                if (SongSingerType < 0 | SongSingerType > 10)
                {
                    SongSingerType = 10;
                    Global.SongLog.Add("【歌庫結構重建】此首歌曲歌手類別數值錯誤,已自動將其數值改為10: " + SongId + "|" + SongSongName);
                }

                if (SongTrack < 0 | SongTrack > 5)
                {
                    SongTrack = 0;
                    Global.SongLog.Add("【歌庫結構重建】此首歌曲聲道數值錯誤,已自動將其數值改為0: " + SongId + "|" + SongSongName);
                }

                // 重建歌檔
                string SongSrcPath = Path.Combine(SongPath, SongFileName);
                string SongDestPath = "";

                if (!HasInvalidChar)
                {
                    SongRecord song = new SongRecord
                    {
                        Id = SongId, Lang = SongLang, SingerType = SongSingerType.ToString(), Singer = SongSinger, SongName = SongSongName,
                        Track = SongTrack.ToString(), SongType = SongSongType, FileName = SongFileName, Path = SongPath,
                    };
                    SongDestPath = LibraryPlacement.Destination(song, PlacementPurpose.Rebuild, RebuildSongPath);
                    SongPath = Path.GetDirectoryName(SongDestPath) + @"\";
                    SongFileName = Path.GetFileName(SongDestPath);
                }

                bool FileIOError = false;
                if (File.Exists(SongSrcPath) & !HasInvalidChar)
                {
                    if (RebuildMode == SongAddMode.Move || RebuildMode == SongAddMode.Copy)
                    {
                        if (LibraryPlacement.Overwrite(SongSrcPath, SongDestPath, RebuildMode == SongAddMode.Copy) != PlacementOutcome.Placed)
                        {
                            FileIOError = true;
                            Global.SongLog.Add("【歌庫結構重建】檔案處理發生錯誤: " + SongSrcPath + " (唯讀或使用中)");
                        }
                    }
                }
                else
                {
                    FileIOError = true;
                    if (HasInvalidChar)
                    {
                        Global.SongLog.Add("【歌庫結構重建】檔案處理發生錯誤: " + SongId + "|" + SongLang + "|" + SongSinger + "|" + SongSongName + "|" + SongSongType + " (含有非法字元)");
                    }
                    else
                    {
                        Global.SongLog.Add("【歌庫結構重建】此首歌曲檔案不存在,已自動忽略重建檔案: " + SongId + "|" + SongSrcPath);
                    }
                    lock (LockThis) { Global.TotalList[1]++; }
                }

                if (!FileIOError)
                {
                    string RebuildSongFileValue = SongId + "|" + SongSingerType + "|" + SongTrack + "|" + SongFileName + "|" + SongPath;
                    RebuildSongFileValueList.Add(RebuildSongFileValue);
                    lock (LockThis) { Global.TotalList[0]++; }

                    this.BeginInvoke((Action)delegate()
                    {
                        SongMaintenance_Tooltip_Label.Text = "已成功將 " + Global.TotalList[0] + " 首歌曲重建至重建資料夾,請稍待...";
                    });
                }
            }

            SongLibraryStore store = SongLibraryStore.Open();

            List<string> valuelist = new List<string>();

            foreach (string str in RebuildSongFileValueList)
            {
                valuelist = new List<string>(str.Split('|'));

                try
                {
                    store.UpdateSong(new SongRecord { Id = valuelist[0], SingerType = valuelist[1], Track = valuelist[2], FileName = valuelist[3], Path = valuelist[4] }, SongColumns.Id | SongColumns.SingerType | SongColumns.Track | SongColumns.FileName | SongColumns.Path);
                    lock (LockThis) { Global.TotalList[3]++; }
                }
                catch
                {
                    Global.SongLog.Add("【歌庫轉換】寫入重建檔案路徑至資料庫時發生錯誤: " + valuelist[0] + "|" + valuelist[3] + "|" + valuelist[4]);
                    lock (LockThis)
                    {
                        Global.TotalList[0]--;
                        Global.TotalList[2]++;
                    }
                }

                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = "正在更新第 " + Global.TotalList[3] + " 首歌曲的資料庫資料,請稍待...";
                });
            }
            RebuildSongFileValueList.Clear();
            store.Dispose();
        }


        #endregion


        #region --- 我的最愛 ---

        private void SongMaintenance_GetFavoriteUserList()
        {
            SongMaintenance_Favorite_ListBox.DataSource = CommonFunc.GetFavoriteUserList(1);
            SongMaintenance_Favorite_ListBox.DisplayMember = "Display";
            SongMaintenance_Favorite_ListBox.ValueMember = "Value";
            SongMaintenance_Favorite_ListBox.SelectedValue = 1;
        }

        private void SongMaintenance_Favorite_ListBox_Enter(object sender, EventArgs e)
        {
            SongMaintenance_Tooltip_Label.Text = "";
            SongMaintenance_Favorite_Button.Text = "移除";
        }

        private void SongMaintenance_Favorite_TextBox_Enter(object sender, EventArgs e)
        {
            SongMaintenance_Favorite_TextBox.ImeMode = ImeMode.OnHalf;
            SongMaintenance_Tooltip_Label.Text = "";
            SongMaintenance_Favorite_Button.Text = "加入";
        }

        private void SongMaintenance_Favorite_Button_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string SongQuerySqlStr = "";
            string UserId = "";
            string UserName = "";

            switch (SongMaintenance_Favorite_Button.Text)
            {
                case "加入":
                    if (SongMaintenance_Favorite_TextBox.Text != "")
                    {
                        if (SongMaintenance_Tooltip_Label.Text == "尚未輸入要加入的最愛用戶名稱!") SongMaintenance_Tooltip_Label.Text = "";
                        bool AddUser = true;
                        dt = (DataTable)SongMaintenance_Favorite_ListBox.DataSource;

                        if (dt.Rows.Count > 0)
                        {
                            if (SongMaintenance_Tooltip_Label.Text == "已有此最愛用戶名稱!") SongMaintenance_Tooltip_Label.Text = "";
                            foreach (DataRow row in dt.AsEnumerable())
                            {
                                if (row["Display"].ToString() == SongMaintenance_Favorite_TextBox.Text)
                                {
                                    SongMaintenance_Tooltip_Label.Text = "已有此最愛用戶名稱!";
                                    dt = new DataTable();
                                    AddUser = false;
                                    break;
                                }
                            }
                        }

                        if (AddUser)
                        {
                            SongQuerySqlStr = "select User_Id, User_Name from ktv_User";
                            dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");

                            List<string> UserIdList = new List<string>();
                            foreach (DataRow row in dt.Rows)
                            {
                                if (CommonFunc.IsUserId(row["User_Id"].ToString()))
                                {
                                    UserIdList.Add(row["User_Id"].ToString());
                                }
                            }

                            for (int i = 1; i < 9999; i++)
                            {
                                if (UserIdList.IndexOf(i.ToString("D4")) < 0)
                                {
                                    UserId = i.ToString("D4");
                                    UserName = SongMaintenance_Favorite_TextBox.Text;
                                    break;
                                }
                            }

                            if (UserId != "")
                            {
                                using (SongLibraryStore store = SongLibraryStore.Open())
                                {
                                    store.AddFavoriteUser(UserId, UserName);
                                }
                                SongMaintenance_Favorite_TextBox.Text = "";

                                SongQuery_GetFavoriteUserList();
                                SongMaintenance_GetFavoriteUserList();
                                if (Global.SongQueryQueryType == "FavoriteQuery")
                                {
                                    Global.SongQueryQueryType = "SongQuery";
                                    SongQuery_DataGridView.DataSource = null;
                                    if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                                    SongQuery_QueryStatus_Label.Text = "";
                                }
                            }
                            else
                            {
                                SongMaintenance_Tooltip_Label.Text = "最愛用戶數量已滿!";
                            }
                        }
                    }
                    else
                    {
                        SongMaintenance_Tooltip_Label.Text = "尚未輸入要加入的最愛用戶名稱!";
                    }
                    break;
                case "移除":
                    UserName = SongMaintenance_Favorite_ListBox.Text;
                    if (UserName == "")
                    {
                        SongMaintenance_Tooltip_Label.Text = "已無最愛用戶可移除!";
                    }
                    else
                    {
                        if (SongMaintenance_Tooltip_Label.Text == "已無最愛用戶可移除!") SongMaintenance_Tooltip_Label.Text = "";
                        if (MessageBox.Show("你確定要移除此最愛用戶【" + UserName + "】及其所有最愛歌曲嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            SongQuerySqlStr = "select User_Id, User_Name from ktv_User";
                            dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");

                            var query = from row in dt.AsEnumerable()
                                        where row.Field<string>("User_Name").Equals(UserName)
                                        select row;

                            if (query.Count<DataRow>() > 0)
                            {
                                foreach (DataRow row in query)
                                {
                                    UserId = row["User_Id"].ToString();
                                    break;
                                }
                            }

                            if (UserId != "")
                            {
                                using (SongLibraryStore store = SongLibraryStore.Open())
                                {
                                    store.RemoveFavoriteUser(UserId);
                                }
                                SongQuery_GetFavoriteUserList();
                                SongMaintenance_GetFavoriteUserList();
                                if (Global.SongQueryQueryType == "FavoriteQuery")
                                {
                                    Global.SongQueryQueryType = "SongQuery";
                                    SongQuery_DataGridView.DataSource = null;
                                    if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                                    SongQuery_QueryStatus_Label.Text = "";
                                }
                            }
                        }
                    }
                    break;
            }
            dt.Dispose();
        }


        private void SongMaintenance_FavoriteExport_Button_Click(object sender, EventArgs e)
        {
            SongMaintenance.CreateSongDataTable();

            var tasks = new List<Task>()
            {
                Task.Factory.StartNew(() => SongMaintenance_FavoriteExportTask())
            };

            Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
            {
                this.BeginInvoke((Action)delegate()
                {
                    SongMaintenance_Tooltip_Label.Text = @"已將我的最愛資料匯出至【SongMgr\Backup\Favorite.txt】檔案。";
                    SongMaintenance.DisposeSongDataTable();
                });
            });
        }

        public static void SongMaintenance_FavoriteExportTask()
        {
            List<string> list = new List<string>();
            List<string> UserIDList = new List<string>();

            string SongQuerySqlStr = "select User_Id, User_Name from ktv_User";
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, ""))
            {
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.AsEnumerable())
                    {
                        string UserID = Convert.ToString(row["User_Id"]);
                        Regex r = new Regex(@"^\*\*\*\*$|^####$|^\^CB\d$|^\^CN\d{6}$|^\^CG\d$|^9999$");
                        if (!r.IsMatch(UserID) && !Convert.ToString(row["User_Name"]).Contains("新歌快報"))
                        {
                            UserIDList.Add(UserID);
                            r = new Regex(@"^\^#\d{4}$");
                            if (r.IsMatch(UserID))
                            {
                                UserID = Regex.Replace(UserID, @"\^#", "");
                            }
                        }
                        list.Add("ktv_User|" + UserID + "|" + Convert.ToString(row["User_Name"]));
                    }
                }
            }

            SongQuerySqlStr = "select User_Id, Song_Id from ktv_Favorite";
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, ""))
            {
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.AsEnumerable())
                    {
                        var query = from QueryRow in Global.SongDT.AsEnumerable()
                                    where QueryRow.Field<string>("Song_Id").Equals(row["Song_Id"].ToString())
                                    select QueryRow;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow songrow in query)
                            {
                                string UserID = Convert.ToString(row["User_Id"]);
                                Regex r = new Regex(@"^\*\*\*\*$|^####$|^\^CB\d$|^\^CN\d{6}$|^\^CG\d$|^9999$");
                                if (!r.IsMatch(UserID) && UserIDList.IndexOf(UserID) >= 0)
                                {
                                    r = new Regex(@"^\^#\d{4}$");
                                    if (r.IsMatch(UserID))
                                    {
                                        UserID = Regex.Replace(UserID, @"\^#", "");
                                    }
                                }
                                list.Add("ktv_Favorite|" + UserID + "|" + Convert.ToString(songrow["Song_Lang"]) + "|" + Convert.ToString(songrow["Song_Singer"]) + "|" + Convert.ToString(songrow["Song_SongName"]));
                                break;
                            }
                        }
                    }
                }
            }

            if (YoutubeFavorites.TableExists(Global.CrazyktvDatabaseFile))
            {
                SongQuerySqlStr = "select User_Id, VideoId, Name, Length, CreatDate from YoutubeFavorite order by CreatDate, VideoId";
                using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, ""))
                {
                    foreach (DataRow row in dt.AsEnumerable())
                    {
                        string UserID = Convert.ToString(row["User_Id"]);
                        if (UserIDList.IndexOf(UserID) >= 0 && Regex.IsMatch(UserID, @"^\^#\d{4}$")) UserID = Regex.Replace(UserID, @"\^#", "");
                        list.Add(SongMaintenance.YoutubeFavoriteLine(UserID, row));
                    }
                }
            }
            UserIDList.Clear();
            UserIDList = null;

            if (!Directory.Exists(Application.StartupPath + @"\SongMgr\Backup")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Backup");
            StreamWriter sw = new StreamWriter(Application.StartupPath + @"\SongMgr\Backup\Favorite.txt");
            foreach (string str in list)
            {
                sw.WriteLine(str);
            }
            sw.Close();
            list.Clear();
            list = null;
        }

        private void SongMaintenance_FavoriteImport_Button_Click(object sender, EventArgs e)
        {
            if (File.Exists(Application.StartupPath + @"\SongMgr\Backup\Favorite.txt"))
            {
                if (SongMaintenance_Tooltip_Label.Text == @"【SongMgr\Backup\Favorite.txt】我的最愛備份檔案不存在!") SongMaintenance_Tooltip_Label.Text = "";
                if (MessageBox.Show("你確定要重置並匯入我的最愛嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Global.TimerStartTime = DateTime.Now;
                    Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                    SongMaintenance.CreateSongDataTable();
                    Common_SwitchSetUI(false);

                    SongMaintenance_Tooltip_Label.Text = "正在匯入我的最愛,請稍待...";

                    var tasks = new List<Task<bool>>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_FavoriteImportTask())
                    };

                    Task.Factory.ContinueWhenAll(tasks.ToArray(), FavoriteImportEndTask =>
                    {
                        Global.TimerEndTime = DateTime.Now;
                        bool imported = FavoriteImportEndTask[0].Result;
                        this.BeginInvoke((Action)delegate()
                        {
                            SongMaintenance_Tooltip_Label.Text = imported
                                ? "總共匯入 " + Global.TotalList[0] + " 位最愛用戶及 " + Global.TotalList[1] + " 首最愛歌曲,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。"
                                : "匯入我的最愛失敗,我的最愛維持原樣。";
                            Common_SwitchSetUI(true);

                            SongQuery_GetFavoriteUserList();
                            SongMaintenance_GetFavoriteUserList();
                            if (Global.SongQueryQueryType == "FavoriteQuery")
                            {
                                Global.SongQueryQueryType = "SongQuery";
                                SongQuery_DataGridView.DataSource = null;
                                if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                                SongQuery_QueryStatus_Label.Text = "";
                            }
                            SongMaintenance.DisposeSongDataTable();
                        });
                    });
                }
            }
            else
            {
                SongMaintenance_Tooltip_Label.Text = @"【SongMgr\Backup\Favorite.txt】我的最愛備份檔案不存在!";
            }
        }


        /// <summary>Returns false, with 我的最愛 left as they were, when the backup could not be read or written.</summary>
        private bool SongMaintenance_FavoriteImportTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            try
            {
                string[] lines = File.ReadAllLines(Application.StartupPath + @"\SongMgr\Backup\Favorite.txt", Encoding.UTF8);
                using (SongLibraryStore store = SongLibraryStore.OpenOrThrow())
                {
                    var written = SongMaintenance.ImportFavorites(store, lines, Global.SongDT);
                    Global.TotalList[0] = written.Users;
                    Global.TotalList[1] = written.Songs;
                }
                return true;
            }
            catch (Exception ex)
            {
                Global.SongLog.Add("【我的最愛匯入】匯入失敗,我的最愛維持原樣: " + ex.Message);
                return false;
            }
        }


        #endregion


        #region --- 錢櫃排行 (我的最愛) ---


        private void SongMaintenance_Favorite_UpdateNewbill_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要更新錢櫃新歌排行榜至我的最愛嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在更新錢櫃新歌排行榜至我的最愛,請稍待...";

                var tasks = new List<Task>()
                {
                        Task.Factory.StartNew(() => SongMaintenance_Favorite_UpdateNewbillTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), FavoriteImportEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 位最愛用戶及 " + Global.TotalList[1] + " 首排行歌曲,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);

                        SongQuery_GetFavoriteUserList();
                        SongMaintenance_GetFavoriteUserList();
                        if (Global.SongQueryQueryType == "FavoriteQuery")
                        {
                            Global.SongQueryQueryType = "SongQuery";
                            SongQuery_DataGridView.DataSource = null;
                            if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                            SongQuery_QueryStatus_Label.Text = "";
                        }
                        Cashbox.DisposeSongDataTable();
                    });
                });
            }
        }


        private void SongMaintenance_Favorite_UpdateNewbillTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                List<string> clist = new List<string>();
                List<string> tlist = new List<string>();

                if (CommonFunc.CheckFor45PlusVersion())
                {
                    string url = "https://www.cashboxparty.com/ashx/ah_Song.ashx";
                    string poststr = string.Format("[{0}\"m\":\"bl\",\"t\":\"1\",\"lang\":\"01,02\"{1}]", "{", "}");
                    string jsonString = CommonFunc.DownloadCashboxData(url, poststr);

                    using (DataTable dt = CommonFunc.JSONtoDataTable(jsonString))
                    {
                        if (dt.Rows.Count > 0 && dt.Columns.Contains("SongName"))
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                List<string> list = new List<string>()
                                 {
                                     row["SongNo"].ToString(),
                                     row["LangName"].ToString(),
                                     row["SongName"].ToString(),
                                     row["Singer"].ToString(),
                                     row["LangCode"].ToString()
                                 };

                                if (CommonFunc.IsSongId(list[0]) && list[1] != "" && list[2] != "" && list[3] != "" && list[4] != "")
                                {
                                    switch (list[4])
                                    {
                                        case "01":
                                            clist.Add(list[0]);
                                            break;
                                        case "02":
                                            tlist.Add(list[0]);
                                            break;
                                    }
                                }
                                list.Clear();
                                list = null;
                            }
                        }
                    }
                }
                else
                {
                    string url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_newbill.json";
                    MemoryStream ms = CommonFunc.Download(url);
                    try
                    {
                        if (ms.Length > 0)
                        {
                            ms.Position = 0;
                            string jsonString = string.Empty;

                            using (StreamReader sr = new StreamReader(ms))
                            {
                                jsonString = sr.ReadToEnd();
                                DataTable dt = CommonFunc.JSONtoDataTable(jsonString);

                                if (dt.Rows.Count > 0)
                                {
                                    var cquery = from row in dt.AsEnumerable()
                                                where row.Field<string>("Song_Lang") == "國語"
                                                select row;

                                    if (cquery.Count<DataRow>() > 0)
                                    {
                                        string CashBoxID = string.Empty;
                                        foreach (DataRow row in cquery)
                                        {
                                            CashBoxID = row["Cashbox_Id"].ToString();
                                            if (CommonFunc.IsSongId(CashBoxID)) clist.Add(CashBoxID);
                                        }
                                    }

                                    var tquery = from row in dt.AsEnumerable()
                                                 where row.Field<string>("Song_Lang") == "台語"
                                                 select row;

                                    if (tquery.Count<DataRow>() > 0)
                                    {
                                        string CashBoxID = string.Empty;
                                        foreach (DataRow row in tquery)
                                        {
                                            CashBoxID = row["Cashbox_Id"].ToString();
                                            if (CommonFunc.IsSongId(CashBoxID)) tlist.Add(CashBoxID);
                                        }
                                    }
                                }
                                ms = null;
                            }
                        }
                    }
                    finally
                    {
                        ms?.Dispose();
                    }
                }

                List<string> CUpdateList = new List<string>();
                List<string> TUpdateList = new List<string>();
                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    // 國語新歌排行
                    if (clist.Count > 0)
                    {
                        CUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, clist, 1);
                    }
                    clist.Clear();
                    clist = null;

                    // 台語新歌排行
                    if (tlist.Count > 0)
                    {
                        TUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, tlist, 1);
                    }
                    tlist.Clear();
                    tlist = null;
                }

                if (CUpdateList.Count > 0 || TUpdateList.Count > 0)
                {
                    SongMaintenance.ReplaceFavoriteUsers(new List<(string UserId, string Name, List<string> SongIds)>
                    {
                        ("^CB1", "錢櫃新歌排行榜 (國語)", CUpdateList),
                        ("^CB2", "錢櫃新歌排行榜 (台語)", TUpdateList),
                    });
                }
                CUpdateList.Clear();
                TUpdateList.Clear();
            }
        }


        private void SongMaintenance_Favorite_UpdateTotalbill_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要更新錢櫃點播總排行至我的最愛嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在更新錢櫃點播總排行至我的最愛,請稍待...";

                var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_Favorite_UpdateTotalbillTask())
                    };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), FavoriteImportEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 位最愛用戶及 " + Global.TotalList[1] + " 首排行歌曲,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);

                        SongQuery_GetFavoriteUserList();
                        SongMaintenance_GetFavoriteUserList();
                        if (Global.SongQueryQueryType == "FavoriteQuery")
                        {
                            Global.SongQueryQueryType = "SongQuery";
                            SongQuery_DataGridView.DataSource = null;
                            if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                            SongQuery_QueryStatus_Label.Text = "";
                        }
                        Cashbox.DisposeSongDataTable();
                    });
                });
            }
        }


        private void SongMaintenance_Favorite_UpdateTotalbillTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                List<string> clist = new List<string>();
                List<string> tlist = new List<string>();

                if (CommonFunc.CheckFor45PlusVersion())
                {
                    string url = "https://www.cashboxparty.com/ashx/ah_Song.ashx";
                    string poststr = string.Format("[{0}\"m\":\"bl\",\"t\":\"2\",\"lang\":\"01,02\"{1}]", "{", "}");
                    string jsonString = CommonFunc.DownloadCashboxData(url, poststr);

                    using (DataTable dt = CommonFunc.JSONtoDataTable(jsonString))
                    {
                        if (dt.Rows.Count > 0 && dt.Columns.Contains("SongName"))
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                List<string> list = new List<string>()
                                 {
                                     row["SongNo"].ToString(),
                                     row["LangName"].ToString(),
                                     row["SongName"].ToString(),
                                     row["Singer"].ToString(),
                                     row["LangCode"].ToString()
                                 };

                                if (CommonFunc.IsSongId(list[0]) && list[1] != "" && list[2] != "" && list[3] != "" && list[4] != "")
                                {
                                    switch (list[4])
                                    {
                                        case "01":
                                            clist.Add(list[0]);
                                            break;
                                        case "02":
                                            tlist.Add(list[0]);
                                            break;
                                    }
                                }
                                list.Clear();
                                list = null;
                            }
                        }
                    }
                }
                else
                {
                    string url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_totalbill.json";

                    MemoryStream ms = CommonFunc.Download(url);
                    try
                    {
                        if (ms.Length > 0)
                        {
                            ms.Position = 0;
                            string jsonString = string.Empty;

                            using (StreamReader sr = new StreamReader(ms))
                            {
                                jsonString = sr.ReadToEnd();
                                DataTable dt = CommonFunc.JSONtoDataTable(jsonString);

                                if (dt.Rows.Count > 0)
                                {
                                    var cquery = from row in dt.AsEnumerable()
                                                 where row.Field<string>("Song_Lang") == "國語"
                                                 select row;

                                    if (cquery.Count<DataRow>() > 0)
                                    {
                                        string CashBoxID = string.Empty;
                                        foreach (DataRow row in cquery)
                                        {
                                            CashBoxID = row["Cashbox_Id"].ToString();
                                            if (CommonFunc.IsSongId(CashBoxID)) clist.Add(CashBoxID);
                                        }
                                    }

                                    var tquery = from row in dt.AsEnumerable()
                                                 where row.Field<string>("Song_Lang") == "台語"
                                                 select row;

                                    if (tquery.Count<DataRow>() > 0)
                                    {
                                        string CashBoxID = string.Empty;
                                        foreach (DataRow row in tquery)
                                        {
                                            CashBoxID = row["Cashbox_Id"].ToString();
                                            if (CommonFunc.IsSongId(CashBoxID)) tlist.Add(CashBoxID);
                                        }
                                    }
                                }
                                ms = null;
                            }
                        }
                    }
                    finally
                    {
                        ms?.Dispose();
                    }
                }

                List<string> CUpdateList = new List<string>();
                List<string> TUpdateList = new List<string>();
                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    // 國語新歌排行
                    if (clist.Count > 0)
                    {
                        CUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, clist, 1);
                    }
                    clist.Clear();
                    clist = null;

                    // 台語新歌排行
                    if (tlist.Count > 0)
                    {
                        TUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, tlist, 1);
                    }
                    tlist.Clear();
                    tlist = null;
                }

                if (CUpdateList.Count > 0 || TUpdateList.Count > 0)
                {
                    SongMaintenance.ReplaceFavoriteUsers(new List<(string UserId, string Name, List<string> SongIds)>
                    {
                        ("^CB3", "錢櫃國語排行榜", CUpdateList),
                        ("^CB4", "錢櫃台語排行榜", TUpdateList),
                    });
                }
                CUpdateList.Clear();
                TUpdateList.Clear();
            }
        }


        private void SongMaintenance_Favorite_UpdateOtherlangbill_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要更新錢櫃外語排行榜至我的最愛嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在更新錢櫃外語排行榜至我的最愛,請稍待...";

                var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_Favorite_UpdateOtherlangbillTask())
                    };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), FavoriteImportEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 位最愛用戶及 " + Global.TotalList[1] + " 首排行歌曲,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);

                        SongQuery_GetFavoriteUserList();
                        SongMaintenance_GetFavoriteUserList();
                        if (Global.SongQueryQueryType == "FavoriteQuery")
                        {
                            Global.SongQueryQueryType = "SongQuery";
                            SongQuery_DataGridView.DataSource = null;
                            if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                            SongQuery_QueryStatus_Label.Text = "";
                        }
                        Cashbox.DisposeSongDataTable();
                    });
                });
            }
        }


        private void SongMaintenance_Favorite_UpdateOtherlangbillTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                List<string> hlist = new List<string>();
                List<string> elist = new List<string>();
                List<string> jlist = new List<string>();

                if (CommonFunc.CheckFor45PlusVersion())
                {
                    string url = "https://www.cashboxparty.com/ashx/ah_Song.ashx";
                    string poststr = string.Format("[{0}\"m\":\"bl\",\"t\":\"2\",\"lang\":\"03,04,05\"{1}]", "{", "}");
                    string jsonString = CommonFunc.DownloadCashboxData(url, poststr);

                    using (DataTable dt = CommonFunc.JSONtoDataTable(jsonString))
                    {
                        if (dt.Rows.Count > 0 && dt.Columns.Contains("SongName"))
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                List<string> list = new List<string>()
                                 {
                                     row["SongNo"].ToString(),
                                     row["LangName"].ToString(),
                                     row["SongName"].ToString(),
                                     row["Singer"].ToString(),
                                     row["LangCode"].ToString()
                                 };

                                if (CommonFunc.IsSongId(list[0]) && list[1] != "" && list[2] != "" && list[3] != "" && list[4] != "")
                                {
                                    switch (list[4])
                                    {
                                        case "03":
                                            hlist.Add(list[0]);
                                            break;
                                        case "04":
                                            jlist.Add(list[0]);
                                            break;
                                        case "05":
                                            elist.Add(list[0]);
                                            break;
                                    }
                                }
                                list.Clear();
                                list = null;
                            }
                        }
                    }
                }
                else
                {
                    string url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_cantonesebill.json";
                    MemoryStream ms = CommonFunc.Download(url);
                    try
                    {
                        if (ms.Length > 0)
                        {
                            ms.Position = 0;
                            string jsonString = string.Empty;

                            using (StreamReader sr = new StreamReader(ms))
                            {
                                jsonString = sr.ReadToEnd();
                                DataTable dt = CommonFunc.JSONtoDataTable(jsonString);

                                if (dt.Rows.Count > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) hlist.Add(CashBoxID);
                                    }
                                }
                                ms = null;
                            }
                        }
                    }
                    finally
                    {
                        ms?.Dispose();
                    }
                    
                    url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_englishbill.json";
                    ms = CommonFunc.Download(url);
                    try
                    {
                        if (ms.Length > 0)
                        {
                            ms.Position = 0;
                            string jsonString = string.Empty;

                            using (StreamReader sr = new StreamReader(ms))
                            {
                                jsonString = sr.ReadToEnd();
                                DataTable dt = CommonFunc.JSONtoDataTable(jsonString);

                                if (dt.Rows.Count > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) elist.Add(CashBoxID);
                                    }
                                }
                                ms = null;
                            }
                        }
                    }
                    finally
                    {
                        ms?.Dispose();
                    }

                    url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_japanesebill.json";
                    ms = CommonFunc.Download(url);
                    try
                    {
                        if (ms.Length > 0)
                        {
                            ms.Position = 0;
                            string jsonString = string.Empty;

                            using (StreamReader sr = new StreamReader(ms))
                            {
                                jsonString = sr.ReadToEnd();
                                DataTable dt = CommonFunc.JSONtoDataTable(jsonString);

                                if (dt.Rows.Count > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) jlist.Add(CashBoxID);
                                    }
                                }
                                ms = null;
                            }
                        }
                    }
                    finally
                    {
                        ms?.Dispose();
                    }
                }

                List<string> HUpdateList = new List<string>();
                List<string> EUpdateList = new List<string>();
                List<string> JUpdateList = new List<string>();
                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    // 粵語新歌排行
                    if (hlist.Count > 0)
                    {
                        HUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, hlist, 1);
                    }
                    hlist.Clear();
                    hlist = null;

                    // 英語新歌排行
                    if (elist.Count > 0)
                    {
                        EUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, elist, 1);
                    }
                    elist.Clear();
                    elist = null;
                    // 日語新歌排行
                    if (jlist.Count > 0)
                    {
                        JUpdateList = SongMaintenance.MatchCashboxData(CashboxDT, jlist, 1);
                    }
                    jlist.Clear();
                    jlist = null;
                }

                if (HUpdateList.Count > 0 || EUpdateList.Count > 0 || JUpdateList.Count > 0)
                {
                    SongMaintenance.ReplaceFavoriteUsers(new List<(string UserId, string Name, List<string> SongIds)>
                    {
                        ("^CB5", "錢櫃粵語排行榜", HUpdateList),
                        ("^CB6", "錢櫃英語排行榜", EUpdateList),
                        ("^CB7", "錢櫃日語排行榜", JUpdateList),
                    });
                }
                HUpdateList.Clear();
                EUpdateList.Clear();
                JUpdateList.Clear();
            }
        }


        private void SongMaintenance_Favorite_Update3456Gold_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要更新錢櫃懷念金曲至我的最愛嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在更新錢櫃懷念金曲至我的最愛,請稍待...";

                var tasks = new List<Task>()
                    {
                        Task.Factory.StartNew(() => SongMaintenance_Favorite_Update3456GoldTask())
                    };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), FavoriteImportEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 位最愛用戶及 " + Global.TotalList[1] + " 首排行歌曲,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);

                        SongQuery_GetFavoriteUserList();
                        SongMaintenance_GetFavoriteUserList();
                        if (Global.SongQueryQueryType == "FavoriteQuery")
                        {
                            Global.SongQueryQueryType = "SongQuery";
                            SongQuery_DataGridView.DataSource = null;
                            if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                            SongQuery_QueryStatus_Label.Text = "";
                        }
                        Cashbox.DisposeSongDataTable();
                    });
                });
            }
        }


        private void SongMaintenance_Favorite_Update3456GoldTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                List<string> c34list = new List<string>();
                List<string> t34list = new List<string>();
                List<string> c56list = new List<string>();
                List<string> t56list = new List<string>();

                string url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_3456goldsong.json";
                MemoryStream ms = CommonFunc.Download(url);
                try
                {
                    if (ms.Length > 0)
                    {
                        ms.Position = 0;
                        string jsonString = string.Empty;

                        using (StreamReader sr = new StreamReader(ms))
                        {
                            jsonString = sr.ReadToEnd();
                            DataTable dt = CommonFunc.JSONtoDataTable(jsonString);

                            if (dt.Rows.Count > 0)
                            {
                                var query = from row in dt.AsEnumerable()
                                            where row.Field<string>("Song_Type") == "1"
                                            select row;

                                if (query.Count<DataRow>() > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in query)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) t34list.Add(CashBoxID);
                                    }
                                }

                                query = from row in dt.AsEnumerable()
                                        where row.Field<string>("Song_Type") == "2"
                                        select row;

                                if (query.Count<DataRow>() > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in query)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) c34list.Add(CashBoxID);
                                    }
                                }

                                query = from row in dt.AsEnumerable()
                                        where row.Field<string>("Song_Type") == "3"
                                        select row;

                                if (query.Count<DataRow>() > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in query)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) t56list.Add(CashBoxID);
                                    }
                                }

                                query = from row in dt.AsEnumerable()
                                        where row.Field<string>("Song_Type") == "4"
                                        select row;

                                if (query.Count<DataRow>() > 0)
                                {
                                    string CashBoxID = string.Empty;
                                    foreach (DataRow row in query)
                                    {
                                        CashBoxID = row["Cashbox_Id"].ToString();
                                        if (CommonFunc.IsSongId(CashBoxID)) c56list.Add(CashBoxID);
                                    }
                                }
                            }
                            ms = null;
                        }
                    }
                }
                finally
                {
                    ms?.Dispose();
                }

                List<string> C34UpdateList = new List<string>();
                List<string> T34UpdateList = new List<string>();
                List<string> C56UpdateList = new List<string>();
                List<string> T56UpdateList = new List<string>();
                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    // 三四年級國語金曲
                    if (c34list.Count > 0)
                    {
                        C34UpdateList = SongMaintenance.MatchCashboxData(CashboxDT, c34list, 1);
                    }
                    c34list.Clear();
                    c34list = null;

                    // 三四年級台語金曲
                    if (t34list.Count > 0)
                    {
                        T34UpdateList = SongMaintenance.MatchCashboxData(CashboxDT, t34list, 1);
                    }
                    t34list.Clear();
                    t34list = null;

                    // 五六年級國語金曲
                    if (c56list.Count > 0)
                    {
                        C56UpdateList = SongMaintenance.MatchCashboxData(CashboxDT, c56list, 1);
                    }
                    c56list.Clear();
                    c56list = null;

                    // 五六年級台語金曲
                    if (t56list.Count > 0)
                    {
                        T56UpdateList = SongMaintenance.MatchCashboxData(CashboxDT, t56list, 1);
                    }
                    t56list.Clear();
                    t56list = null;
                }

                if (C34UpdateList.Count > 0 || T34UpdateList.Count > 0 || C56UpdateList.Count > 0 || T56UpdateList.Count > 0)
                {
                    SongMaintenance.ReplaceFavoriteUsers(new List<(string UserId, string Name, List<string> SongIds)>
                    {
                        ("^CG1", "三四年級國語金曲", C34UpdateList),
                        ("^CG2", "三四年級台語金曲", T34UpdateList),
                        ("^CG3", "五六年級國語金曲", C56UpdateList),
                        ("^CG4", "五六年級台語金曲", T56UpdateList),
                    });
                }
                C34UpdateList.Clear();
                C34UpdateList = null;
                T34UpdateList.Clear();
                T34UpdateList = null;
                C56UpdateList.Clear();
                C56UpdateList = null;
                T56UpdateList.Clear();
                T56UpdateList = null;
            }
        }


        private void SongMaintenance_Favorite_UpdateNewsong_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要更新錢櫃新歌快報至我的最愛嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在更新錢櫃新歌快報至我的最愛,請稍待...";

                string DateItem = SongMaintenance_Favorite_UpdateNewsong_ComboBox.Text;

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_Favorite_UpdateNewsongTask(DateItem))
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), FavoriteImportEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 位最愛用戶及 " + Global.TotalList[1] + " 首排行歌曲,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);

                        SongQuery_GetFavoriteUserList();
                        SongMaintenance_GetFavoriteUserList();
                        if (Global.SongQueryQueryType == "FavoriteQuery")
                        {
                            Global.SongQueryQueryType = "SongQuery";
                            SongQuery_DataGridView.DataSource = null;
                            if (SongQuery_DataGridView.Columns.Count > 0) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
                            SongQuery_QueryStatus_Label.Text = "";
                        }
                        Cashbox.DisposeSongDataTable();
                    });
                });
            }
        }


        private void SongMaintenance_Favorite_UpdateNewsongTask(string DateItem)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                List<string> UpdateDateList = new List<string>();
                List<string> UpdateCashboxIdList = new List<string>();
                string UpdateDate = string.Empty;
                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    if (DateItem != "")
                    {
                        var query = from row in Global.CashboxNewSongDateDT.AsEnumerable()
                                    where row.Field<string>("Date").Equals(DateItem)
                                    select row;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow row in query)
                            {
                                UpdateDateList.AddRange(row["Value"].ToString().Split('|'));
                            }
                        }
                    }

                    if (UpdateDateList.Count > 0)
                    {
                        UpdateDate = "^CN" +  DateTime.Parse(UpdateDateList[0]).ToString("yyyyMM", CultureInfo.InvariantCulture);
                        foreach (string date in UpdateDateList)
                        {
                            var query = from row in CashboxDT.AsEnumerable()
                                        where row.Field<DateTime>("Song_CreatDate").ToString("yyyy/MM/dd", CultureInfo.InvariantCulture).Equals(date)
                                        select row;
                            
                            if (query.Count<DataRow>() > 0)
                            {
                                foreach (DataRow row in query)
                                {
                                    string SongId = row["Cashbox_Id"].ToString();
                                    if (UpdateCashboxIdList.IndexOf(SongId) < 0)
                                    {
                                        UpdateCashboxIdList.Add(SongId);
                                    }
                                }
                            }
                        }
                    }
                    UpdateDateList.Clear();

                    List<string> UpdateList = new List<string>();
                    if (UpdateCashboxIdList.Count > 0)
                    {
                        UpdateList = SongMaintenance.MatchCashboxData(CashboxDT, UpdateCashboxIdList, 1);
                    }
                    UpdateCashboxIdList.Clear();

                    SongMaintenance.ReplaceFavoriteUsers(new List<(string UserId, string Name, List<string> SongIds)>
                    {
                        (UpdateDate, Regex.Replace(DateItem, "份", ""), UpdateList),
                    });
                    UpdateList.Clear();
                }
            }
        }


        #endregion


        #region --- 自訂語系 ---

        private void SongMaintenance_SetCustomLangControl()
        {
            TextBox[] SongMaintenance_Lang_TextBox =
            {
                SongMaintenance_Lang1_TextBox,
                SongMaintenance_Lang2_TextBox,
                SongMaintenance_Lang3_TextBox,
                SongMaintenance_Lang4_TextBox,
                SongMaintenance_Lang5_TextBox,
                SongMaintenance_Lang6_TextBox,
                SongMaintenance_Lang7_TextBox,
                SongMaintenance_Lang8_TextBox,
                SongMaintenance_Lang9_TextBox,
                SongMaintenance_Lang10_TextBox
            };

            for (int i = 0; i < SongMaintenance_Lang_TextBox.Count<TextBox>(); i++)
            {
                SongMaintenance_Lang_TextBox[i].Text = Global.CrazyktvSongLangList[i];
            }

            TextBox[] SongMaintenance_LangIDStr_TextBox =
            {
              SongMaintenance_Lang1IDStr_TextBox,
              SongMaintenance_Lang2IDStr_TextBox,
              SongMaintenance_Lang3IDStr_TextBox,
              SongMaintenance_Lang4IDStr_TextBox,
              SongMaintenance_Lang5IDStr_TextBox,
              SongMaintenance_Lang6IDStr_TextBox,
              SongMaintenance_Lang7IDStr_TextBox,
              SongMaintenance_Lang8IDStr_TextBox,
              SongMaintenance_Lang9IDStr_TextBox,
              SongMaintenance_Lang10IDStr_TextBox
            };

            for (int i = 0; i < SongMaintenance_LangIDStr_TextBox.Count<TextBox>(); i++)
            {
                SongMaintenance_LangIDStr_TextBox[i].Text = Global.CrazyktvSongLangKeyWordList[i];
            }
        }


        private void SongMaintenance_TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (SongMaintenance_TabControl.SelectedTab.Name)
            {
                case "SongMaintenance_CustomLang_TabPage":
                    SongMaintenance_Save_Button.Text = "更新語系";
                    SongMaintenance_Save_Button.Enabled = true;
                    break;
                case "SongMaintenance_FFmpeg_TabPage":
                    SongMaintenance_Save_Button.Text = "儲存設定";
                    SongMaintenance_Save_Button.Enabled = true;
                    break;
                case "SongMaintenance_DBVer_TabPage":
                    SongMaintenance_Save_Button.Text = "儲存設定";
                    if (SongMaintenance_TabControl.Enabled == true) SongMaintenance_Save_Button.Enabled = true;
                    break;
                default:
                    SongMaintenance_Save_Button.Text = "儲存設定";
                    SongMaintenance_Save_Button.Enabled = false;
                    break;
            }
        }


        private void SongMaintenance_SongLangUpdateTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            List<string> list = new List<string>();

            if (File.Exists(Application.StartupPath + @"\SongMgr\Custom.lang"))
            {
                if (!Directory.Exists(Application.StartupPath + @"\Lang")) Directory.CreateDirectory(Application.StartupPath + @"\Lang");
                File.Copy(Application.StartupPath + @"\SongMgr\Custom.lang", Application.StartupPath + @"\Lang\Custom.lang", true);

                if (File.Exists(Application.StartupPath + @"\Lang\Custom.lang"))
                {
                    StreamReader sr = new StreamReader(Application.StartupPath + @"\Lang\Custom.lang", Encoding.Unicode);
                    while (!sr.EndOfStream)
                    {
                        list.Add(sr.ReadLine());
                    }
                    sr.Close();

                    list[2] = string.Join(",", Global.CrazyktvSongLangList);

                    StreamWriter sw = new StreamWriter(Application.StartupPath + @"\Lang\Custom.lang", false, Encoding.Unicode);
                    foreach (string str in list)
                    {
                        sw.WriteLine(str);
                    }
                    sw.Close();

                    CommonFunc.SaveConfigXmlFile(Global.CrazyktvCfgFile, "Language", "Custom");
                }
            }

            DbConnection conn = Db.OpenConn(Global.CrazyktvDatabaseFile, "");
            string sqlColumnStr = "Langauage_Name = @LangauageName, Langauage_KeyWord = @LangauageKeyWord";
            string SongUpdateSqlStr = "update ktv_Langauage set " + sqlColumnStr + " where Langauage_Id = @LangauageId";
            DbCommand cmd = Db.Command(SongUpdateSqlStr, conn);

            foreach (string str in Global.CrazyktvSongLangList)
            {
                cmd.Parameters.AddWithValue("@LangauageName", str);
                cmd.Parameters.AddWithValue("@LangauageKeyWord", Global.CrazyktvSongLangKeyWordList[Global.CrazyktvSongLangList.IndexOf(str)]);
                cmd.Parameters.AddWithValue("@LangauageId", Global.CrazyktvSongLangList.IndexOf(str));

                try
                {
                    cmd.ExecuteNonQuery();
                    lock (LockThis)
                    {
                        Global.TotalList[0]++;
                    }
                }
                catch
                {
                    lock (LockThis)
                    {
                        Global.TotalList[1]++;
                    }
                    Global.SongLog.Add("【自訂語系】更新資料庫時發生錯誤: " + str);
                }
                cmd.Parameters.Clear();
            }
            conn.Close();
        }


        #endregion


        #region --- FFmpeg - 音量平衡 ---


        private void SongMaintenance_ReplayGain_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要使用音量平衡功能來變更音量嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                if (SongMaintenance_ReplayGain_TextBox.Text != Global.SongMaintenanceReplayGainVolume)
                {
                    Global.SongMaintenanceReplayGainVolume = SongMaintenance_ReplayGain_TextBox.Text;
                    SongMgrSettings.Save("SongMaintenanceReplayGainVolume");
                }

                SongMaintenance_Tooltip_Label.Text = "正在使用音量平衡功能來變更音量,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_ReplayGain_VolumeChangeTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), VolumeChangeEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共更新 " + Global.TotalList[0] + " 首歌曲的音量資料,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }

        private void SongMaintenance_ReplayGain_VolumeChangeTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            
            if (File.Exists(Global.CrazyktvDatabaseFile))
            {
                NativeMethods.SystemSleepManagement.PreventSleep(false);
                int basevolume = Convert.ToInt32(Global.SongMaintenanceReplayGainVolume);

                List<string> list = new List<string>();
                Parallel.ForEach(Global.SongDT.AsEnumerable(), new ParallelOptions { MaxDegreeOfParallelism = 2 }, (row, loopState) =>
                {
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "正在解析第 " + (list.Count + 1) + " 首歌曲的音量資料,請稍待...";
                    });

                    string file = Path.Combine(row["Song_Path"].ToString(), row["Song_FileName"].ToString());
                    string gain = row["Song_ReplayGain"].ToString();
                    double GainDB = 0;

                    string SongVolume = Global.SongMaintenanceReplayGainVolume;

                    if (!string.IsNullOrEmpty(gain))
                    {
                        GainDB = Convert.ToDouble(gain);
                    }
                    else
                    {
                        MediaProbe.SongVolumeValue result = MediaProbe.MeasureGain(file);
                        GainDB = result.GainDB;
                        NativeMethods.SystemSleepManagement.ResetSleepTimer(false);
                    }
                    SongVolume = MediaProbe.SongVolume(basevolume, GainDB);
                    lock (LockThis) list.Add(SongVolume + "|" + GainDB.ToString() + "|" + row["Song_Id"].ToString());
                });

                if (list.Count > 0)
                {
                    using (SongLibraryStore store = SongLibraryStore.Open())
                    {
                        int mCount = 0;
                        int lCount = 0;

                        List<string> valuelist = new List<string>();

                        foreach (string str in list)
                        {
                            valuelist = new List<string>(str.Split('|'));

                            if (Convert.ToInt32(valuelist[0]) > VolumeRange.Library.Max)
                            {
                                valuelist[0] = VolumeRange.Library.Max.ToString();
                                mCount++;
                            }
                            else if (Convert.ToInt32(valuelist[0]) < VolumeRange.Library.Min)
                            {
                                valuelist[0] = VolumeRange.Library.Min.ToString();
                                lCount++;
                            }

                            try
                            {
                                store.UpdateSong(new SongRecord { Id = valuelist[2], Volume = valuelist[0], ReplayGain = valuelist[1] }, SongColumns.Volume | SongColumns.ReplayGain);
                                lock (LockThis) Global.TotalList[0]++;
                            }
                            catch
                            {
                                lock (LockThis) Global.TotalList[1]++;

                                Global.SongLog.Add("【音量平衡】更新資料庫時發生錯誤: " + str);
                            }

                            valuelist.Clear();
                            valuelist = null;

                            this.BeginInvoke((Action)delegate ()
                            {
                                SongMaintenance_Tooltip_Label.Text = "正在轉換第 " + Global.TotalList[0] + " 首歌曲的音量資料,請稍待...";
                            });
                        }
                        list.Clear();
                        list = null;

                        Global.SongLog.Add("【音量平衡】檢測到 " + mCount.ToString() + " 首音量大於 100 的歌曲,已自動變更為 100。");
                        Global.SongLog.Add("【音量平衡】檢測到 " + lCount.ToString() + " 首音量小於 10 的歌曲,已自動變更為 10。");
                    }
                }
                NativeMethods.SystemSleepManagement.ResotreSleep();
            }
        }

        private void SongMaintenance_ReplayGainReset_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要重置偵測出來的播放增益資料嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                if (SongMaintenance_ReplayGain_TextBox.Text != Global.SongMaintenanceReplayGainVolume)
                {
                    Global.SongMaintenanceReplayGainVolume = SongMaintenance_ReplayGain_TextBox.Text;
                    SongMgrSettings.Save("SongMaintenanceReplayGainVolume");
                }

                SongMaintenance_Tooltip_Label.Text = "正在重置播放增益資料,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_ReplayGain_ResetTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), VolumeChangeEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        if (Global.TotalList[0] > 0)
                        {
                            SongMaintenance_Tooltip_Label.Text = "已成功將所有歌曲的播放增益資料重置,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        }
                        else
                        {
                            SongMaintenance_Tooltip_Label.Text = "重置所有歌曲的播放增益資料失敗,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        }
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }

        private void SongMaintenance_ReplayGain_ResetTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            if (File.Exists(Global.CrazyktvDatabaseFile))
            {
                using (SongLibraryStore store = SongLibraryStore.Open())
                {
                    try
                    {
                        store.ClearAllReplayGain();
                        lock (LockThis) Global.TotalList[0]++;
                    }
                    catch
                    {
                        lock (LockThis) Global.TotalList[1]++;
                        Global.SongLog.Add("【音量平衡】更新資料庫時發生錯誤！");
                    }

                }
            }
        }

        #endregion


        #region --- FFmpeg - 檢查檔案完整性 ---

        private void SongMaintenance_VerifyFile_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要開始檢查檔案完整性嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                SongMaintenance.CreateSongDataTable();
                Common_SwitchSetUI(false);

                SongMaintenance_Tooltip_Label.Text = "正在檢查檔案完整性,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMaintenance_VerifyFileTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), VolumeChangeEndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "總共檢查 " + Global.TotalList[0] + " 首歌曲的檔案完整性資料,損壞 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    SongMaintenance.DisposeSongDataTable();
                });
            }
        }

        private void SongMaintenance_VerifyFileTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (File.Exists(Global.CrazyktvDatabaseFile))
            {
                NativeMethods.SystemSleepManagement.PreventSleep(false);

                List<string> badlist = new List<string>();
                
                Parallel.ForEach(Global.SongDT.AsEnumerable(), new ParallelOptions { MaxDegreeOfParallelism = 2 }, (row, loopState) =>
                {
                    lock (LockThis) Global.TotalList[0]++;
                    this.BeginInvoke((Action)delegate ()
                    {
                        SongMaintenance_Tooltip_Label.Text = "正在解析第 " + Global.TotalList[0] + " 首歌曲的完整性資料,目前損壞 " + Global.TotalList[1] + " 首,請稍待...";
                    });

                    string file = Path.Combine(row["Song_Path"].ToString(), row["Song_FileName"].ToString());
                    bool status = MediaProbe.VerifyFile(file);
                    NativeMethods.SystemSleepManagement.ResetSleepTimer(false);

                    if (!status)
                    {
                        lock (LockThis)
                        {
                            Global.TotalList[1]++;
                            badlist.Add(row["Song_Id"].ToString());
                        }
                    }
                });

                if (badlist.Count > 0)
                {
                    foreach (string songid in badlist)
                    {
                        Global.SongLog.Add("【檔案完整性】這首歌曲的檔案可能有問題: " + songid);
                    }
                    badlist.Clear();
                    badlist = null;
                }
                NativeMethods.SystemSleepManagement.ResotreSleep();
            }
        }

        #endregion




    }




    class SongMaintenance
    {
        public static void CreateSongDataTable()
        {
            Global.SongDT = new DataTable();
            string SongQuerySqlStr = "select Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_PlayCount, Song_FileName, Song_Path, Song_ReplayGain from ktv_Song order by Song_Id";
            Global.SongDT = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");

            Global.SingerList = new List<string>();
            Global.SingerLowCaseList = new List<string>();
            Global.SingerTypeList = new List<string>();

            Global.SingerDT = new DataTable();
            string SongSingerQuerySqlStr = "select Singer_Id, Singer_Name, Singer_Type from ktv_Singer";
            Global.SingerDT = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongSingerQuerySqlStr, "");

            foreach (DataRow row in Global.SingerDT.AsEnumerable())
            {
                Global.SingerList.Add(row["Singer_Name"].ToString());
                Global.SingerLowCaseList.Add(row["Singer_Name"].ToString().ToLower());
                Global.SingerTypeList.Add(row["Singer_Type"].ToString());
            }

            Global.AllSingerList = new List<string>();
            Global.AllSingerLowCaseList = new List<string>();
            Global.AllSingerTypeList = new List<string>();

            Global.AllSingerDT = new DataTable();
            string SongAllSingerQuerySqlStr = "select Singer_Id, Singer_Name, Singer_Type from ktv_AllSinger";
            Global.AllSingerDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SongAllSingerQuerySqlStr, "");

            foreach (DataRow row in Global.AllSingerDT.AsEnumerable())
            {
                Global.AllSingerList.Add(row["Singer_Name"].ToString());
                Global.AllSingerLowCaseList.Add(row["Singer_Name"].ToString().ToLower());
                Global.AllSingerTypeList.Add(row["Singer_Type"].ToString());
            }
        }

        public static void DisposeSongDataTable()
        {
            Global.SingerList.Clear();
            Global.SingerList = null;
            Global.SingerLowCaseList.Clear();
            Global.SingerLowCaseList = null;
            Global.SingerTypeList.Clear();
            Global.SingerTypeList = null;
            Global.AllSingerList.Clear();
            Global.AllSingerList = null;
            Global.AllSingerLowCaseList.Clear();
            Global.AllSingerLowCaseList = null;
            Global.AllSingerTypeList.Clear();
            Global.AllSingerTypeList = null;
            Global.SongDT.Dispose();
            Global.SongDT = null;
            Global.SingerDT.Dispose();
            Global.SingerDT = null;
            Global.AllSingerDT.Dispose();
            Global.AllSingerDT = null;
        }

        #region --- SongMaintenance 歌曲查詢下拉清單 ---

        public static DataTable GetNewsongDateList()
        {
            using (DataTable list = new DataTable())
            {
                list.Columns.Add(new DataColumn("Display", typeof(string)));
                list.Columns.Add(new DataColumn("Value", typeof(int)));

                Global.CashboxNewSongDateDT = new DataTable();
                Global.CashboxNewSongDateDT.Columns.Add(new DataColumn("Date", typeof(string)));
                Global.CashboxNewSongDateDT.Columns.Add(new DataColumn("Value", typeof(string)));

                List<string> ItemList = new List<string>();

                string QuerySqlStr = "SELECT First(Song_CreatDate) AS Song_CreatDate, Count(Song_CreatDate) AS Song_CreatDateCount FROM ktv_Cashbox GROUP BY Song_CreatDate HAVING (Count(Song_CreatDate)>0) ORDER BY First(Song_CreatDate) desc";
                using (DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, QuerySqlStr, ""))
                {
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow row in dt.AsEnumerable())
                        {
                            string DateItem = (CultureInfo.CurrentCulture.Name == "zh-TW") ? DateTime.Parse(row["Song_CreatDate"].ToString()).ToString("yyyy年MM月份") : DateTime.Parse(row["Song_CreatDate"].ToString()).ToString("yyyy/MM");
                            string DateValue = DateTime.Parse(row["Song_CreatDate"].ToString()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
                            if (ItemList.IndexOf(DateItem) < 0)
                            {
                                ItemList.Add(DateItem);
                            }

                            var query = from drow in Global.CashboxNewSongDateDT.AsEnumerable()
                                        where drow.Field<string>("Date").Equals(DateItem)
                                        select drow;

                            if (query.Count<DataRow>() > 0)
                            {
                                foreach (DataRow drow in query)
                                {
                                    drow["Value"] = drow["Value"] + "|" + DateValue;
                                }
                            }
                            else
                            {
                                DataRow drow = Global.CashboxNewSongDateDT.NewRow();
                                drow["Date"] = DateItem;
                                drow["Value"] = DateValue;
                                Global.CashboxNewSongDateDT.Rows.Add(drow);
                            }
                        }
                    }
                }

                foreach (string str in ItemList)
                {
                    list.Rows.Add(list.NewRow());
                    list.Rows[list.Rows.Count - 1][0] = str;
                    list.Rows[list.Rows.Count - 1][1] = list.Rows.Count;
                }
                ItemList.Clear();
                return list;
            }
        }

        #endregion

        #region --- SongMaintenance 最大音量下拉清單 ---

        public static DataTable GetMaxVolumeList()
        {
            using (DataTable list = new DataTable())
            {
                list.Columns.Add(new DataColumn("Display", typeof(string)));
                list.Columns.Add(new DataColumn("Value", typeof(int)));

                List<string> ItemList = new List<string>() { "-18 dB", "-17 dB", "-16 dB", "-15 dB", "-14 dB", "-13 dB", "-12 dB", "-11 dB", "-10 dB", "-9 dB" };

                foreach (string str in ItemList)
                {
                    list.Rows.Add(list.NewRow());
                    list.Rows[list.Rows.Count - 1][0] = str;
                    list.Rows[list.Rows.Count - 1][1] = list.Rows.Count;
                }
                ItemList.Clear();
                return list;
            }
        }

        #endregion


        public static List<string> MatchCashboxData(DataTable CashboxDT, List<string> songidlist, int TotalNum)
        {
            List<string> ResultList = new List<string>();
            foreach (string songid in songidlist)
            {
                var query = from row in CashboxDT.AsEnumerable()
                            where row.Field<string>("Cashbox_Id").Equals(songid)
                            select row;

                if (query.Count<DataRow>() > 0)
                {
                    foreach (DataRow row in query)
                    {
                        int SongDataIndex = CommonFunc.MatchCashboxSong("CashboxSong", row, "", "", "", Cashbox.SongDataLowCaseList, Cashbox.SongDataFuzzyList);

                        string CashboxId = row["Cashbox_Id"].ToString();

                        if (SongDataIndex >= 0)
                        {
                            Global.TotalList[TotalNum]++;
                            ResultList.Add(Cashbox.SongIdList[SongDataIndex]);
                            break;
                        }
                    }
                }
            }
            return ResultList;
        }


        /// <summary>
        /// 錢櫃排行 / 金曲 / 新歌 lists: each non-empty list replaces its 最愛用戶, all in one transaction.
        /// Adds the users written to Global.TotalList[0]; a failure writes nothing and is logged.
        /// </summary>
        public static void ReplaceFavoriteUsers(IEnumerable<(string UserId, string Name, List<string> SongIds)> users)
        {
            var lists = users.Where(u => u.SongIds.Count > 0).ToList();
            if (lists.Count == 0) return;
            try
            {
                using (SongLibraryStore store = SongLibraryStore.OpenOrThrow())
                {
                    store.BeginTransaction();
                    try
                    {
                        foreach (var u in lists) store.ReplaceFavoriteUser(u.UserId, u.Name, u.SongIds);
                        store.Commit();
                    }
                    catch
                    {
                        store.Rollback();
                        throw;
                    }
                }
                Global.TotalList[0] += lists.Count;
            }
            catch (Exception ex)
            {
                Global.SongLog.Add("【錢櫃排行】更新我的最愛失敗,我的最愛維持原樣: " + ex.Message);
            }
        }

        private const string YoutubeFavoritePrefix = "YoutubeFavorite";

        /// <summary>
        /// A YouTube 我的最愛 as 我的最愛匯出 saves it: YoutubeFavorite|User_Id|VideoId|Length|CreatDate|Name, the 名稱 last
        /// since it may hold '|'. <paramref name="row"/> has YoutubeFavorite's VideoId, Name, Length and CreatDate.
        /// </summary>
        public static string YoutubeFavoriteLine(string userId, DataRow row)
        {
            DateTime? created = YoutubeFavorites.ReadDate(row["CreatDate"]);
            return YoutubeFavoritePrefix + "|" + userId + "|" + Convert.ToString(row["VideoId"]) + "|" +
                ((row["Length"] == DBNull.Value) ? 0 : Convert.ToInt32(row["Length"])) + "|" +
                ((created == null) ? "" : created.Value.ToString(YoutubeFavorites.DateFormat, CultureInfo.InvariantCulture)) + "|" +
                Convert.ToString(row["Name"]);
        }

        /// <summary>
        /// 我的最愛匯入: empties ktv_User and ktv_Favorite and writes back the lines 我的最愛匯出 saved, all or nothing.
        /// A ktv_Favorite line finds its song by 語系, 歌手 and 歌名 in <paramref name="songs"/> (lines whose song is
        /// gone are skipped); a repeated User_Id keeps its first line. When the backup has YouTube 我的最愛 lines they
        /// replace YoutubeFavorite too; an older backup without them keeps the YouTube 我的最愛 of the users it brings back.
        /// Returns the 最愛用戶 and songs (YouTube ones included) written; throws with nothing changed when a line cannot
        /// be read or a write fails.
        /// </summary>
        public static (int Users, int Songs) ImportFavorites(SongLibraryStore store, IEnumerable<string> lines, DataTable songs)
        {
            var users = new List<(string Id, string Name)>();
            var userIds = new HashSet<string>();
            var favorites = new Dictionary<string, List<string>>();
            var favoriteOrder = new List<string>();
            var youtube = new Dictionary<string, List<YoutubeFavoriteItem>>();
            var youtubeOrder = new List<string>();

            foreach (string line in lines)
            {
                if (line.StartsWith(YoutubeFavoritePrefix + "|", StringComparison.Ordinal))
                {
                    string[] y = line.Split(new[] { '|' }, 6);
                    if (y.Length < 6) throw new FormatException("YouTube 我的最愛資料格式錯誤: " + line);
                    if (!youtube.TryGetValue(y[1], out List<YoutubeFavoriteItem> items))
                    {
                        items = new List<YoutubeFavoriteItem>();
                        youtube[y[1]] = items;
                        youtubeOrder.Add(y[1]);
                    }
                    int.TryParse(y[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out int length);
                    items.Add(new YoutubeFavoriteItem(y[2], y[5], length, YoutubeFavorites.ReadDate(y[4])));
                    continue;
                }

                string[] f = line.Split('|');
                switch (f[0])
                {
                    case "ktv_User":
                        if (userIds.Add(f[1])) users.Add((f[1], f[2]));
                        break;
                    case "ktv_Favorite":
                        DataRow song = songs.AsEnumerable().FirstOrDefault(r =>
                            r.Field<string>("Song_Lang") == f[2] &&
                            r.Field<string>("Song_Singer") == f[3] &&
                            r.Field<string>("Song_SongName") == f[4]);
                        if (song == null) break;
                        if (!favorites.TryGetValue(f[1], out List<string> ids))
                        {
                            ids = new List<string>();
                            favorites[f[1]] = ids;
                            favoriteOrder.Add(f[1]);
                        }
                        ids.Add(song["Song_Id"].ToString());
                        break;
                }
            }

            int songCount = 0;
            store.BeginTransaction();
            try
            {
                store.ClearFavoriteUsers();
                foreach (var u in users) store.AddFavoriteUser(u.Id, u.Name);
                foreach (string userId in favoriteOrder) songCount += store.AddFavorites(userId, favorites[userId]);
                if (youtubeOrder.Count > 0)
                {
                    store.ClearYoutubeFavorites();
                    foreach (string userId in youtubeOrder) songCount += store.AddYoutubeFavorites(userId, youtube[userId]);
                }
                else
                {
                    store.RemoveYoutubeFavoritesExcept(userIds);
                }
                store.Commit();
            }
            catch
            {
                store.Rollback();
                throw;
            }
            return (users.Count, songCount);
        }


    }
}

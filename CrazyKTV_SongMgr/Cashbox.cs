using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
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
        #region --- Cashbox 控制項事件 ---

        private void Cashbox_QueryType_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (Cashbox_QueryType_ComboBox.SelectedValue.ToString())
            {
                case "1":
                    Cashbox_QueryValue_TextBox.ImeMode = ImeMode.OnHalf;
                    Cashbox_QueryValue_TextBox.Text = "";
                    Cashbox_QueryValue_TextBox.Enabled = true;
                    Cashbox_QueryValue_ComboBox.Visible = false;
                    Cashbox_QueryValue_TextBox.Visible = true;
                    Cashbox_Paste_Button.Enabled = true;
                    Cashbox_Clear_Button.Enabled = true;
                    Cashbox_QueryValue_TextBox.Focus();
                    break;
                case "2":
                    Cashbox_QueryValue_TextBox.ImeMode = ImeMode.OnHalf;
                    Cashbox_QueryValue_TextBox.Text = "";
                    Cashbox_QueryValue_TextBox.Enabled = true;
                    Cashbox_QueryValue_ComboBox.Visible = false;
                    Cashbox_QueryValue_TextBox.Visible = true;
                    Cashbox_Paste_Button.Enabled = true;
                    Cashbox_Clear_Button.Enabled = true;
                    Cashbox_QueryValue_TextBox.Focus();
                    break;
                case "3":
                    Cashbox_QueryValue_TextBox.ImeMode = ImeMode.Close;
                    Cashbox_QueryValue_TextBox.Text = "";
                    Cashbox_QueryValue_TextBox.Enabled = true;
                    Cashbox_QueryValue_ComboBox.Visible = false;
                    Cashbox_QueryValue_TextBox.Visible = true;
                    Cashbox_Paste_Button.Enabled = false;
                    Cashbox_Clear_Button.Enabled = false;
                    Cashbox_QueryValue_TextBox.Focus();
                    break;
                case "4":
                    Cashbox_QueryValue_TextBox.ImeMode = ImeMode.Close;
                    Cashbox_QueryValue_TextBox.Text = "*";
                    Cashbox_QueryValue_TextBox.Enabled = false;
                    Cashbox_QueryValue_ComboBox.Visible = false;
                    Cashbox_QueryValue_TextBox.Visible = true;
                    Cashbox_Paste_Button.Enabled = false;
                    Cashbox_Clear_Button.Enabled = false;
                    Cashbox_Query_Button.PerformClick();
                    break;
                case "5":
                    Cashbox_QueryValue_TextBox.ImeMode = ImeMode.Close;
                    Cashbox_QueryValue_TextBox.Text = "100";
                    Cashbox_QueryValue_TextBox.Enabled = true;
                    Cashbox_QueryValue_ComboBox.Visible = false;
                    Cashbox_QueryValue_TextBox.Visible = true;
                    Cashbox_Paste_Button.Enabled = false;
                    Cashbox_Clear_Button.Enabled = false;
                    Cashbox_Query_Button.PerformClick();
                    break;
            }
        }

        private void Cashbox_QueryFilter_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Cashbox_QueryFilter_ComboBox.SelectedValue.ToString() != "System.Data.DataRowView")
            {
                Global.CashboxQueryFilter = Cashbox_QueryFilter_ComboBox.Text;
            }
        }

        private void Cashbox_QueryValue_TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (Cashbox_QueryType_ComboBox.SelectedValue.ToString())
            {
                case "3":
                case "5":
                    if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8 & (int)e.KeyChar != 13)
                    {
                        e.Handled = true;
                    }
                    break;
            }
            if ((int)e.KeyChar == 13)
            {
                Cashbox_Query_Button.PerformClick();
            }
        }

        private void Cashbox_Paste_Button_Click(object sender, EventArgs e)
        {
            Cashbox_QueryValue_TextBox.Text = Clipboard.GetText();
        }

        private void Cashbox_Clear_Button_Click(object sender, EventArgs e)
        {
            Cashbox_QueryValue_TextBox.Text = "";
        }

        private void Cashbox_SynonymousQuery_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Global.CashboxSynonymousQuery = Cashbox_SynonymousQuery_CheckBox.Checked;
        }

        private void Cashbox_FuzzyQuery_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Global.CashboxFuzzyQuery = Cashbox_FuzzyQuery_CheckBox.Checked;
        }

        private void Cashbox_CommonFilter_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (((ComboBox)sender).Focused)
            {
                if (((ComboBox)sender).Name == "Cashbox_LangFilter_ComboBox")
                {
                    List<string> SongSingerFilterList = new List<string>();
                    ((DataTable)Cashbox_DataGridView.DataSource).DefaultView.RowFilter = "";

                    using (DataTable dt = (DataTable)Cashbox_DataGridView.DataSource)
                    {
                        using (DataTable SongSingerDistinctDT = dt.DefaultView.ToTable(true, "Song_Lang", "Song_Singer"))
                        {
                            if (SongSingerDistinctDT.Rows.Count > 0)
                            {
                                string SongSinger = string.Empty;
                                foreach (DataRow row in SongSingerDistinctDT.AsEnumerable())
                                {
                                    SongSinger = row["Song_Singer"].ToString();
                                    if (SongSinger.Contains("&"))
                                    {
                                        string ChorusSongSingerName = SongSinger;
                                        List<string> list = CommonFunc.GetChorusSingerList(ChorusSongSingerName);
                                        foreach (string str in list)
                                        {
                                            string SingerStr = Regex.Replace(str, @"^\s*|\s*$", ""); //去除頭尾空白
                                            if (SongSingerFilterList.IndexOf(SingerStr) < 0)
                                            {
                                                if (row["Song_Lang"].ToString() == Cashbox_LangFilter_ComboBox.Text || Cashbox_LangFilter_ComboBox.Text == "全部")
                                                {
                                                    SongSingerFilterList.Add(SingerStr);
                                                }
                                            }
                                        }
                                        list.Clear();
                                        list = null;
                                    }
                                    else
                                    {
                                        if (SongSingerFilterList.IndexOf(SongSinger) < 0)
                                        {
                                            if (row["Song_Lang"].ToString() == Cashbox_LangFilter_ComboBox.Text || Cashbox_LangFilter_ComboBox.Text == "全部")
                                            {
                                                SongSingerFilterList.Add(SongSinger);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (SongSingerFilterList.Count > 0)
                    {
                        SongSingerFilterList.Sort();
                        Cashbox_SongSingerFilter_ComboBox.Enabled = true;
                        Cashbox_SongSingerFilter_ComboBox.DataSource = Cashbox.GetQueryFilterList(SongSingerFilterList);
                        Cashbox_SongSingerFilter_ComboBox.DisplayMember = "Display";
                        Cashbox_SongSingerFilter_ComboBox.ValueMember = "Value";
                        SongSingerFilterList.Clear();
                    }
                }

                string FilterStr = "";
                if (Cashbox_LangFilter_ComboBox.Text != "全部" && Cashbox_LangFilter_ComboBox.Text != "") FilterStr = "Song_Lang = '" + Cashbox_LangFilter_ComboBox.Text + "'";
                if (Cashbox_SongSingerFilter_ComboBox.Text != "全部" && Cashbox_SongSingerFilter_ComboBox.Text != "")
                {
                    string SongSinger = Cashbox_SongSingerFilter_ComboBox.Text;
                    SongSinger = Regex.Replace(SongSinger, @"(%)|(_)|(\[)|(])|(\^)", "[$1$2$3$4$5]");

                    FilterStr += (FilterStr != "") ? " and (" : "";
                    FilterStr += "Song_Singer = '" + SongSinger + "' or Song_Singer like '%&" + SongSinger + "%' or Song_Singer like '%" + SongSinger + "&%'" + ((FilterStr != "") ? ")" : "");
                }
                ((DataTable)Cashbox_DataGridView.DataSource).DefaultView.RowFilter = FilterStr;
            }
        }

        #endregion

        #region --- Cashbox 查詢歌曲 ---

        private void Cashbox_Query_Button_Click(object sender, EventArgs e)
        {
            if (Global.CrazyktvDatabaseStatus)
            {
                string QueryText = Cashbox_QueryValue_TextBox.Text;
                CashboxSearchKind Kind;
                string SongQueryStatusText;

                switch (Cashbox_QueryType_ComboBox.SelectedValue.ToString())
                {
                    case "1":
                        Kind = CashboxSearchKind.SongName;
                        SongQueryStatusText = QueryText;
                        Cashbox_QueryStatus_Label.Text = "正在查詢歌曲名稱為『" + SongQueryStatusText + "』的相關歌曲,請稍待...";
                        break;
                    case "2":
                        Kind = CashboxSearchKind.SingerName;
                        SongQueryStatusText = QueryText;
                        Cashbox_QueryStatus_Label.Text = "正在查詢歌手名稱為『" + SongQueryStatusText + "』的相關歌曲,請稍待...";
                        break;
                    case "3":
                        Kind = CashboxSearchKind.CashboxId;
                        SongQueryStatusText = "歌曲編號中包含 " + QueryText;
                        Cashbox_QueryStatus_Label.Text = "正在查詢『" + SongQueryStatusText + "』的相關歌曲,請稍待...";
                        break;
                    case "4":
                        Kind = CashboxSearchKind.AllSongs;
                        SongQueryStatusText = "全部歌曲";
                        Cashbox_QueryStatus_Label.Text = "正在查詢" + SongQueryStatusText + ",請稍待...";
                        break;
                    case "5":
                        Kind = CashboxSearchKind.NewSongs;
                        SongQueryStatusText = "新進歌曲";
                        Cashbox_QueryStatus_Label.Text = "正在查詢" + SongQueryStatusText + ",請稍待...";
                        break;
                    default:
                        return;
                }

                // 必須輸入查詢條件: the kinds that read the text box need some text.
                if (Kind != CashboxSearchKind.AllSongs && QueryText == "")
                {
                    Cashbox_DataGridView.DataSource = null;
                    Cashbox_ShowNoResult("必須輸入查詢條件才能查詢...");
                    return;
                }

                var Criteria = new CashboxSearchCriteria(Kind, QueryText, Global.CashboxQueryFilter, Global.CashboxFuzzyQuery, Global.CashboxSynonymousQuery);
                Cashbox_RunSearch(Criteria,
                    "查無『" + SongQueryStatusText + "』的相關歌曲,請重新查詢...",
                    Count => "總共查詢到 " + Count + " 筆有關『" + SongQueryStatusText + "』的歌曲。");
            }
        }

        /// <summary>Runs a 錢櫃查詢 in the background with the UI locked, then shows the songs found.</summary>
        private void Cashbox_RunSearch(CashboxSearchCriteria criteria, string noResultText, Func<int, string> resultText)
        {
            Cashbox_Query_Button.Enabled = false;
            Common_SwitchSetUI(false);

            Cashbox_DataGridView.DataSource = null;
            GC.Collect();

            Task.Factory.StartNew(() =>
            {
                Thread.CurrentThread.Priority = ThreadPriority.Lowest;
                DataTable dt = null;
                try
                {
                    SongSearchReference reference = CashboxSearch.CaptureReference(criteria.Kind);
                    dt = CashboxSearch.Run(criteria, reference, Global.CrazyktvSongMgrDatabaseFile).Songs;
                }
                catch
                {
                    dt = null;
                }

                this.BeginInvoke((Action)delegate()
                {
                    if (dt == null) Cashbox_ShowNoResult("查詢條件輸入錯誤,請重新輸入...");
                    else if (dt.Rows.Count == 0) Cashbox_ShowNoResult(noResultText);
                    else Cashbox_ShowResult(dt, resultText(dt.Rows.Count));

                    Common_SwitchSetUI(true);
                    Cashbox_Query_Button.Enabled = true;
                    if (Cashbox_DataGridView.Rows.Count > 0) Cashbox_DataGridView.Focus();
                });
            });
        }

        private void Cashbox_ShowNoResult(string statusText)
        {
            Cashbox_EditMode_CheckBox.Enabled = false;
            Cashbox_LangFilter_ComboBox.Enabled = false;
            Cashbox_SongSingerFilter_ComboBox.Enabled = false;
            Cashbox_QueryStatus_Label.Text = statusText;
        }

        /// <summary>Binds the 錢櫃 songs found to the 錢櫃 list, with the 語系 and 歌手 filters.</summary>
        private void Cashbox_ShowResult(DataTable dt, string statusText)
        {
            Cashbox_EditMode_CheckBox.Enabled = true;
            Cashbox_QueryStatus_Label.Text = statusText;

            Cashbox_DataGridView.DataSource = dt;

            List<string> LangFilterList = dt.AsEnumerable().Select(row => row["Song_Lang"].ToString()).Distinct().ToList();
            List<string> SongSingerFilterList = new List<string>();

            foreach (string SongSinger in dt.AsEnumerable().Select(row => row["Song_Singer"].ToString()).Distinct())
            {
                IEnumerable<string> singers = SongSinger.Contains("&")
                    ? CommonFunc.GetChorusSingerList(SongSinger).Select(s => Regex.Replace(s, @"^\s*|\s*$", ""))
                    : new[] { SongSinger };
                foreach (string singer in singers)
                {
                    if (SongSingerFilterList.IndexOf(singer) < 0) SongSingerFilterList.Add(singer);
                }
            }

            if (LangFilterList.Count > 0)
            {
                Cashbox_LangFilter_ComboBox.Enabled = true;
                Cashbox_LangFilter_ComboBox.DataSource = Cashbox.GetQueryFilterList(LangFilterList);
                Cashbox_LangFilter_ComboBox.DisplayMember = "Display";
                Cashbox_LangFilter_ComboBox.ValueMember = "Value";
            }

            if (SongSingerFilterList.Count > 0)
            {
                SongSingerFilterList.Sort();
                Cashbox_SongSingerFilter_ComboBox.Enabled = true;
                Cashbox_SongSingerFilter_ComboBox.DataSource = Cashbox.GetQueryFilterList(SongSingerFilterList);
                Cashbox_SongSingerFilter_ComboBox.DisplayMember = "Display";
                Cashbox_SongSingerFilter_ComboBox.ValueMember = "Value";
            }

            for (int i = 0; i < Cashbox_DataGridView.ColumnCount; i++)
            {
                List<string> DataGridViewColumnName = Cashbox.GetDataGridViewColumnSet(Cashbox_DataGridView.Columns[i].Name);
                Cashbox_DataGridView.Columns[i].HeaderText = DataGridViewColumnName[0];

                if (DataGridViewColumnName[1].ToString() == "0")
                {
                    Cashbox_DataGridView.Columns[i].Visible = false;
                }

                if (DataGridViewColumnName[2].ToString() != "none")
                {
                    ((DataGridViewTextBoxColumn)Cashbox_DataGridView.Columns[i]).MaxInputLength = int.Parse(DataGridViewColumnName[2]);
                }

                Cashbox_DataGridView.Columns[i].Width = int.Parse(DataGridViewColumnName[1]);
            }

            int ColumnWidth_140 = Convert.ToInt32(140 * Global.DPIScalingFactor * MainUIScale.UIScalingFactor);
            Cashbox_DataGridView.Columns["Song_CreatDate"].MinimumWidth = ColumnWidth_140;
            Cashbox_DataGridView.Columns["Song_CreatDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        #endregion

        #region --- Cashbox 其它查詢 ---

        private void Cashbox_OtherQuery_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Global.CrazyktvDatabaseStatus && Cashbox_OtherQuery_ComboBox.SelectedValue.ToString() != "System.Data.DataRowView")
            {
                string SongQueryStatusText = Cashbox_OtherQuery_ComboBox.Text;

                switch (Cashbox_OtherQuery_ComboBox.SelectedValue.ToString())
                {
                    case "1":
                        Cashbox_QueryStatus_Label.Text = "正在查詢" + SongQueryStatusText + ",請稍待...";
                        Cashbox_RunSearch(new CashboxSearchCriteria(CashboxSearchKind.Missing),
                            "查無『" + SongQueryStatusText + "』的相關歌曲,請重新查詢...",
                            Count => "總共查詢到 " + Count + " 筆歌庫所沒有的錢櫃歌曲。");
                        break;
                    case "2":
                        Cashbox_QueryStatus_Label.Text = "正在查詢" + SongQueryStatusText + ",請稍待...";
                        Cashbox_RunSearch(new CashboxSearchCriteria(CashboxSearchKind.Duplicate),
                            "查無『" + SongQueryStatusText + "』的相關歌曲,請重新查詢...",
                            Count => "總共查詢到 " + Count + " 筆有關『" + SongQueryStatusText + "』的歌曲。");
                        break;
                }
            }
        }

        private void Cashbox_DateQuery_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Global.CrazyktvDatabaseStatus && Cashbox_DateQuery_ComboBox.SelectedValue.ToString() != "System.Data.DataRowView")
            {
                string SongQueryStatusText = Cashbox_DateQuery_ComboBox.Text;
                if (!DateTime.TryParseExact(SongQueryStatusText, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime Date)) return;

                Cashbox_QueryStatus_Label.Text = "正在查詢更新日期為『" + SongQueryStatusText + "』的歌曲,請稍待...";
                Cashbox_RunSearch(new CashboxSearchCriteria(CashboxSearchKind.UpdatedOn, date: Date),
                    "查無『" + SongQueryStatusText + "』的相關歌曲,請重新查詢...",
                    Count => "總共查詢到 " + Count + " 筆更新日期為『" + SongQueryStatusText + "』的歌曲。");
            }
        }

        #endregion

        #region --- Cashbox 更新歌曲資料 ---

        private void Cashbox_UpdDate_Button_Click(object sender, EventArgs e)
        {
            if (Global.CrazyktvDatabaseStatus)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Common_SwitchSetUI(false);

                Cashbox_QueryStatus_Label.Text = "正在更新錢櫃資料,請稍待...";

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => Cashbox_UpdDateTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    this.BeginInvoke((Action)delegate()
                    {
                        if (Cashbox_QueryStatus_Label.Text != "錢櫃資料已是最新!")
                            Common_InitializeSongData(false, false, true, false, false, false);
                        Common_SwitchSetUI(true);
                        Global.TimerEndTime = DateTime.Now;
                        if (Cashbox_QueryStatus_Label.Text != "錢櫃資料已是最新!")
                            Cashbox_QueryStatus_Label.Text = "總共更新 " + Global.TotalList[0] + " 筆資料,失敗 " + Global.TotalList[1] + " 筆,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                    });
                });
            }
        }

        private void Cashbox_UpdDateTask()
        {
            string SongDate = string.Empty;
            List<string> SongDataList = new List<string>();
            List<string> SongIdList = new List<string>();
            bool GetDataFromCashboxWeb = false;
            bool GetDataFromGithubWeb = false;

            DateTime DatePrevUpdDate = DateTime.Parse(((DataTable)Cashbox_DateQuery_ComboBox.DataSource).Rows[0][0].ToString());
            
            if ((DateTime.Now - DatePrevUpdDate).Days < 7)
            {
                this.BeginInvoke((Action)delegate ()
                {
                    Cashbox_QueryStatus_Label.Text = "錢櫃資料已是最新!";
                });
                return;
            }
            else
            {
                if (CommonFunc.CheckFor45PlusVersion())
                {
                    GetDataFromCashboxWeb = true;
                    this.BeginInvoke((Action)delegate ()
                    {
                        Cashbox_QueryStatus_Label.Text = "正在從錢櫃網站獲取資料...";
                    });
                }
                else
                {
                    GetDataFromGithubWeb = true;
                    this.BeginInvoke((Action)delegate ()
                    {
                        Cashbox_QueryStatus_Label.Text = "正在從 Github 網站獲取資料...";
                    });
                }
            }

            string CashboxQuerySqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox order by Cashbox_Id, Song_CreatDate";
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, CashboxQuerySqlStr, ""))
            {
                Parallel.ForEach(dt.AsEnumerable(), (row, loopState) =>
                {
                    string SongCreatDate = DateTime.Parse(row["Song_CreatDate"].ToString()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
                    lock (LockThis)
                    {
                        SongIdList.Add(row["Cashbox_Id"].ToString());
                    }
                });
            }

            if (GetDataFromCashboxWeb)
            {
                DateTime DateValidDate = DateTime.Parse("2016/01/01");
                DateTime DateEndDate = DateTime.Now;

                if (DateTime.Compare(DateEndDate, DatePrevUpdDate) < 0)
                {
                    this.BeginInvoke((Action)delegate ()
                    {
                        Cashbox_QueryStatus_Label.Text = "系統時間錯誤!";
                    });
                    return;
                }

                if (DateTime.Compare(DatePrevUpdDate, DateValidDate) < 0 || DateTime.Compare(DateEndDate, DateValidDate) < 0)
                {
                    this.BeginInvoke((Action)delegate ()
                    {
                        Cashbox_QueryStatus_Label.Text = "僅支援更新2016年後的錢櫃歌曲!";
                    });
                    return;
                }

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                HtmlNode table;
                HtmlNodeCollection child;

                List<string> sDateList = new List<string>();

                int DaysCount = (DateEndDate - DatePrevUpdDate).Days;

                for (int i = 1; i <= DaysCount; i++)
                {
                    if (sDateList.IndexOf(DatePrevUpdDate.AddDays(i).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)) < 0) sDateList.Add(DatePrevUpdDate.AddDays(i).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture));
                }
                if (sDateList.IndexOf(DateEndDate.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)) < 0) sDateList.Add(DateEndDate.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture));

                foreach (string sdate in sDateList)
                {
                    string url = "https://www.cashboxparty.com/ashx/ah_Song.ashx";
                    string poststr = string.Format("[{1}\"m\":\"sn\",\"lang\":\"01,02\",\"date\":\"{0}\"{2}]", sdate, "{", "}");
                    string jsonString = CommonFunc.DownloadCashboxData(url, poststr);

                    using (DataTable dt = CommonFunc.JSONtoDataTable(jsonString))
                    {
                        if (dt.Rows.Count > 0 && dt.Columns.Contains("SongName"))
                        {
                            this.BeginInvoke((Action)delegate ()
                            {
                                Cashbox_QueryStatus_Label.Text = "正在分析第 " + (sDateList.IndexOf(sdate) + 1) + " / " + sDateList.Count + " 天的更新歌曲,請稍待...";
                            });

                            foreach (DataRow row in dt.Rows)
                            {
                                List<string> list = new List<string>()
                                 {
                                     row["SongNo"].ToString(),
                                     row["LangName"].ToString(),
                                     row["SongName"].ToString(),
                                     row["Singer"].ToString(),
                                     row["Indate"].ToString()
                                 };

                                if (list[1] == "") list[1] = "其它";
                                if (CommonFunc.IsSongId(list[0]) && list[1] != "" && list[2] != "" && list[3] != "")
                                {
                                    list[3] = Regex.Replace(list[3], "、", "&");

                                    if (SongIdList.IndexOf(list[0]) < 0)
                                    {
                                        SongIdList.Add(list[0]);
                                        list.Add("AddSong");
                                    }
                                    else
                                    {
                                        list.Add("UpdSong");
                                    }
                                    SongDataList.Add(string.Join("|", list));
                                }
                                list.Clear();
                                list = null;
                            }
                        }
                    }
                }
                SongIdList = null;
            }

            if (GetDataFromGithubWeb)
            {
                string url = "https://raw.githubusercontent.com/CrazyKTV/WebUpdater/master/CrazyKTV_WebUpdater/Cashbox/cashbox_newsong.json";
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
                            using (DataTable dt = CommonFunc.JSONtoDataTable(jsonString))
                            {
                                if (dt.Rows.Count > 0)
                                {
                                    var query = from row in dt.AsEnumerable()
                                                where row.Field<DateTime>("Song_CreatDate") >= DatePrevUpdDate
                                                select row;

                                    if (query.Count<DataRow>() > 0)
                                    {
                                        foreach (DataRow row in query)
                                        {
                                            List<string> list = new List<string>()
                                            {
                                                row["Cashbox_Id"].ToString(),
                                                row["Song_Lang"].ToString(),
                                                row["Song_SongName"].ToString(),
                                                row["Song_Singer"].ToString(),
                                                DateTime.Parse(row["Song_CreatDate"].ToString()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                                            };

                                            if (list[1] == "") list[1] = "其它";
                                            if (CommonFunc.IsSongId(list[0]) && list[1] != "" && list[2] != "" && list[3] != "" && list[4] != "")
                                            {
                                                list[3] = Regex.Replace(list[3], "、", "&");

                                                if (SongIdList.IndexOf(list[0]) < 0)
                                                {
                                                    SongIdList.Add(list[0]);
                                                    list.Add("AddSong");
                                                }
                                                else
                                                {
                                                    list.Add("UpdSong");
                                                }
                                                SongDataList.Add(string.Join("|", list));
                                            }
                                            list.Clear();
                                            list = null;
                                        }
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
                SongIdList = null;
            }

            if (SongDataList.Count > 0)
            {
                using (DbConnection conn = Db.OpenConn(Global.CrazyktvSongMgrDatabaseFile, ""))
                {
                    string sqlAddStr = "Cashbox_Id, Song_Lang, Song_SongName, Song_Singer, Song_CreatDate";
                    string sqlValuesStr = "@CashboxId, @SongLang, @SongSongName, @SongSinger, @SongCreatDate";
                    string AddSqlStr = "insert into ktv_Cashbox ( " + sqlAddStr + " ) values ( " + sqlValuesStr + " )";
                    string sqlUpdStr = "Cashbox_Id = @CashboxId, Song_Lang = @SongLang, Song_SongName = @SongSongName, Song_Singer = @SongSinger, Song_CreatDate = @SongCreatDate";
                    string UpdSqlStr = "update ktv_Cashbox set " + sqlUpdStr + " where Cashbox_Id = @OldCashboxId";

                    foreach (string SongData in SongDataList)
                    {
                        List<string> valuelist = new List<string>(SongData.Split('|'));

                        switch (valuelist[5])
                        {
                            case "AddSong":
                                using (DbCommand AddCmd = Db.Command(AddSqlStr, conn))
                                {
                                    AddCmd.Parameters.AddWithValue("@CashboxId", valuelist[0]);
                                    AddCmd.Parameters.AddWithValue("@SongLang", valuelist[1]);
                                    AddCmd.Parameters.AddWithValue("@SongSongName", valuelist[2]);
                                    AddCmd.Parameters.AddWithValue("@SongSinger", valuelist[3]);
                                    AddCmd.Parameters.AddWithValue("@SongCreatDate", Db.DateParam(conn, valuelist[4]));

                                    try
                                    {
                                        AddCmd.ExecuteNonQuery();
                                        Global.TotalList[0]++;
                                        this.BeginInvoke((Action)delegate ()
                                        {
                                            Cashbox_QueryStatus_Label.Text = "正在將第 " + Global.TotalList[0] + " 首歌曲寫入資料庫,請稍待...";
                                        });
                                    }
                                    catch
                                    {
                                        Global.TotalList[1]++;
                                        Global.SongLog.Add("加入錢櫃資料時發生未知的錯誤: " + SongData);
                                    }
                                }
                                break;
                            case "UpdSong":
                                using (DbCommand UpdCmd = Db.Command(UpdSqlStr, conn))
                                {
                                    UpdCmd.Parameters.AddWithValue("@CashboxId", valuelist[0]);
                                    UpdCmd.Parameters.AddWithValue("@SongLang", valuelist[1]);
                                    UpdCmd.Parameters.AddWithValue("@SongSongName", valuelist[2]);
                                    UpdCmd.Parameters.AddWithValue("@SongSinger", valuelist[3]);
                                    UpdCmd.Parameters.AddWithValue("@SongCreatDate", Db.DateParam(conn, valuelist[4]));
                                    UpdCmd.Parameters.AddWithValue("@OldCashboxId", valuelist[0]);

                                    try
                                    {
                                        UpdCmd.ExecuteNonQuery();
                                        Global.TotalList[0]++;
                                        this.BeginInvoke((Action)delegate ()
                                        {
                                            Cashbox_QueryStatus_Label.Text = "正在將第 " + Global.TotalList[0] + " 首歌曲寫入資料庫,請稍待...";
                                        });
                                    }
                                    catch
                                    {
                                        Global.TotalList[1]++;
                                        Global.SongLog.Add("更新錢櫃資料時發生未知的錯誤: " + SongData);
                                    }
                                }
                                break;
                        }
                        valuelist.Clear();
                        valuelist = null;
                    }
                }
                SongDataList.Clear();
                SongDataList = null;

                using (DbConnection conn = Db.OpenConn(Global.CrazyktvSongMgrDatabaseFile, ""))
                {
                    Global.CashboxUpdDate = DateTime.Now;
                    string CashboxUpdDateSqlStr = "CashboxUpdDate = @CashboxUpdDate";
                    string CashboxUpdDateUpdateSqlStr = "update ktv_Version set " + CashboxUpdDateSqlStr + " where Id = @Id";

                    using (DbCommand Versioncmd = Db.Command(CashboxUpdDateUpdateSqlStr, conn))
                    {
                        Versioncmd.Parameters.AddWithValue("@CashboxUpdDate", Db.DateParam(conn, Global.CashboxUpdDate.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)));
                        Versioncmd.Parameters.AddWithValue("@Id", "1");
                        Versioncmd.ExecuteNonQuery();
                    }
                }

                this.BeginInvoke((Action)delegate ()
                {
                    Cashbox_UpdDateValue_Label.Text = (CultureInfo.CurrentCulture.Name == "zh-TW") ? Global.CashboxUpdDate.ToLongDateString() : Global.CashboxUpdDate.ToShortDateString();

                    Cashbox_DateQuery_ComboBox.SelectedIndexChanged -= new EventHandler(Cashbox_DateQuery_ComboBox_SelectedIndexChanged);
                    Cashbox_DateQuery_ComboBox.DataSource = Cashbox.GetDateQueryList();
                    Cashbox_DateQuery_ComboBox.DisplayMember = "Display";
                    Cashbox_DateQuery_ComboBox.ValueMember = "Value";
                    Cashbox_DateQuery_ComboBox.SelectedValue = 1;
                    Cashbox_DateQuery_ComboBox.SelectedIndexChanged += new EventHandler(Cashbox_DateQuery_ComboBox_SelectedIndexChanged);

                    SongMaintenance_Favorite_UpdateNewsong_ComboBox.DataSource = SongMaintenance.GetNewsongDateList();
                    SongMaintenance_Favorite_UpdateNewsong_ComboBox.DisplayMember = "Display";
                    SongMaintenance_Favorite_UpdateNewsong_ComboBox.ValueMember = "Value";
                    SongMaintenance_Favorite_UpdateNewsong_ComboBox.SelectedValue = 1;
                });
            }
        }

        #endregion

        #region --- Cashbox 套用錢櫃編號 ---

        private void Cashbox_ApplyCashboxId_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要套用錢櫃編號嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox_QueryStatus_Label.Text = "正在解析錢櫃歌曲編號,請稍待...";

                Common_SwitchSetUI(false);
                Cashbox.CreateSongDataTable();

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => Cashbox_ApplyCashboxIdTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        Cashbox_QueryStatus_Label.Text = "總共變更 " + Global.TotalList[0] + " 首歌曲為錢櫃編號,失敗 " + Global.TotalList[1] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_SwitchSetUI(true);
                    });
                    Cashbox.DisposeSongDataTable();
                });
            }
        }

        private void Cashbox_ApplyCashboxIdTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                // The 我的最愛 move with each song (SongLibraryStore.RenumberSong), so nothing is exported and re-imported.
                var library = new List<CashboxIdPlan.Song>();
                for (int i = 0; i < Cashbox.SongIdList.Count; i++)
                {
                    string[] data = Cashbox.SongDataList[i].Split('|');
                    library.Add(new CashboxIdPlan.Song(Cashbox.SongIdList[i], data[0], data[1], data[2]));
                }
                var matches = new List<(CashboxIdPlan.Song Song, string CashboxId)>();

                string MaxDigitCode = (Global.SongMgrMaxDigitCode == "1") ? "D5" : "D6";

                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    Parallel.ForEach(Global.CashboxSongLangList, (langstr, loopState) =>
                    {
                        var query = from row in CashboxDT.AsEnumerable()
                                    where row.Field<string>("Song_Lang").Equals(langstr)
                                    select row;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow row in query)
                            {
                                string CashboxId = Convert.ToInt32(row["Cashbox_Id"].ToString()).ToString(MaxDigitCode);
                                int SongDataIndex = CommonFunc.MatchCashboxSong("CashboxSong", row, "", "", "", Cashbox.SongDataLowCaseList, Cashbox.SongDataFuzzyList);

                                if (SongDataIndex >= 0)
                                {
                                    lock (LockThis) { Global.TotalList[2]++; }

                                    if (CashboxId != Cashbox.SongIdList[SongDataIndex])
                                    {
                                        lock (LockThis) { matches.Add((library[SongDataIndex], CashboxId)); }
                                    }

                                    this.BeginInvoke((Action)delegate()
                                    {
                                        Cashbox_QueryStatus_Label.Text = "已在歌庫找到 " + Global.TotalList[2] + " 首錢櫃歌曲...";
                                    });
                                }
                            }
                        }
                    });
                }
                CashboxIdPlan.Result plan = CashboxIdPlan.Plan(library, matches, SongIdAllocator.FromSettings);
                foreach (var skipped in plan.Skipped)
                {
                    Global.SongLog.Add("略過套用編號: " + skipped.Song.Id + " => " + skipped.CashboxId + " (" + skipped.Song + "),此錢櫃編號已分配給其它歌曲或此歌曲已有錢櫃編號");
                }

                // Each song once, from where it was to where it ends up (a swap stops on a spare number in between).
                var changes = plan.Moves.GroupBy(m => m.Song)
                    .Select(g => (Song: g.Key, From: g.First().From, To: g.Last().To, Cashbox: g.Any(m => m.Why == CashboxIdPlan.Reason.Cashbox)))
                    .ToList();

                using (SongLibraryStore store = SongLibraryStore.Open())
                {
                    if (CashboxIdPlan.Apply(store, plan.Moves, out Exception error))
                    {
                        foreach (var c in changes)
                        {
                            if (c.Cashbox)
                            {
                                Global.TotalList[0]++;
                                Global.SongLog.Add("已成功套用編號: " + c.From + " => " + c.To + " (" + c.Song + ")");
                            }
                            else
                            {
                                Global.SongLog.Add("已成功讓出編號: " + c.From + " => " + c.To + " (" + c.Song + ")");
                            }
                        }
                    }
                    else
                    {
                        Global.TotalList[1] += changes.Count(c => c.Cashbox);
                        Global.SongLog.Add("套用錢櫃編號時發生錯誤,已還原所有變更: " + error.Message);
                    }
                }
            }
        }

        #endregion

        #region --- Cashbox 記錄已有歌曲 ---

        private void Cashbox_WriteHaveSongData_Button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你確定要記錄已有歌曲嗎?", "確認提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Global.TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0 };
                Cashbox_QueryStatus_Label.Text = "正在解析已有歌曲,請稍待...";

                Common_SwitchSetUI(false);
                Cashbox.CreateSongDataTable();

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => Cashbox_WriteHaveSongDataTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    Global.TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        Cashbox_QueryStatus_Label.Text = "總共解析到 " + Global.TotalList[0] + " 首已有歌曲,更新 " + Global.TotalList[1] + " 首,失敗 " + Global.TotalList[2] + " 首,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                        Common_InitializeSongData(false, false, false, false, false, true);
                        Common_SwitchSetUI(true);
                    });
                    Cashbox.DisposeSongDataTable();
                });
            }
        }

        private void Cashbox_WriteHaveSongDataTask()
        {
            Thread.CurrentThread.Priority = ThreadPriority.Lowest;

            if (Global.CrazyktvDatabaseStatus)
            {
                List<string> UpdateList = new List<string>();
                string SqlStr = "select Cashbox_Id, Song_Lang, Song_Singer, Song_SongName, Song_CreatDate, Song_HaveSong from ktv_Cashbox";
                using (DataTable CashboxDT = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, SqlStr, ""))
                {
                    Parallel.ForEach(Global.CashboxSongLangList, (langstr, loopState) =>
                    {
                        var query = from row in CashboxDT.AsEnumerable()
                                    where row.Field<string>("Song_Lang").Equals(langstr)
                                    select row;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow row in query)
                            {
                                int SongDataIndex = CommonFunc.MatchCashboxSong("CashboxSong", row, "", "", "", Cashbox.SongDataLowCaseList, Cashbox.SongDataFuzzyList);

                                string CashboxId = row["Cashbox_Id"].ToString();

                                if (SongDataIndex >= 0)
                                {
                                    lock (LockThis)
                                    { 
                                        Global.TotalList[0]++;
                                        UpdateList.Add(CashboxId);
                                    }

                                    this.BeginInvoke((Action)delegate()
                                    {
                                        Cashbox_QueryStatus_Label.Text = "已在歌庫找到 " + Global.TotalList[0] + " 首錢櫃歌曲...";
                                    });
                                }
                            }
                        }
                    });
                }

                if (!Directory.Exists(Application.StartupPath + @"\SongMgr")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Backup");
                using (StreamWriter sw = new StreamWriter(Application.StartupPath + @"\SongMgr\CashboxLog.txt"))
                {
                    foreach (string str in UpdateList)
                    {
                        sw.WriteLine(str);
                        Global.TotalList[1]++;

                        this.BeginInvoke((Action)delegate()
                        {
                            Cashbox_QueryStatus_Label.Text = "正在將第 " + Global.TotalList[1] + " 首已有歌曲寫入記錄,請稍待...";
                        });
                    }
                }
                UpdateList.Clear();
            }
        }

        #endregion

    }


    class Cashbox
    {
        
        #region --- Cashbox 建立資料表 ---

        public static List<string> SongIdList;
        public static List<string> SongDataList;
        public static List<string> SongDataLowCaseList;
        public static List<string> SongDataFuzzyList;

        public static void CreateSongDataTable()
        {
            SongIdList = new List<string>();
            SongDataList = new List<string>();
            SongDataLowCaseList = new List<string>();
            SongDataFuzzyList = new List<string>();

            string SongQuerySqlStr = "select Song_Id, Song_Lang, Song_Singer, Song_SongName from ktv_Song order by Song_Id";
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, ""))
            {
                string SongId = string.Empty;
                string SongLang = string.Empty;
                string SongSinger = string.Empty;
                string SongSongName = string.Empty;
                string FuzzyStr = string.Empty;
                foreach (DataRow row in dt.AsEnumerable())
                {
                    SongId = row["Song_Id"].ToString();
                    SongLang = row["Song_Lang"].ToString();
                    SongSinger = row["Song_Singer"].ToString();
                    SongSongName = row["Song_SongName"].ToString();

                    SongIdList.Add(SongId);
                    SongDataList.Add(SongLang + "|" + SongSinger + "|" + SongSongName);
                    SongDataLowCaseList.Add(SongLang + "|" + SongSinger.ToLower() + "|" + SongSongName.ToLower());

                    FuzzyStr = SongLang + "|" + Regex.Replace(SongSinger.ToLower(), @"\s?[\{\(\[｛（［【].+?[】］）｝\]\)\}]\s?|\s|" + Global.CashboxNonSymbolList, "") + "|" + Regex.Replace(SongSongName.ToLower(), @"\s?[\{\(\[｛（［【].+?[】］）｝\]\)\}]\s?|\s|" + Global.CashboxNonSymbolList, "");
                    if (Global.CashboxFullMatchSongList.IndexOf(FuzzyStr) > 0 || Global.GroupSingerLowCaseList.IndexOf(SongSinger.ToLower()) >= 0) FuzzyStr = CommonFunc.GetFuzzyMatchStr("UserSong", SongLang, SongSinger.ToLower(), SongSongName.ToLower(), false, "");
                    SongDataFuzzyList.Add(FuzzyStr);
                }
            }
        }

        public static void DisposeSongDataTable()
        {
            SongIdList.Clear();
            SongDataList.Clear();
            SongDataLowCaseList.Clear();
            GC.Collect();
        }

        #endregion

        #region --- Cashbox 歌曲查詢下拉清單 ---

        public static DataTable GetQueryTypeList()
        {
            using (DataTable list = new DataTable())
            {
                list.Columns.Add(new DataColumn("Display", typeof(string)));
                list.Columns.Add(new DataColumn("Value", typeof(int)));
                List<string> ItemList = new List<string>() { "歌曲名稱", "歌手名稱", "歌曲編號", "全部歌曲", "新進歌曲" };

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

        public static DataTable GetQueryFilterList(List<string> FilterList)
        {
            using (DataTable list = new DataTable())
            {
                list.Columns.Add(new DataColumn("Display", typeof(string)));
                list.Columns.Add(new DataColumn("Value", typeof(int)));
                list.Rows.Add(list.NewRow());
                list.Rows[0][0] = "全部";
                list.Rows[0][1] = 1;

                foreach (string str in FilterList)
                {
                    list.Rows.Add(list.NewRow());
                    list.Rows[list.Rows.Count - 1][0] = str;
                    list.Rows[list.Rows.Count - 1][1] = list.Rows.Count;
                }
                return list;
            }
        }

        public static DataTable GetOtherQueryList()
        {
            using (DataTable list = new DataTable())
            {
                list.Columns.Add(new DataColumn("Display", typeof(string)));
                list.Columns.Add(new DataColumn("Value", typeof(int)));
                List<string> ItemList = new List<string>() { "錢櫃缺歌", "重複歌曲" };

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

        public static DataTable GetDateQueryList()
        {
            using (DataTable list = new DataTable())
            {
                list.Columns.Add(new DataColumn("Display", typeof(string)));
                list.Columns.Add(new DataColumn("Value", typeof(int)));
                List<string> ItemList = new List<string>();

                string QuerySqlStr = "SELECT First(Song_CreatDate) AS Song_CreatDate, Count(Song_CreatDate) AS Song_CreatDateCount FROM ktv_Cashbox GROUP BY Song_CreatDate HAVING (Count(Song_CreatDate)>0) ORDER BY First(Song_CreatDate) desc";
                using (DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, QuerySqlStr, ""))
                {
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow row in dt.AsEnumerable())
                        {
                            string DateItem = DateTime.Parse(row["Song_CreatDate"].ToString()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
                            if (ItemList.IndexOf(DateItem) < 0)
                            {
                                ItemList.Add(DateItem);
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

        #region --- Cashbox 歌曲列表欄位設定 ---

        public static List<string> GetDataGridViewColumnSet(string ColumnName)
        {
            List<string> list = new List<string>();

            string ColumnWidth_120 = Convert.ToInt32((120 * Global.DPIScalingFactor * MainUIScale.UIScalingFactor)).ToString();
            string ColumnWidth_140 = Convert.ToInt32((140 * Global.DPIScalingFactor * MainUIScale.UIScalingFactor)).ToString();
            string ColumnWidth_190 = Convert.ToInt32((190 * Global.DPIScalingFactor * MainUIScale.UIScalingFactor)).ToString();
            string ColumnWidth_320 = Convert.ToInt32((320 * Global.DPIScalingFactor * MainUIScale.UIScalingFactor)).ToString();

            // List<string>() { "欄位名稱", "欄位寬度", "欄位字數" };
            switch (ColumnName)
            {
                case "Cashbox_Id":
                    list = new List<string>() { "歌曲編號", ColumnWidth_120, "6" };
                    break;
                case "Song_Lang":
                    list = new List<string>() { "語系類別", ColumnWidth_120, "none" };
                    break;
                case "Song_Singer":
                    list = new List<string>() { "歌手名稱", ColumnWidth_190, "none" };
                    break;
                case "Song_SongName":
                    list = new List<string>() { "歌曲名稱", ColumnWidth_320, "none" };
                    break;
                case "Song_CreatDate":
                    list = new List<string>() { "更新日期", ColumnWidth_140, "none" };
                    break;
                case "Song_HaveSong":
                    list = new List<string>() { "已有歌曲", "0", "none" };
                    break;
            }
            return list;
        }

        #endregion

    }
}

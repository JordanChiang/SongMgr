using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    // ─────────────────────────────────────────────────────────────────────────
    //  SongNormVol – 音量正規化 tab
    //  Uses Global.SongMaintenanceReplayGainVolume as the base ("基準音量").
    //  Writes the suggested volume to DB (Song_Volume field), NOT to the file.
    // ─────────────────────────────────────────────────────────────────────────
    public partial class MainForm : Form
    {
        // ── Controls (created by InitSongNormVolTab) ──────────────────────────
        private TabPage      NormVol_TabPage;
        private TextBox      NormVol_Singer_TextBox;
        private Button       NormVol_Search_Button;
        private Label        NormVol_BaseVol_Label;      // shows current 基準音量
        private CheckBox     NormVol_FastMode_CheckBox;  // 快速分析 option
        private ComboBox     NormVol_SearchType_ComboBox;
        private DataGridView NormVol_DataGridView;
        private RichTextBox  NormVol_Log_TextBox;
        private Button       NormVol_Start_Button;
        private Button       NormVol_Cancel_Button;
        private Button       NormVol_Write_Button;
        private Label        NormVol_Status_Label;

        // ── State ─────────────────────────────────────────────────────────────
        private CancellationTokenSource NormVol_CTS;
        private ManualResetEventSlim    NormVol_PauseEvent = new ManualResetEventSlim(true);

        // ── Column name constants ─────────────────────────────────────────────
        private const string NV_COL_LANG      = "NormVol_Lang";
        private const string NV_COL_NAME      = "NormVol_SongName";
        private const string NV_COL_SINGER    = "NormVol_Singer";
        private const string NV_COL_TRACKDISP = "NormVol_TrackDisp";   // visible: 音軌
        private const string NV_COL_ORIGVOL   = "NormVol_OrigVol";
        private const string NV_COL_SUGVOL    = "NormVol_SugVol";
        // hidden helper columns
        private const string NV_COL_SONGID    = "NormVol_SongId";
        private const string NV_COL_PATH      = "NormVol_Path";
        private const string NV_COL_FILE      = "NormVol_File";
        private const string NV_COL_TRACK     = "NormVol_Track";       // hidden: raw int for FFmpeg
        private const string NV_COL_RGAIN     = "NormVol_ReplayGain";

        // ─────────────────────────────────────────────────────────────────────
        //  Construction / initialisation
        // ─────────────────────────────────────────────────────────────────────
        public void InitSongNormVolTab()
        {
            Font fontReg  = new Font("微軟正黑體", 12F, FontStyle.Regular, GraphicsUnit.Point, 136);
            Font fontBold = new Font("微軟正黑體", 12F, FontStyle.Bold,    GraphicsUnit.Point, 136);

            // ── Tab page ──────────────────────────────────────────────────────
            NormVol_TabPage = new TabPage
            {
                Name    = "NormVol_TabPage",
                Text    = "音量正規化",
                Padding = new Padding(8, 6, 8, 6),
                Font    = fontReg,
                UseVisualStyleBackColor = true
            };

            // ── Row 1: search criteria ─────────────────────────────────────────
            NormVol_SearchType_ComboBox = new ComboBox
            {
                Name          = "NormVol_SearchType_ComboBox",
                Font          = fontReg,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location      = new Point(8, 10),
                Size          = new Size(100, 29)
            };
            NormVol_SearchType_ComboBox.Items.AddRange(new object[] { "歌手", "歌名", "國語", "台語", "其它語系", "全部歌曲" });
            NormVol_SearchType_ComboBox.SelectedIndex = 0;

            NormVol_Singer_TextBox = new TextBox
            {
                Name     = "NormVol_Singer_TextBox",
                Font     = fontReg,
                Location = new Point(115, 10),
                Size     = new Size(200, 29),
                ImeMode  = ImeMode.OnHalf
            };
            NormVol_Singer_TextBox.KeyPress += NormVol_Singer_TextBox_KeyPress;

            NormVol_Search_Button = new Button
            {
                Name     = "NormVol_Search_Button",
                Text     = "搜尋",
                Font     = fontBold,
                Location = new Point(322, 8),
                Size     = new Size(64, 32)
            };
            NormVol_Search_Button.Click += NormVol_Search_Button_Click;

            // ── 基準音量 display ───────────────────────────────────────────────
            Label baseVolTitleLabel = new Label
            {
                Text     = "基準音量:",
                Font     = fontReg,
                AutoSize = true,
                Location = new Point(400, 14)
            };

            NormVol_BaseVol_Label = new Label
            {
                Name      = "NormVol_BaseVol_Label",
                Text      = Global.SongMaintenanceReplayGainVolume,
                Font      = fontBold,
                AutoSize  = true,
                Location  = new Point(484, 14),
                ForeColor = Color.DarkRed
            };

            // ── 快速分析 checkbox ──────────────────────────────────────────────
            NormVol_FastMode_CheckBox = new CheckBox
            {
                Name     = "NormVol_FastMode_CheckBox",
                Text     = "快速分析",
                Font     = fontReg,
                AutoSize = true,
                Location = new Point(530, 12),
                Checked  = true   // default ON: analyse 30~90 s only
            };
            ToolTip fastModeTip = new ToolTip();
            fastModeTip.SetToolTip(NormVol_FastMode_CheckBox, "勾選：僅分析歌曲 30~90 秒處 (推薦，省時)\n未勾選：從第 5 秒開始分析至歌曲結束前5秒 (完整精確)");

            // ── Status label ──────────────────────────────────────────────────
            NormVol_Status_Label = new Label
            {
                Name      = "NormVol_Status_Label",
                Text      = "",
                Font      = fontReg,
                AutoSize  = true,
                Location  = new Point(8, 48),
                ForeColor = Color.DarkBlue
            };

            // ── DataGridView ──────────────────────────────────────────────────
            NormVol_DataGridView = new DataGridView
            {
                Name                        = "NormVol_DataGridView",
                Dock                        = DockStyle.Fill,
                SelectionMode               = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect                 = false,
                ReadOnly                    = true,
                AllowUserToAddRows          = false,
                AllowUserToDeleteRows       = false,
                RowHeadersVisible           = false,
                // Column header – explicit so 12pt Chinese font is fully visible
                ColumnHeadersVisible             = true,
                ColumnHeadersHeightSizeMode      = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                EnableHeadersVisualStyles         = false,  // must be false to apply custom BackColor/ForeColor
                BackgroundColor             = SystemColors.Window,
                Font                        = fontReg,
                AutoSizeColumnsMode         = DataGridViewAutoSizeColumnsMode.None,
                ScrollBars                  = ScrollBars.Both,
                BorderStyle                 = BorderStyle.Fixed3D
            };
            // Style the column headers to stand out
            NormVol_DataGridView.ColumnHeadersDefaultCellStyle.Font            = fontBold;
            NormVol_DataGridView.ColumnHeadersDefaultCellStyle.BackColor       = Color.FromArgb(70, 130, 180);   // steel blue
            NormVol_DataGridView.ColumnHeadersDefaultCellStyle.ForeColor       = Color.White;
            NormVol_DataGridView.ColumnHeadersDefaultCellStyle.Alignment       = DataGridViewContentAlignment.MiddleCenter;
            NormVol_DataGridView.MakeDoubleBuffered(true);
            NormVol_DataGridView.CellDoubleClick += NormVol_DataGridView_CellDoubleClick;

            // Visible columns – first column frozen so it stays in view when scrolling
            NormVol_AddColumn(NV_COL_LANG,      "語系",    60,  true,  true);
            NormVol_AddColumn(NV_COL_NAME,      "歌名",    220, true,  false);
            NormVol_AddColumn(NV_COL_SINGER,    "歌手",    120, true,  false);
            NormVol_AddColumn(NV_COL_TRACKDISP, "音軌",    50,  true,  false);
            NormVol_AddColumn(NV_COL_ORIGVOL,   "原始音量", 80,  true,  false);
            NormVol_AddColumn(NV_COL_SUGVOL,    "建議音量", 80,  true,  false);
            // Hidden helper columns
            NormVol_AddColumn(NV_COL_SONGID,    "歌曲ID",     0, false,  false);
            NormVol_AddColumn(NV_COL_PATH,      "路徑",       0, false,  false);
            NormVol_AddColumn(NV_COL_FILE,      "檔名",       0, false,  false);
            NormVol_AddColumn(NV_COL_TRACK,     "Track",     0, false,  false);
            NormVol_AddColumn(NV_COL_RGAIN,     "ReplayGain", 0, false, false);

            // ── Header panel (Dock.Top) – contains all row-1 controls + status ─
            Panel headerPanel = new Panel
            {
                Dock    = DockStyle.Top,
                Height  = 76,
                Padding = new Padding(0)
            };
            headerPanel.Controls.Add(NormVol_SearchType_ComboBox);
            headerPanel.Controls.Add(NormVol_Singer_TextBox);
            headerPanel.Controls.Add(NormVol_Search_Button);
            headerPanel.Controls.Add(baseVolTitleLabel);
            headerPanel.Controls.Add(NormVol_BaseVol_Label);
            headerPanel.Controls.Add(NormVol_FastMode_CheckBox);
            headerPanel.Controls.Add(NormVol_Status_Label);

            // ── Bottom buttons ────────────────────────────────────────────────
            NormVol_Start_Button = new Button
            {
                Name    = "NormVol_Start_Button",
                Text    = "開始分析",
                Font    = fontBold,
                Size    = new Size(110, 32),
                Enabled = false
            };
            NormVol_Start_Button.Click += NormVol_Start_Button_Click;

            NormVol_Cancel_Button = new Button
            {
                Name    = "NormVol_Cancel_Button",
                Text    = "取消",
                Font    = fontBold,
                Size    = new Size(110, 32),
                Enabled = false
            };
            NormVol_Cancel_Button.Click += NormVol_Cancel_Button_Click;

            NormVol_Write_Button = new Button
            {
                Name    = "NormVol_Write_Button",
                Text    = "寫入資料庫",
                Font    = fontBold,
                Size    = new Size(110, 32),
                Enabled = false
            };
            NormVol_Write_Button.Click += NormVol_Write_Button_Click;

            FlowLayoutPanel btnPanel = new FlowLayoutPanel
            {
                Dock          = DockStyle.Bottom,
                Height        = 42,
                FlowDirection = FlowDirection.LeftToRight,
                Padding       = new Padding(0, 4, 0, 4),
                WrapContents  = false
            };
            btnPanel.Controls.Add(NormVol_Start_Button);
            btnPanel.Controls.Add(NormVol_Cancel_Button);
            btnPanel.Controls.Add(NormVol_Write_Button);

            // ── Log TextBox ──────────────────────────────────────────────────
            NormVol_Log_TextBox = new RichTextBox
            {
                Name            = "NormVol_Log_TextBox",
                Dock            = DockStyle.Fill,
                ReadOnly        = true,
                BackColor       = Color.Black,
                ForeColor       = Color.LightGreen,
                Font            = new Font("Consolas", 10F, FontStyle.Regular),
                BorderStyle     = BorderStyle.Fixed3D,
                WordWrap        = false
            };

            // ── Footer panel (Dock.Bottom) – contains Log + Buttons ──────────
            Panel footerPanel = new Panel
            {
                Dock   = DockStyle.Bottom,
                Height = 180   // log area height
            };
            footerPanel.Controls.Add(NormVol_Log_TextBox);
            footerPanel.Controls.Add(btnPanel);

            // ── Assemble tab ──────────────────────────────────────────────────
            // WinForms dock layout processes controls in REVERSE Controls-index order
            // (highest index = processed first). Correct order:
            //   index 0 → DataGridView (Dock.Fill)   processed LAST  → gets remaining space
            //   index 1 → footerPanel  (Dock.Bottom)  processed FIRST → takes bottom strip
            //   index 2 → headerPanel  (Dock.Top)     processed FIRST → takes top strip
            NormVol_TabPage.Controls.Add(NormVol_DataGridView);  // 0 – Fill
            NormVol_TabPage.Controls.Add(footerPanel);           // 1 – Bottom
            NormVol_TabPage.Controls.Add(headerPanel);           // 2 – Top

            MainTabControl.TabPages.Add(NormVol_TabPage);
        }

        private void NormVol_AppendLog(string text, Color? color = null)
        {
            Console.WriteLine(text);
            if (NormVol_Log_TextBox == null) return;
            if (NormVol_Log_TextBox.InvokeRequired)
            {
                NormVol_Log_TextBox.Invoke((Action)(() => NormVol_AppendLog(text, color)));
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            Color textColor = color ?? ((text.Contains("失敗") || text.Contains("錯誤")) ? Color.Red : Color.LightGreen);

            NormVol_Log_TextBox.SelectionStart = NormVol_Log_TextBox.TextLength;
            NormVol_Log_TextBox.SelectionLength = 0;
            NormVol_Log_TextBox.SelectionColor = textColor;
            NormVol_Log_TextBox.AppendText(string.Format("[{0}] {1}\r\n", timestamp, text));
            NormVol_Log_TextBox.SelectionColor = NormVol_Log_TextBox.ForeColor;
            NormVol_Log_TextBox.SelectionStart = NormVol_Log_TextBox.Text.Length;
            NormVol_Log_TextBox.ScrollToCaret();
        }

        private void NormVol_AddColumn(string name, string header, int width, bool visible, bool frozen = false)
        {
            var col = new DataGridViewTextBoxColumn
            {
                Name         = name,
                HeaderText   = header,
                Width        = visible ? width : 0,
                MinimumWidth = visible ? Math.Max(width, 2) : 2,
                Visible      = visible,
                Frozen       = frozen,
                ReadOnly     = true,
                SortMode     = DataGridViewColumnSortMode.NotSortable
            };
            NormVol_DataGridView.Columns.Add(col);
        }

        // Called when the tab becomes visible so the label reflects any change
        // made in the SongMaintenance tab.
        private void NormVol_RefreshBaseVol()
        {
            if (NormVol_BaseVol_Label != null)
                NormVol_BaseVol_Label.Text = Global.SongMaintenanceReplayGainVolume;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Search
        // ─────────────────────────────────────────────────────────────────────
        private void NormVol_Singer_TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == 13) NormVol_Search_Button_Click(sender, EventArgs.Empty);
        }

        private void NormVol_Search_Button_Click(object sender, EventArgs e)
        {
            string keyword = NormVol_Singer_TextBox.Text.Trim();
            int searchType = NormVol_SearchType_ComboBox.SelectedIndex;

            // Optional validation: "歌手" or "歌名" modes usually require keywords
            if ((searchType == 0 || searchType == 1) && string.IsNullOrEmpty(keyword))
            {
                NormVol_Status_Label.Text = "請輸入搜尋關鍵字。";
                return;
            }

            if (!Global.CrazyktvDatabaseStatus)
            {
                NormVol_Status_Label.Text = "資料庫尚未連線。";
                return;
            }

            NormVol_RefreshBaseVol();
            NormVol_DataGridView.Rows.Clear();
            NormVol_Status_Label.Text     = "正在搜尋...";
            NormVol_Start_Button.Enabled  = false;
            NormVol_Write_Button.Enabled  = false;
            NormVol_Cancel_Button.Enabled = false;

            Task.Factory.StartNew(() =>
            {
                string whereClause = "";
                List<OleDbParameter> parameters = new List<OleDbParameter>();

                switch (searchType)
                {
                    case 0: // 歌手
                        whereClause = "WHERE Song_Singer LIKE @Keyword ";
                        parameters.Add(new OleDbParameter("@Keyword", "%" + keyword + "%"));
                        break;
                    case 1: // 歌名
                        whereClause = "WHERE Song_SongName LIKE @Keyword ";
                        parameters.Add(new OleDbParameter("@Keyword", "%" + keyword + "%"));
                        break;
                    case 2: // 國語
                        whereClause = "WHERE Song_Lang = '國語' ";
                        if (!string.IsNullOrEmpty(keyword))
                        {
                            whereClause += "AND Song_Singer LIKE @Keyword ";
                            parameters.Add(new OleDbParameter("@Keyword", "%" + keyword + "%"));
                        }
                        break;
                    case 3: // 台語
                        whereClause = "WHERE Song_Lang = '台語' ";
                        if (!string.IsNullOrEmpty(keyword))
                        {
                            whereClause += "AND Song_Singer LIKE @Keyword ";
                            parameters.Add(new OleDbParameter("@Keyword", "%" + keyword + "%"));
                        }
                        break;
                    case 4: // 其它
                        whereClause = "WHERE Song_Lang NOT IN ('國語', '台語') ";
                        if (!string.IsNullOrEmpty(keyword))
                        {
                            whereClause += "AND Song_Singer LIKE @Keyword ";
                            parameters.Add(new OleDbParameter("@Keyword", "%" + keyword + "%"));
                        }
                        break;
                    case 5: // 全部歌曲
                        whereClause = "";
                        break;
                }

                string sql =
                    "SELECT Song_Id, Song_Lang, Song_SongName, Song_Singer, " +
                    "       Song_Volume, Song_Track, Song_Path, Song_FileName, Song_ReplayGain " +
                    "FROM ktv_Song " +
                    whereClause +
                    "ORDER BY Song_Lang, Song_Singer, Song_SongName";

                DataTable dt = null;
                try
                {
                    using (OleDbConnection conn = new OleDbConnection(
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Global.CrazyktvDatabaseFile))
                    {
                        conn.Open();
                        using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                        {
                            foreach (var p in parameters) cmd.Parameters.Add(p);

                            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                            dt = new DataTable();
                            da.Fill(dt);
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)delegate ()
                    {
                        NormVol_Status_Label.Text = "搜尋失敗: " + ex.Message;
                    });
                    return;
                }

                if (dt != null)
                {
                    this.BeginInvoke((Action)delegate ()
                    {
                        NormVol_DataGridView.Rows.Clear();
                        foreach (DataRow row in dt.Rows)
                        {
                            string sLang   = row["Song_Lang"]?.ToString()     ?? "";
                            string sName   = row["Song_SongName"]?.ToString() ?? "";
                            string sSinger = row["Song_Singer"]?.ToString()   ?? "";
                            string sTrack  = row["Song_Track"]?.ToString()    ?? "1";
                            string sVol    = row["Song_Volume"]?.ToString()   ?? "100";
                            string sSug    = ""; 
                            string sId     = row["Song_Id"]?.ToString()       ?? "";
                            string sPath   = row["Song_Path"]?.ToString()     ?? "";
                            string sFile   = row["Song_FileName"]?.ToString() ?? "";
                            string sRG     = row["Song_ReplayGain"]?.ToString() ?? "";

                            // Must match the 11 columns in InitSongNormVolTab
                            NormVol_DataGridView.Rows.Add(
                                sLang, sName, sSinger, sTrack, sVol, sSug, 
                                sId, sPath, sFile, sTrack, sRG
                            );
                        }

                        NormVol_Status_Label.Text = string.Format("搜尋完成，共 {0} 首歌曲。", dt.Rows.Count);
                        NormVol_Start_Button.Enabled = (dt.Rows.Count > 0);
                    });
                }
            });
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Row double-click → play via DShowForm
        // ─────────────────────────────────────────────────────────────────────
        private void NormVol_DataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            DataGridViewRow row = NormVol_DataGridView.Rows[e.RowIndex];
            string songPath = row.Cells[NV_COL_PATH].Value?.ToString() ?? "";
            string songFile = row.Cells[NV_COL_FILE].Value?.ToString() ?? "";
            string filePath = Path.Combine(songPath, songFile);

            if (!File.Exists(filePath))
            {
                string msg = "【" + songFile + "】檔案不存在。";
                NormVol_Status_Label.Text = msg;
                NormVol_AppendLog(string.Format("播放失敗: {0} ({1})", msg, filePath));
                return;
            }

            string songId     = row.Cells[NV_COL_SONGID].Value?.ToString() ?? "";
            string songLang   = row.Cells[NV_COL_LANG].Value?.ToString()   ?? "";
            string singer     = row.Cells[NV_COL_SINGER].Value?.ToString()  ?? "";
            string songName   = row.Cells[NV_COL_NAME].Value?.ToString()    ?? "";
            string track      = row.Cells[NV_COL_TRACK].Value?.ToString()   ?? "1";
            string volume     = row.Cells[NV_COL_ORIGVOL].Value?.ToString() ?? "100";
            string replayGain = row.Cells[NV_COL_RGAIN].Value?.ToString()   ?? "";

            var playerInfo = new List<string>
            {
                songId, songLang, singer, songName,
                track, volume, replayGain, filePath,
                e.RowIndex.ToString(), "NormVol"
            };

            Global.PlayerUpdateSongValueList = new List<string>();
            if (Global.MainCfgPlayerCore == "1")
            {
                DShowForm playerForm = new DShowForm(this, playerInfo);
                playerForm.Show();
            }
            this.Hide();
        }

        // ─────────────────────────────────────────────────────────────────────
        //  開始 – run FFmpeg replaygain; suggested volume uses 基準音量 from Global
        // ─────────────────────────────────────────────────────────────────────
        private void NormVol_Start_Button_Click(object sender, EventArgs e)
        {
            if (NormVol_Start_Button.Text == "停止分析")
            {
                NormVol_CTS?.Cancel();
                NormVol_PauseEvent.Set(); // Resume if paused so it can see the cancellation
                NormVol_Start_Button.Enabled = false;
                return;
            }

            if (NormVol_DataGridView.Rows.Count == 0) return;

            // Read base volume from Global (same value as SongMaintenance_ReplayGain_TextBox)
            int baseVolume;
            if (!int.TryParse(Global.SongMaintenanceReplayGainVolume, out baseVolume) || baseVolume <= 0)
                baseVolume = 50;

            NormVol_CTS = new CancellationTokenSource();
            CancellationToken ct = NormVol_CTS.Token;
            NormVol_PauseEvent.Set(); // Reset to running state

            NormVol_Start_Button.Text     = "停止分析";
            NormVol_Start_Button.Enabled  = true;
            NormVol_Cancel_Button.Text    = "暫停分析";
            NormVol_Cancel_Button.Enabled = true;
            NormVol_Write_Button.Enabled  = false;
            NormVol_Status_Label.Text     = string.Format("正在分析音量 (基準音量={0})，請稍待...", baseVolume);

            NormVol_Log_TextBox.Clear();
            NormVol_AppendLog(string.Format("=== 開始音量分析 (基準音量: {0}) ===", baseVolume), Color.Cyan);

            // Snapshot row data
            bool fastMode = NormVol_FastMode_CheckBox.Checked;
            var rows = new List<(int idx, string lang, string name, string singer, string path, string file, int track)>();
            foreach (DataGridViewRow row in NormVol_DataGridView.Rows)
            {
                int t;
                if (!int.TryParse(row.Cells[NV_COL_TRACK].Value?.ToString(), out t) || t < 1)
                    t = 1;

                rows.Add((
                    row.Index,
                    row.Cells[NV_COL_LANG].Value?.ToString()   ?? "",
                    row.Cells[NV_COL_NAME].Value?.ToString()   ?? "",
                    row.Cells[NV_COL_SINGER].Value?.ToString() ?? "",
                    row.Cells[NV_COL_PATH].Value?.ToString()   ?? "",
                    row.Cells[NV_COL_FILE].Value?.ToString()   ?? "",
                    t
                ));
            }

            Task.Factory.StartNew(() =>
            {
                int done  = 0;
                int total = rows.Count;

                foreach (var r in rows)
                {
                    NormVol_PauseEvent.Wait(ct);
                    if (ct.IsCancellationRequested) break;

                    Stopwatch sw = Stopwatch.StartNew();
                    string sugVol   = "";
                    string filePath = Path.Combine(r.path, r.file);
                    string errorMsg = "";

                    if (!File.Exists(filePath))
                    {
                        errorMsg = "檔案不存在 (" + filePath + ")";
                    }
                    else
                    {
                        // Compute seekArgs in background to avoid UI hang
                        string seekArgs;
                        if (fastMode)
                        {
                            seekArgs = "-ss 30 -t 60";
                        }
                        else
                        {
                            double dur = FFmpeg.GetFileDuration(filePath);
                            seekArgs = (dur > 15)
                                ? string.Format(System.Globalization.CultureInfo.InvariantCulture, "-ss 5 -t {0:F2}", Math.Max(dur - 10, 1))
                                : "";
                        }

                        try
                        {
                            // Use Song_Track + seek args (fast/full mode)
                            FFmpeg.SongVolumeValue sv = FFmpeg.GetSongVolume(filePath, r.track, seekArgs);
                            if (sv.Success)
                            {
                                sugVol = FFmpeg.CalSongVolume(baseVolume, sv.GainDB);
                                // Clamp suggested volume to valid range [5, 100]
                                int volClamped;
                                if (int.TryParse(sugVol, out volClamped))
                                    sugVol = Math.Max(5, Math.Min(100, volClamped)).ToString();
                            }
                            else
                            {
                                errorMsg = string.IsNullOrEmpty(sv.ErrorMessage) ? "分析失敗" : sv.ErrorMessage;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorMsg = "分析失敗: " + ex.Message;
                            sugVol = "";
                        }
                    }
                    sw.Stop();

                    int    rowIdx        = r.idx;
                    string sugVolCapture = sugVol;
                    done++;
                    int doneCapture = done;

                    string logMsg;
                    if (!string.IsNullOrEmpty(errorMsg))
                    {
                        logMsg = string.Format(
                            "[{0}] {1} ({2}) | 音軌:{3} | 失敗: {4}",
                            r.lang, r.name, r.singer, r.track, errorMsg);
                    }
                    else
                    {
                        logMsg = string.Format(
                            "[{0}] {1} ({2}) | 音軌:{3} | 建議:{4} | 耗時:{5:F2}s",
                            r.lang, r.name, r.singer, r.track,
                            sugVolCapture,
                            sw.Elapsed.TotalSeconds);
                    }

                    this.BeginInvoke((Action)delegate ()
                    {
                        if (rowIdx < NormVol_DataGridView.Rows.Count)
                            NormVol_DataGridView.Rows[rowIdx].Cells[NV_COL_SUGVOL].Value = sugVolCapture;
                        NormVol_Status_Label.Text = string.Format("已分析 {0} / {1} 首...", doneCapture, total);
                        NormVol_AppendLog(logMsg);
                    });
                }

                bool cancelled = ct.IsCancellationRequested;
                this.BeginInvoke((Action)delegate ()
                {
                    NormVol_Start_Button.Text     = "開始分析";
                    NormVol_Start_Button.Enabled  = true;
                    NormVol_Cancel_Button.Text    = "取消";
                    NormVol_Cancel_Button.Enabled = false;
                    NormVol_Write_Button.Enabled  = !cancelled;
                    NormVol_Status_Label.Text = cancelled
                        ? "已取消分析。"
                        : string.Format("分析完成，共 {0} 首，按「寫入資料庫」可將建議音量存入資料庫。", total);
                });
            }, ct);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  取消
        // ─────────────────────────────────────────────────────────────────────
        private void NormVol_Cancel_Button_Click(object sender, EventArgs e)
        {
            if (NormVol_Cancel_Button.Text == "暫停分析")
            {
                NormVol_PauseEvent.Reset();
                NormVol_Cancel_Button.Text = "繼續分析";
                NormVol_Status_Label.Text  = "分析已暫停。";
                return;
            }
            if (NormVol_Cancel_Button.Text == "繼續分析")
            {
                NormVol_PauseEvent.Set();
                NormVol_Cancel_Button.Text = "暫停分析";
                NormVol_Status_Label.Text  = "正在分析音量...";
                return;
            }

            // Fallback for "取消" or manual cancel
            NormVol_CTS?.Cancel();
            NormVol_PauseEvent.Set(); // Resume if paused so it can cancel
            NormVol_Cancel_Button.Enabled = false;
            NormVol_Status_Label.Text = "正在取消...";
        }

        // ─────────────────────────────────────────────────────────────────────
        //  寫入 – save SUGGESTED volume to DB (Song_Volume field), not the file
        // ─────────────────────────────────────────────────────────────────────
        private void NormVol_Write_Button_Click(object sender, EventArgs e)
        {
            if (NormVol_DataGridView.Rows.Count == 0) return;

            // Collect rows that have a non-empty suggested volume
            var updateList = new List<(string songId, string sugVol)>();
            foreach (DataGridViewRow row in NormVol_DataGridView.Rows)
            {
                string sug = row.Cells[NV_COL_SUGVOL].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(sug))
                    updateList.Add((row.Cells[NV_COL_SONGID].Value?.ToString() ?? "", sug));
            }

            if (updateList.Count == 0)
            {
                NormVol_Status_Label.Text = "尚無建議音量可寫入，請先按「開始」分析。";
                return;
            }

            if (MessageBox.Show(
                    string.Format("確定要將 {0} 首歌曲的建議音量寫入資料庫嗎?", updateList.Count),
                    "確認寫入", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            NormVol_Write_Button.Enabled  = false;
            NormVol_Start_Button.Enabled  = false;
            NormVol_Status_Label.Text     = "正在寫入資料庫...";

            Task.Factory.StartNew(() =>
            {
                int success = 0, fail = 0;
                try
                {
                    using (OleDbConnection conn = new OleDbConnection(
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Global.CrazyktvDatabaseFile))
                    {
                        conn.Open();
                        const string sql = "UPDATE ktv_Song SET Song_Volume = @Vol WHERE Song_Id = @SongId";
                        using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                        {
                            foreach (var item in updateList)
                            {
                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("@Vol",    item.sugVol);
                                cmd.Parameters.AddWithValue("@SongId", item.songId);
                                try { cmd.ExecuteNonQuery(); success++; }
                                catch { fail++; }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)delegate ()
                    {
                        NormVol_Status_Label.Text    = "寫入失敗: " + ex.Message;
                        NormVol_Write_Button.Enabled = true;
                        NormVol_Start_Button.Enabled = true;
                    });
                    return;
                }

                this.BeginInvoke((Action)delegate ()
                {
                    // Sync the 原始音量 column so the grid reflects the saved values
                    foreach (DataGridViewRow row in NormVol_DataGridView.Rows)
                    {
                        string sug = row.Cells[NV_COL_SUGVOL].Value?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(sug))
                        {
                            row.Cells[NV_COL_ORIGVOL].Value = sug;
                            row.Cells[NV_COL_SUGVOL].Value  = "";
                        }
                    }
                    NormVol_Status_Label.Text    = string.Format("寫入完成，成功 {0} 首，失敗 {1} 首。", success, fail);
                    NormVol_Start_Button.Enabled = true;
                });
            });
        }
    }
}

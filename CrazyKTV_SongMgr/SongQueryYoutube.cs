using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 歌庫查詢 → 類型「YouTube 歌曲」: lists YoutubeSongList in SongQuery_DataGridView with its own columns.
    /// In 編輯模式 the editable columns are changed in the grid itself and saved as soon as a cell is left.
    /// </summary>
    public partial class MainForm : Form
    {
        /// <summary>Grid columns: YoutubeSongList column, header, width and the column it saves to (null = read only).</summary>
        private static readonly (string Name, string Header, int Width, YoutubeColumn? Editable)[] SongQuery_YoutubeColumns =
        {
            ("Id", "編號", 70, null),
            ("Name", "名稱", 360, YoutubeColumn.Name),
            ("Url", "影片 ID", 110, null),
            ("Length", "長度", 60, null),
            ("PlayCount", "點播次數", 80, YoutubeColumn.PlayCount),
            ("Track", "聲道", 150, YoutubeColumn.Track),
            ("Volume", "音量", 60, YoutubeColumn.Volume),
            ("FileName", "檔名", 300, null),
            ("FileSize", "檔案大小(MB)", 100, null),
            ("CreatDate", "加入日期", 140, null),
        };

        /// <summary>The YoutubeSongList rows on show; the grid is in YouTube mode while it is the DataSource.</summary>
        private DataTable SongQuery_YoutubeTable;

        private bool SongQuery_IsYoutubeMode => SongQuery_YoutubeTable != null && SongQuery_DataGridView.DataSource == SongQuery_YoutubeTable;

        private void SongQuery_InitializeYoutubeEvents()
        {
            SongQuery_DataGridView.DataSourceChanged += SongQuery_DataGridView_DataSourceChanged;
            SongQuery_DataGridView.CellValidating += SongQuery_DataGridView_CellValidating;
            SongQuery_DataGridView.CellParsing += SongQuery_DataGridView_CellParsing;
            SongQuery_DataGridView.CellValueChanged += SongQuery_DataGridView_CellValueChanged;
            SongQuery_DataGridView.CurrentCellDirtyStateChanged += SongQuery_DataGridView_CurrentCellDirtyStateChanged;
            SongQuery_DataGridView.DataError += SongQuery_DataGridView_DataError;
        }

        #region --- SongQuery YouTube 查詢 ---

        /// <summary>Lists every YouTube 歌曲, or those whose 名稱 contains <paramref name="keyword"/>.</summary>
        private void SongQuery_YoutubeQuery(string keyword)
        {
            SongQuery_DataGridView.DataSource = null;
            if (SongQuery_DataGridView.Columns.Contains("Song_FullPath")) SongQuery_DataGridView.Columns.Remove("Song_FullPath");
            SongQuery_EditMode_CheckBox.Enabled = false;
            SongQuery_LangFilter_ComboBox.Enabled = false;
            SongQuery_WordCountFilter_ComboBox.Enabled = false;

            try
            {
                string CreatedText = "";
                if (!YoutubeSongStore.TableExists(Global.CrazyktvDatabaseFile))
                {
                    if (MessageBox.Show("此資料庫沒有 YouTube 歌曲資料表,是否要建立?", "建立提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    {
                        SongQuery_QueryStatus_Label.Text = "此資料庫沒有 YouTube 歌曲資料表...";
                        return;
                    }
                    YoutubeSongStore.CreateTable(Global.CrazyktvDatabaseFile);
                    CreatedText = "已建立 YouTube 歌曲資料表。";
                }

                DataTable dt = YoutubeSongStore.Query(Global.CrazyktvDatabaseFile, keyword);
                SongQuery_YoutubeTable = dt;
                SongQuery_DataGridView.DataSource = dt;
                SongQuery_YoutubeSetupColumns(dt);
                SongQuery_EditMode_CheckBox.Enabled = dt.Rows.Count > 0;
                SongQuery_YoutubeShowCheckButton(true);
                SongQuery_YoutubeCheck_Button.Enabled = dt.Rows.Count > 0;

                string QueryText = (keyword.Trim() == "") ? "YouTube 歌曲" : "名稱包含『" + keyword.Trim() + "』的 YouTube 歌曲";
                SongQuery_QueryStatus_Label.Text = CreatedText + ((dt.Rows.Count > 0) ? "總共查詢到 " + dt.Rows.Count + " 筆" + QueryText + "。" : "查無" + QueryText + "...");
            }
            catch (Exception ex)
            {
                SongQuery_QueryStatus_Label.Text = "查詢 YouTube 歌曲失敗: " + ex.Message;
            }
        }

        private void SongQuery_YoutubeSetupColumns(DataTable dt)
        {
            SongQuery_YoutubeReplaceWithComboBox("Track", combo =>
            {
                // Value has the Track column's own type (Int32 on Access, Int64 on SQLite), or no row would match it
                Type TrackType = dt.Columns["Track"].DataType;
                DataTable TrackList = new DataTable();
                TrackList.Columns.Add("Display", typeof(string));
                TrackList.Columns.Add("Value", TrackType);
                foreach (var track in YoutubeSongStore.Tracks) TrackList.Rows.Add(track.Display, Convert.ChangeType(track.Value, TrackType));
                combo.DataSource = TrackList;
                combo.DisplayMember = "Display";
                combo.ValueMember = "Value";
            });

            for (int i = 0; i < SongQuery_YoutubeColumns.Length; i++)
            {
                var spec = SongQuery_YoutubeColumns[i];
                DataGridViewColumn column = SongQuery_DataGridView.Columns[spec.Name];
                column.HeaderText = spec.Header;
                column.Width = Math.Max(Convert.ToInt32(spec.Width * Global.DPIScalingFactor * MainUIScale.UIScalingFactor), SongQuery_YoutubeHeaderWidth(column));
                column.ReadOnly = spec.Editable == null;
                column.DisplayIndex = i;
            }
            ((DataGridViewTextBoxColumn)SongQuery_DataGridView.Columns["Name"]).MaxInputLength = 255;
        }

        /// <summary>
        /// The width that keeps the header on one line (the header row grows to fit, so a narrow column wraps 長度 into
        /// 長 / 度): what the grid itself needs for the header text in the header font, with its margins and sort arrow.
        /// </summary>
        private static int SongQuery_YoutubeHeaderWidth(DataGridViewColumn column)
        {
            return column.GetPreferredWidth(DataGridViewAutoSizeColumnMode.ColumnHeader, true);
        }

        private void SongQuery_YoutubeReplaceWithComboBox(string name, Action<DataGridViewComboBoxColumn> fill)
        {
            DataGridViewColumn TextColumn = SongQuery_DataGridView.Columns[name];
            int index = TextColumn.Index;
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn()
            {
                Name = name,
                DataPropertyName = name,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                FlatStyle = FlatStyle.Flat,
                SortMode = DataGridViewColumnSortMode.Automatic,
            };
            fill(combo);
            SongQuery_DataGridView.Columns.Remove(TextColumn);
            SongQuery_DataGridView.Columns.Insert(index, combo);
        }

        /// <summary>Leaving the YouTube list: drop its hand-made columns and the editing it switched on.</summary>
        private void SongQuery_DataGridView_DataSourceChanged(object sender, EventArgs e)
        {
            if (SongQuery_YoutubeTable == null || SongQuery_DataGridView.DataSource == SongQuery_YoutubeTable) return;

            SongQuery_YoutubeTable = null;
            SongQuery_YoutubeCheckCancel?.Cancel();
            SongQuery_YoutubeShowCheckButton(false);
            // The 聲道 drop-down was added by hand, so unbinding does not remove it.
            foreach (DataGridViewColumn column in SongQuery_DataGridView.Columns.Cast<DataGridViewColumn>().Where(c => c is DataGridViewComboBoxColumn).ToList())
            {
                SongQuery_DataGridView.Columns.Remove(column);
            }
            SongQuery_DataGridView.EditMode = DataGridViewEditMode.EditProgrammatically;
            if (SongQuery_EditMode_CheckBox.Checked) SongQuery_EditMode_CheckBox.Checked = false;
        }

        #endregion

        #region --- SongQuery YouTube 列表 ---

        private void SongQuery_YoutubeCellFormatting(DataGridViewCellFormattingEventArgs e)
        {
            switch (SongQuery_DataGridView.Columns[e.ColumnIndex].Name)
            {
                case "Length":
                    e.Value = YoutubeSongStore.FormatLength(e.Value);
                    e.FormattingApplied = true;
                    break;
                case "FileSize":
                    e.Value = YoutubeSongStore.FormatMB(e.Value);
                    e.FormattingApplied = true;
                    break;
                case "CreatDate":
                    if (e.Value is DateTime date)
                    {
                        e.Value = date.ToString("yyyy/MM/dd HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                        e.FormattingApplied = true;
                    }
                    break;
                case "Id":
                case "Url":
                    e.CellStyle.ForeColor = Color.Blue;
                    break;
                case SongQuery_YoutubeCheckColumn:
                    string result = Convert.ToString(e.Value);
                    if (result.StartsWith("已失效", StringComparison.Ordinal)) e.CellStyle.ForeColor = Color.Red;
                    else if (result.StartsWith("無法確認", StringComparison.Ordinal)) e.CellStyle.ForeColor = Color.Gray;
                    break;
            }
        }

        private void SongQuery_YoutubeCellDoubleClick(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (SongQuery_EditMode_CheckBox.Checked)
            {
                if (e.ColumnIndex >= 0 && !SongQuery_DataGridView.Columns[e.ColumnIndex].ReadOnly) SongQuery_DataGridView.BeginEdit(true);
            }
            else
            {
                SongQuery_YoutubeOpenInBrowser(e.RowIndex);
            }
        }

        private void SongQuery_YoutubeShowContextMenu(DataGridViewCellMouseEventArgs e)
        {
            if (!SongQuery_DataGridView.Rows[e.RowIndex].Selected) SongQuery_DataGridView.CurrentCell = SongQuery_DataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];

            object VideoId = SongQuery_DataGridView.Rows[e.RowIndex].Cells["Url"].Value;
            ContextMenuStrip GridView_ContextMenu = new ContextMenuStrip();
            GridView_ContextMenu.Items.Add("在瀏覽器開啟", null, (s, a) => SongQuery_YoutubeOpenInBrowser(VideoId));
            GridView_ContextMenu.Items.Add("複製影片網址", null, (s, a) => SongQuery_YoutubeCopyUrl(VideoId));
            if (Global.FavoriteUserDT.Rows.Count > 0)
            {
                ToolStripMenuItem FavoriteItem = new ToolStripMenuItem("加入我的最愛");
                foreach (DataRow row in Global.FavoriteUserDT.AsEnumerable())
                {
                    string UserId = row["User_Id"].ToString();
                    string UserName = row["User_Name"].ToString();
                    FavoriteItem.DropDown.Items.Add(UserName, null, (s, a) => SongQuery_YoutubeAddFavorites(UserId, UserName));
                }
                EnableDropDownMouseWheelScroll(FavoriteItem.DropDown);
                GridView_ContextMenu.Items.Add(FavoriteItem);
            }
            GridView_ContextMenu.Show(MousePosition.X, MousePosition.Y);
        }

        private void SongQuery_YoutubeCopyUrl(object VideoId)
        {
            string url = YoutubeSongStore.WatchUrl(VideoId);
            if (url == "")
            {
                SongQuery_QueryStatus_Label.Text = "此筆資料沒有影片 ID...";
                return;
            }
            Clipboard.SetText(url);
            SongQuery_QueryStatus_Label.Text = "已複製影片網址: " + url;
        }

        private void SongQuery_YoutubeOpenInBrowser(int RowIndex)
        {
            SongQuery_YoutubeOpenInBrowser(SongQuery_DataGridView.Rows[RowIndex].Cells["Url"].Value);
        }

        private void SongQuery_YoutubeOpenInBrowser(object VideoId)
        {
            string url = YoutubeSongStore.WatchUrl(VideoId);
            if (url == "")
            {
                SongQuery_QueryStatus_Label.Text = "此筆資料沒有影片 ID...";
                return;
            }
            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                SongQuery_QueryStatus_Label.Text = "無法開啟瀏覽器: " + ex.Message;
            }
        }

        /// <summary>Puts the selected YouTube 歌曲 on the user's 我的最愛 (UltimateKtv's YoutubeFavorite).</summary>
        private void SongQuery_YoutubeAddFavorites(string UserId, string UserName)
        {
            List<YoutubeFavoriteItem> items = SongQuery_DataGridView.SelectedRows.Cast<DataGridViewRow>()
                .OrderBy(row => row.Index)
                .Select(row => new YoutubeFavoriteItem(
                    Convert.ToString(row.Cells["Url"].Value).Trim(),
                    Convert.ToString(row.Cells["Name"].Value),
                    (row.Cells["Length"].Value == null || row.Cells["Length"].Value == DBNull.Value) ? 0 : Convert.ToInt32(row.Cells["Length"].Value)))
                .ToList();
            int NoId = items.Count(item => !YoutubeFavorites.IsValidVideoId(item.VideoId));

            try
            {
                int added;
                using (SongLibraryStore store = SongLibraryStore.OpenOrThrow())
                {
                    added = store.AddYoutubeFavorites(UserId, items);
                }
                string skipped = (items.Count - added - NoId > 0) ? ", " + (items.Count - added - NoId) + " 筆已在清單中" : "";
                string invalid = (NoId > 0) ? ", " + NoId + " 筆沒有影片 ID" : "";
                SongQuery_QueryStatus_Label.Text = "已將 " + added + " 筆 YouTube 歌曲加入【" + UserName + "】的我的最愛" + skipped + invalid + "。";
            }
            catch (Exception ex)
            {
                SongQuery_QueryStatus_Label.Text = "加入我的最愛失敗: " + ex.Message;
            }
        }

        #region --- SongQuery YouTube 我的最愛 ---

        /// <summary>A row of a 我的最愛 list that is a YouTube 我的最愛 (Song_Id holds the video ID) rather than a song.</summary>
        private bool SongQuery_IsYoutubeFavoriteRow(int RowIndex)
        {
            if (Global.SongQueryQueryType != "FavoriteQuery" || RowIndex < 0) return false;
            DataGridViewRow row = SongQuery_DataGridView.Rows[RowIndex];
            return YoutubeFavorites.IsFavoriteRow(row.Cells["Song_Lang"].Value, row.Cells["Song_Path"].Value);
        }

        private void SongQuery_YoutubeFavoriteShowContextMenu(int RowIndex)
        {
            object VideoId = SongQuery_DataGridView.Rows[RowIndex].Cells["Song_Id"].Value;
            ContextMenuStrip GridView_ContextMenu = new ContextMenuStrip();
            GridView_ContextMenu.Items.Add("在瀏覽器開啟", null, (s, a) => SongQuery_YoutubeOpenInBrowser(VideoId));
            GridView_ContextMenu.Items.Add("複製影片網址", null, (s, a) => SongQuery_YoutubeCopyUrl(VideoId));
            ToolStripMenuItem RemoveItem = new ToolStripMenuItem("從我的最愛移除");
            RemoveItem.Click += new EventHandler(SongQuery_DataGridView_ContextMenuItem_RightClick);
            GridView_ContextMenu.Items.Add(RemoveItem);
            GridView_ContextMenu.Show(MousePosition.X, MousePosition.Y);
        }

        #endregion

        #endregion

        #region --- SongQuery YouTube 檢查影片 ---

        /// <summary>The column 檢查影片 adds to the YouTube list (not a YoutubeSongList column).</summary>
        private const string SongQuery_YoutubeCheckColumn = "CheckResult";

        /// <summary>Set while 檢查影片 runs; 取消檢查 and leaving the YouTube list cancel it.</summary>
        private CancellationTokenSource SongQuery_YoutubeCheckCancel;

        /// <summary>The room between the status text and 編輯模式 when 檢查影片 is hidden, measured the first time.</summary>
        private int SongQuery_StatusLabelGap = -1;

        /// <summary>檢查影片 shows only with the YouTube list; the status text makes room for it.</summary>
        private void SongQuery_YoutubeShowCheckButton(bool show)
        {
            if (SongQuery_StatusLabelGap < 0 && !SongQuery_YoutubeCheck_Button.Visible)
            {
                SongQuery_StatusLabelGap = Math.Max(0, SongQuery_EditMode_CheckBox.Left - SongQuery_QueryStatus_Label.Right);
            }
            SongQuery_YoutubeCheck_Button.Visible = show;
            int gap = Math.Max(0, SongQuery_StatusLabelGap);
            int right = (show ? SongQuery_YoutubeCheck_Button.Left : SongQuery_EditMode_CheckBox.Left) - gap;
            SongQuery_QueryStatus_Label.Width = Math.Max(0, right - SongQuery_QueryStatus_Label.Left);
            if (!show)
            {
                SongQuery_YoutubeCheck_Button.Text = "檢查影片";
                SongQuery_YoutubeCheck_Button.Enabled = true;
            }
        }

        /// <summary>
        /// 檢查影片: asks yt-dlp about every YouTube 歌曲 on show, one at a time, and writes the answer in 檢查結果. Then
        /// offers to remove those whose video is gone and whose file is not in the 下載目錄 (see YoutubeCheckResult.Removable).
        /// Pressed again while it runs (取消檢查), it stops after the video being checked.
        /// </summary>
        private void SongQuery_YoutubeCheck_Button_Click(object sender, EventArgs e)
        {
            if (SongQuery_YoutubeCheckCancel != null)
            {
                SongQuery_YoutubeCheckCancel.Cancel();
                SongQuery_YoutubeCheck_Button.Enabled = false;
                SongQuery_QueryStatus_Label.Text = "正在取消檢查,請稍待...";
                return;
            }
            if (!SongQuery_IsYoutubeMode || SongQuery_YoutubeTable.Rows.Count == 0) return;

            string AppDir = Application.StartupPath;
            string YtDlp = YoutubeAvailability.FindYtDlp(AppDir);
            if (YtDlp == null)
            {
                SongQuery_QueryStatus_Label.Text = @"找不到 yt-dlp (Tools\yt-dlp.exe),請將加歌程式和 UltimateKtv 放在同一個資料夾...";
                return;
            }
            string DownloadDir = YoutubeAvailability.DownloadDir(AppDir, out bool FromSettings);

            if (SongQuery_EditMode_CheckBox.Checked) SongQuery_EditMode_CheckBox.Checked = false;
            DataTable Table = SongQuery_YoutubeTable;
            if (!Table.Columns.Contains(SongQuery_YoutubeCheckColumn))
            {
                Table.Columns.Add(SongQuery_YoutubeCheckColumn, typeof(string));
                DataGridViewColumn column = SongQuery_DataGridView.Columns[SongQuery_YoutubeCheckColumn];
                column.HeaderText = "檢查結果";
                column.ReadOnly = true;
                column.Width = Convert.ToInt32(220 * Global.DPIScalingFactor * MainUIScale.UIScalingFactor);
                column.DisplayIndex = SongQuery_DataGridView.Columns["Name"].DisplayIndex + 1; // in sight, next to 名稱
            }

            // In the order on show; the values are read here, on the UI thread, which alone writes the table.
            var Songs = Table.DefaultView.Cast<DataRowView>()
                .Select(v => (Row: v.Row, Name: Convert.ToString(v.Row["Name"]), VideoId: v.Row["Url"], FileName: v.Row["FileName"]))
                .ToList();
            foreach (var song in Songs) song.Row[SongQuery_YoutubeCheckColumn] = "";

            var Cancel = new CancellationTokenSource();
            SongQuery_YoutubeCheckCancel = Cancel;
            SongQuery_YoutubeCheck_Button.Text = "取消檢查";
            SongQuery_EditMode_CheckBox.Enabled = false;
            SongQuery_Query_Button.Enabled = false;
            string DirText = DownloadDir + (FromSettings ? "" : " (預設)");
            SongQuery_QueryStatus_Label.Text = "正在檢查 " + Songs.Count + " 筆 YouTube 影片,下載目錄: " + DirText + "...";

            Task.Run(() =>
            {
                var Checked = new YoutubeCheckResult?[Songs.Count];
                int Done = 0;
                var Items = Songs.Select(song => (song.VideoId, YoutubeAvailability.HasLocalFile(DownloadDir, song.FileName))).ToList();
                YoutubeAvailability.CheckAll(Items, YoutubeAvailability.OEmbedStatus, YtDlp, TimeSpan.FromSeconds(60), Cancel.Token, (i, result) =>
                {
                    Checked[i] = result;
                    int done = Interlocked.Increment(ref Done);
                    var song = Songs[i];
                    SongQuery_YoutubeCheckInvoke(() =>
                    {
                        if (song.Row.RowState != DataRowState.Detached) song.Row[SongQuery_YoutubeCheckColumn] = result.Text;
                        if (SongQuery_YoutubeTable == Table && !Cancel.IsCancellationRequested) SongQuery_QueryStatus_Label.Text = "正在檢查 YouTube 影片 " + done + "/" + Songs.Count + ": " + song.Name + "...";
                    });
                });
                var Results = Enumerable.Range(0, Songs.Count).Where(i => Checked[i] != null)
                    .Select(i => (Songs[i].Row, Songs[i].Name, Checked[i].Value)).ToList();
                SongQuery_YoutubeCheckInvoke(() => SongQuery_YoutubeCheckFinished(Table, Cancel, Results, Songs.Count, DirText));
            });
        }

        /// <summary>Runs <paramref name="action"/> on the UI thread, unless SongMgr is closing.</summary>
        private void SongQuery_YoutubeCheckInvoke(Action action)
        {
            try
            {
                if (!IsDisposed && IsHandleCreated) BeginInvoke(action);
            }
            catch (InvalidOperationException) { }
        }

        private void SongQuery_YoutubeCheckFinished(DataTable Table, CancellationTokenSource Cancel, List<(DataRow Row, string Name, YoutubeCheckResult Result)> Results, int Total, string DirText)
        {
            bool Cancelled = Cancel.IsCancellationRequested;
            if (SongQuery_YoutubeCheckCancel == Cancel) SongQuery_YoutubeCheckCancel = null;
            Cancel.Dispose();
            SongQuery_Query_Button.Enabled = true;
            if (SongQuery_YoutubeTable != Table) return; // the list was left while it ran

            SongQuery_YoutubeCheck_Button.Text = "檢查影片";
            SongQuery_YoutubeCheck_Button.Enabled = true;
            SongQuery_EditMode_CheckBox.Enabled = Table.Rows.Count > 0;

            int Available = Results.Count(r => r.Result.State == YoutubeVideoState.Available);
            int LocalOnly = Results.Count(r => r.Result.State == YoutubeVideoState.Unavailable && r.Result.HasLocalFile);
            var Removable = Results.Where(r => r.Result.Removable && r.Row.RowState != DataRowState.Detached).ToList();
            int Unknown = Results.Count(r => r.Result.State == YoutubeVideoState.Unknown);
            // Short enough to fit beside 檢查影片: the gone-without-a-file count says what became of them.
            string Head = Cancelled ? "已取消 (查了 " + Results.Count + "/" + Total + " 筆)" : "檢查完成";
            string SummaryOf(string gone) => Head + ": 可用 " + Available + " 筆,失效但本機可播 " + LocalOnly + " 筆," + gone + ",無法確認 " + Unknown + " 筆。";
            SongQuery_QueryStatus_Label.Text = SummaryOf("失效無檔 " + Removable.Count + " 筆");
            if (Removable.Count == 0) return;

            const int Shown = 10;
            string Names = string.Join(Environment.NewLine, Removable.Take(Shown).Select(r => "・" + r.Name));
            if (Removable.Count > Shown) Names += Environment.NewLine + "...等 " + Removable.Count + " 筆";
            string Message = "有 " + Removable.Count + " 筆 YouTube 影片已失效,下載目錄裡也沒有影片檔:" + Environment.NewLine + Environment.NewLine + Names + Environment.NewLine + Environment.NewLine
                + "要移除這些 YouTube 歌曲資料嗎?" + Environment.NewLine + "(已失效但本機可播的不會移除;下載目錄: " + DirText + ")";
            if (MessageBox.Show(Message, "移除提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                SongQuery_QueryStatus_Label.Text = SummaryOf("失效無檔未移除 " + Removable.Count + " 筆");
                return;
            }

            try
            {
                int Count = YoutubeSongStore.Delete(Global.CrazyktvDatabaseFile, Removable.Select(r => r.Row["Id"].ToString()).ToList());
                foreach (var r in Removable) Table.Rows.Remove(r.Row);
                SongQuery_QueryStatus_Label.Text = SummaryOf("失效無檔已移除 " + Count + " 筆");
                if (Table.Rows.Count == 0)
                {
                    SongQuery_EditMode_CheckBox.Enabled = false;
                    SongQuery_YoutubeCheck_Button.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                SongQuery_QueryStatus_Label.Text = "移除已失效的 YouTube 歌曲資料失敗: " + ex.Message;
            }
        }

        #endregion

        #region --- SongQuery YouTube 編輯 ---

        private void SongQuery_YoutubeEditModeChanged()
        {
            if (SongQuery_EditMode_CheckBox.Checked)
            {
                SongQuery_DataGridView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                SongQuery_QueryStatus_Label.Text = "已進入編輯模式,在儲存格上按兩下、F2 或直接輸入即可修改,按 Delete 刪除所選的列...";
            }
            else
            {
                if (SongQuery_DataGridView.IsCurrentCellInEditMode) SongQuery_DataGridView.EndEdit();
                SongQuery_DataGridView.EditMode = DataGridViewEditMode.EditProgrammatically;
                SongQuery_EditMode_CheckBox.Enabled = SongQuery_DataGridView.RowCount > 0;
                SongQuery_QueryStatus_Label.Text = "已進入檢視模式...";
            }
            SongQuery_DataGridView.Focus();
        }

        /// <summary>The YoutubeSongList column a grid column saves to, or null when it is read only.</summary>
        private YoutubeColumn? SongQuery_YoutubeEditableColumn(int ColumnIndex)
        {
            if (ColumnIndex < 0) return null;
            string name = SongQuery_DataGridView.Columns[ColumnIndex].Name;
            return SongQuery_YoutubeColumns.FirstOrDefault(c => c.Name == name).Editable;
        }

        private void SongQuery_DataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (!SongQuery_IsYoutubeMode || !SongQuery_DataGridView.IsCurrentCellDirty) return;
            // 聲道 is picked from a list, so only typed values need checking.
            if (SongQuery_DataGridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) return;

            YoutubeColumn? column = SongQuery_YoutubeEditableColumn(e.ColumnIndex);
            if (column == null) return;

            string error = YoutubeSongStore.Validate(column.Value, e.FormattedValue, out _);
            if (error != "")
            {
                e.Cancel = true;
                SongQuery_QueryStatus_Label.Text = error + ",請重新輸入或按 Esc 取消...";
            }
        }

        private void SongQuery_DataGridView_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (!SongQuery_IsYoutubeMode || SongQuery_DataGridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) return;

            YoutubeColumn? column = SongQuery_YoutubeEditableColumn(e.ColumnIndex);
            if (column != null && YoutubeSongStore.Validate(column.Value, e.Value, out object value) == "")
            {
                e.Value = value;
                e.ParsingApplied = true;
            }
        }

        private void SongQuery_DataGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            // A 聲道 pick is saved at once instead of when the cell is left.
            if (SongQuery_IsYoutubeMode && SongQuery_DataGridView.IsCurrentCellDirty && SongQuery_DataGridView.CurrentCell is DataGridViewComboBoxCell)
            {
                SongQuery_DataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void SongQuery_DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!SongQuery_IsYoutubeMode || e.RowIndex < 0) return;

            YoutubeColumn? column = SongQuery_YoutubeEditableColumn(e.ColumnIndex);
            if (column == null) return;

            DataGridViewRow row = SongQuery_DataGridView.Rows[e.RowIndex];
            string id = row.Cells["Id"].Value.ToString();
            string header = SongQuery_DataGridView.Columns[e.ColumnIndex].HeaderText;
            try
            {
                if (YoutubeSongStore.Update(Global.CrazyktvDatabaseFile, id, column.Value, row.Cells[e.ColumnIndex].Value) == 0) throw new InvalidOperationException("找不到此筆資料");
                SongQuery_QueryStatus_Label.Text = "已更新 " + id + " 的" + header + "。";
            }
            catch (Exception ex)
            {
                // Show the database as it is rather than the value that failed to save.
                string keyword = SongQuery_QueryValue_TextBox.Text;
                this.BeginInvoke((Action)delegate()
                {
                    SongQuery_YoutubeQuery(keyword);
                    SongQuery_QueryStatus_Label.Text = "更新 " + id + " 的" + header + "失敗: " + ex.Message;
                });
            }
        }

        private void SongQuery_DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // A stored value missing from the 聲道 list, or text the grid cannot parse: no dialog.
            if (SongQuery_IsYoutubeMode) e.ThrowException = false;
        }

        private void SongQuery_YoutubeDeleteSelected()
        {
            List<DataGridViewRow> rows = SongQuery_DataGridView.SelectedRows.Cast<DataGridViewRow>().ToList();
            if (rows.Count == 0) return;

            if (MessageBox.Show("你確定要刪除所選的 " + rows.Count + " 筆 YouTube 歌曲資料嗎?" + Environment.NewLine + "(影片檔案不會刪除)", "刪除提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            try
            {
                int count = YoutubeSongStore.Delete(Global.CrazyktvDatabaseFile, rows.Select(row => row.Cells["Id"].Value.ToString()).ToList());
                foreach (DataGridViewRow row in rows) SongQuery_YoutubeTable.Rows.Remove(((DataRowView)row.DataBoundItem).Row);
                SongQuery_QueryStatus_Label.Text = "已刪除 " + count + " 筆 YouTube 歌曲資料。";

                if (SongQuery_DataGridView.RowCount == 0)
                {
                    SongQuery_EditMode_CheckBox.Checked = false;
                    SongQuery_EditMode_CheckBox.Enabled = false;
                    SongQuery_QueryStatus_Label.Text = "已刪除 " + count + " 筆 YouTube 歌曲資料,已無歌曲可編輯!";
                }
            }
            catch (Exception ex)
            {
                SongQuery_QueryStatus_Label.Text = "刪除 YouTube 歌曲資料失敗: " + ex.Message;
            }
        }

        #endregion
    }
}

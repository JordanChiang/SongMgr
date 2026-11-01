using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
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
                DataTable TrackList = new DataTable();
                TrackList.Columns.Add("Display", typeof(string));
                TrackList.Columns.Add("Value", typeof(int));
                foreach (var track in YoutubeSongStore.Tracks) TrackList.Rows.Add(track.Display, track.Value);
                combo.DataSource = TrackList;
                combo.DisplayMember = "Display";
                combo.ValueMember = "Value";
            });

            for (int i = 0; i < SongQuery_YoutubeColumns.Length; i++)
            {
                var spec = SongQuery_YoutubeColumns[i];
                DataGridViewColumn column = SongQuery_DataGridView.Columns[spec.Name];
                column.HeaderText = spec.Header;
                column.Width = Convert.ToInt32(spec.Width * Global.DPIScalingFactor * MainUIScale.UIScalingFactor);
                column.ReadOnly = spec.Editable == null;
                column.DisplayIndex = i;
            }
            ((DataGridViewTextBoxColumn)SongQuery_DataGridView.Columns["Name"]).MaxInputLength = 255;
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

            int RowIndex = e.RowIndex;
            ContextMenuStrip GridView_ContextMenu = new ContextMenuStrip();
            GridView_ContextMenu.Items.Add("在瀏覽器開啟", null, (s, a) => SongQuery_YoutubeOpenInBrowser(RowIndex));
            GridView_ContextMenu.Items.Add("複製影片網址", null, (s, a) =>
            {
                string url = YoutubeSongStore.WatchUrl(SongQuery_DataGridView.Rows[RowIndex].Cells["Url"].Value);
                if (url == "")
                {
                    SongQuery_QueryStatus_Label.Text = "此筆資料沒有影片 ID...";
                    return;
                }
                Clipboard.SetText(url);
                SongQuery_QueryStatus_Label.Text = "已複製影片網址: " + url;
            });
            GridView_ContextMenu.Show(MousePosition.X, MousePosition.Y);
        }

        private void SongQuery_YoutubeOpenInBrowser(int RowIndex)
        {
            string url = YoutubeSongStore.WatchUrl(SongQuery_DataGridView.Rows[RowIndex].Cells["Url"].Value);
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

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    public partial class MainForm : Form
    {
        #region --- SongAdd_DataGridView ---

        private void SongAdd_DataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int val;
            string valStr;
            switch (SongAdd_DataGridView.Columns[e.ColumnIndex].Name)
            {
                case "Song_AddStatus":
                    e.CellStyle.ForeColor = Color.Red;
                    break;
                case "Song_SingerType":
                    val = Convert.ToInt32(e.Value);
                    valStr = CommonFunc.GetSingerTypeStr(val, 1, "null");
                    e.Value = valStr;
                    e.FormattingApplied = true;
                    break;
                case "Song_Track":
                    val = Convert.ToInt32(e.Value);
                    valStr = CommonFunc.GetSongTrackStr(val, 0, "null");
                    e.Value = valStr;
                    e.FormattingApplied = true;
                    break;
                case "Song_CreatDate":
                    e.Value = DateTime.Parse(e.Value.ToString()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
                    e.FormattingApplied = true;
                    break;
                case "Song_WordCount":
                case "Song_MB":
                case "Song_SrcPath":
                    e.CellStyle.ForeColor = Color.Blue;
                    break;
            }
        }

        private void SongAdd_DataGridView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((e.RowIndex < 0) || (e.ColumnIndex < 0)) return;

            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip GridView_ContextMenu;
                ToolStripMenuItem[] GridView_ContextMenuItem;
                string valStr = "";

                if (SongAdd_Save_Button.Text == "取消加入" || SongAdd_Save_Button.Text == "取消更新")
                {
                    if (!SongAdd_DataGridView.Rows[e.RowIndex].Selected) SongAdd_DataGridView.CurrentCell = SongAdd_DataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];

                    if (SongAdd_DataGridView.SelectedRows.Count > 1)
                    {
                        GridView_ContextMenu = new ContextMenuStrip();
                        int ContextMenuCount = Convert.ToInt32(SongQuery.GetContextMenuStr(0, 0, true));
                        GridView_ContextMenuItem = new ToolStripMenuItem[ContextMenuCount];

                        for (int i = 0; i < ContextMenuCount; i++)
                        {
                            valStr = SongQuery.GetContextMenuStr(i, 0, false);
                            GridView_ContextMenuItem[i] = new ToolStripMenuItem(valStr);
                            GridView_ContextMenu.Items.Add(GridView_ContextMenuItem[i]);
                            GridView_ContextMenuItem[i].Click += new EventHandler(SongAdd_DataGridView_ContextMenuItem_RightClick);
                        }
                        GridView_ContextMenu.Show(MousePosition.X, MousePosition.Y);
                    }
                    else
                    {
                        GridView_ContextMenu = new ContextMenuStrip();
                        int ContextMenuCount = Convert.ToInt32(SongQuery.GetContextMenuStr(0, 1, true));
                        GridView_ContextMenuItem = new ToolStripMenuItem[ContextMenuCount];

                        for (int i = 0; i < ContextMenuCount; i++)
                        {
                            valStr = SongQuery.GetContextMenuStr(i, 1, false);
                            GridView_ContextMenuItem[i] = new ToolStripMenuItem(valStr);
                            GridView_ContextMenu.Items.Add(GridView_ContextMenuItem[i]);
                            GridView_ContextMenuItem[i].Click += new EventHandler(SongAdd_DataGridView_ContextMenuItem_RightClick);
                        }
                        GridView_ContextMenu.Show(MousePosition.X, MousePosition.Y);
                    }
                }
            }
        }

        private void SongAdd_DataGridView_ContextMenuItem_RightClick(object sender, EventArgs e)
        {
            string file = SongAdd_DataGridView.Rows[SongAdd_DataGridView.CurrentCell.RowIndex].Cells["Song_SrcPath"].Value.ToString();
            string SongPath = Path.GetDirectoryName(file);
            string SongFileName = Path.GetFileName(file);

            switch (sender.ToString())
            {
                case "開啟資料夾":
                    if (!Directory.Exists(SongPath) || !File.Exists(file))
                    {
                        SongAdd_Tooltip_Label.Text = "選取歌曲的資料夾或檔案不存在...";
                    }
                    else
                    {
                        string arg = "/select, " + file;
                        Process.Start("explorer", arg);
                    }
                    break;
                case "播放檔案":
                    if (!File.Exists(file))
                    {
                        SongAdd_Tooltip_Label.Text = "【" + SongFileName + "】檔案不存在...";
                    }
                    else
                    {
                        int i = SongAdd_DataGridView.CurrentCell.RowIndex;

                        PlayerRequest request = new PlayerRequest
                        {
                            Song = SongRecord.ForPlayer(SongAdd_DataGridView.Rows[i]),
                            FilePath = file,
                            RowIndex = i.ToString(),
                            Source = "SongAdd",
                        };

                        Global.PlayerTrackUpdate = null;
                        if (Global.MainCfgPlayerCore == "1")
                        {
                            DShowForm newPlayerForm = new DShowForm(this, request);
                            newPlayerForm.Show();
                        }
                        this.Hide();
                    }
                    break;
                case "刪除資料列":
                    if (MessageBox.Show("你確定要刪除資料嗎?", "刪除提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        foreach (DataGridViewRow row in SongAdd_DataGridView.SelectedRows)
                        {
                            SongAdd_DataGridView.Rows.Remove(row);
                        }

                        if (SongAdd_DataGridView.RowCount == 0)
                        {
                            SongAdd_DataGridView.DataSource = null;
                            SongAdd_Add_Button.Enabled = false;
                            SongAdd_InitializeEditControl();
                            SongAdd_Tooltip_Label.Text = "已無歌曲可加入!";
                        }
                    }
                    break;
            }
        }

        private void SongAdd_DataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1)
            {
                Global.SongAddDataGridViewRestoreCurrentRow = SongAdd_DataGridView.CurrentRow.Cells["Song_SrcPath"].Value.ToString();
                Global.SongAddDataGridViewRestoreSelectList = new List<string>();
                foreach (DataGridViewRow row in SongAdd_DataGridView.SelectedRows)
                {
                    string SongPath = row.Cells["Song_SrcPath"].Value.ToString();
                    Global.SongAddDataGridViewRestoreSelectList.Add(SongPath);
                }
            }
        }

        private void SongAdd_DataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (SongAdd_Save_Button.Text == "取消加入" || SongAdd_Save_Button.Text == "取消更新")
            {
                if (SongAdd_DataGridView.SelectedRows.Count > 0)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.Delete:
                            if (MessageBox.Show("你確定要刪除資料嗎?", "刪除提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                foreach (DataGridViewRow row in SongAdd_DataGridView.SelectedRows)
                                {
                                    SongAdd_DataGridView.Rows.Remove(row);
                                }

                                if (SongAdd_DataGridView.RowCount == 0)
                                {
                                    SongAdd_DataGridView.DataSource = null;
                                    SongAdd_Add_Button.Enabled = false;
                                    SongAdd_InitializeEditControl();
                                    SongAdd_Tooltip_Label.Text = "已無歌曲可加入!";
                                }
                            }
                            break;
                    }
                }
            }
        }

        private void SongAdd_DataGridView_SelectionChanged(object sender, EventArgs e)
        {
            if (SongAdd_DataGridView.SelectedRows.Count > 0)
            {
                int SelectedRowsCount = SongAdd_DataGridView.SelectedRows.Count;
                List<string> SongSongTypeList = new List<string>(Global.SongMgrSongType.Split(','));

                if (SongAnalysis.SongAnalysisCompleted)
                {
                    SongAnalysis.SongAnalysisCompleted = false;
                }
                else
                {
                    if (SongAdd_Tooltip_Label.Text != "") SongAdd_Tooltip_Label.Text = "";
                }

                Global.SongAddDataGridViewSelectList.Clear();

                if (SelectedRowsCount > 1)
                {
                    Global.SongAddMultiEditUpdateList = new List<bool>() { false, false, false, false, false, false, false, false };

                    if (!Global.SongAddMultiEdit)
                    {
                        SongAdd_Edit_GroupBox.Text = "批次編輯";
                        SongAdd_GetSongEditComboBoxList(true);

                        SongAdd_EditSongId_TextBox.Enabled = false;
                        SongAdd_EditSongLang_ComboBox.Enabled = true;
                        SongAdd_EditSongCreatDate_DateTimePicker.Enabled = true;
                        SongAdd_EditSongSinger_TextBox.Enabled = true;
                        SongAdd_EditSongSingerType_ComboBox.Enabled = true;
                        SongAdd_EditSongSongName_TextBox.Enabled = false;
                        SongAdd_EditSongSongType_ComboBox.Enabled = true;
                        SongAdd_EditSongSpell_TextBox.Enabled = false;
                        SongAdd_EditSongWordCount_TextBox.Enabled = false;
                        SongAdd_EditSongSrcPath_TextBox.Enabled = false;
                        SongAdd_EditSongTrack_ComboBox.Enabled = true;
                        SongAdd_EditSongTrack_Button.Enabled = false;
                        SongAdd_EditSongVolume_TextBox.Enabled = true;
                        SongAdd_EditSongPlayCount_TextBox.Enabled = true;
                        SongAdd_EditApplyChanges_Button.Enabled = true;

                        SongAdd_EditSongId_TextBox.Text = "";
                        SongAdd_EditSongLang_ComboBox.SelectedValue = 1;
                        SongAdd_EditSongCreatDate_DateTimePicker.Value = DateTime.Now;
                        SongAdd_EditSongSinger_TextBox.Text = "";
                        SongAdd_EditSongSingerType_ComboBox.SelectedValue = 1;
                        SongAdd_EditSongSongName_TextBox.Text = "";
                        SongAdd_EditSongSongType_ComboBox.SelectedValue = 1;
                        SongAdd_EditSongSpell_TextBox.Text = "";
                        SongAdd_EditSongWordCount_TextBox.Text = "";
                        SongAdd_EditSongSrcPath_TextBox.Text = "";
                        SongAdd_EditSongTrack_ComboBox.SelectedValue = 1;
                        SongAdd_EditSongVolume_TextBox.Text = "";
                        SongAdd_EditSongPlayCount_TextBox.Text = "";
                    }
                    
                    Global.SongAddDataGridViewSelectList = new List<SongRecord>();

                    foreach (DataGridViewRow row in SongAdd_DataGridView.SelectedRows)
                    {
                        Global.SongAddDataGridViewSelectList.Add(SongRecord.FromAddGridRow(row));
                    }
                }
                else if (SelectedRowsCount == 1)
                {
                    SongAdd_Edit_GroupBox.Text = "單曲編輯";
                    SongAdd_GetSongEditComboBoxList(false);

                    SongAdd_EditSongId_TextBox.Enabled = false;
                    SongAdd_EditSongLang_ComboBox.Enabled = true;
                    SongAdd_EditSongCreatDate_DateTimePicker.Enabled = true;
                    SongAdd_EditSongSinger_TextBox.Enabled = true;
                    SongAdd_EditSongSingerType_ComboBox.Enabled = true;
                    SongAdd_EditSongSongName_TextBox.Enabled = true;
                    SongAdd_EditSongSongType_ComboBox.Enabled = true;
                    SongAdd_EditSongSpell_TextBox.Enabled = true;
                    SongAdd_EditSongWordCount_TextBox.Enabled = true;
                    SongAdd_EditSongSrcPath_TextBox.Enabled = true;
                    SongAdd_EditSongTrack_ComboBox.Enabled = true;
                    SongAdd_EditSongTrack_Button.Enabled = true;
                    SongAdd_EditSongVolume_TextBox.Enabled = true;
                    SongAdd_EditSongPlayCount_TextBox.Enabled = true;
                    SongAdd_EditApplyChanges_Button.Enabled = true;

                    SongRecord song = SongRecord.FromAddGridRow(SongAdd_DataGridView.SelectedRows[0]);

                    SongAdd_EditSongId_TextBox.Text = song.Id;
                    SongAdd_EditSongLang_ComboBox.SelectedValue = (song.Lang == "未知") ? 11 : Global.CrazyktvSongLangList.IndexOf(song.Lang) + 1;
                    SongAdd_EditSongCreatDate_DateTimePicker.Value = DateTime.Parse(song.CreatDate);
                    SongAdd_EditSongSinger_TextBox.Text = song.Singer;
                    string SongSingerTypeStr = CommonFunc.GetSingerTypeStr(song.SingerTypeValue, 1, "null");
                    SongAdd_EditSongSingerType_ComboBox.SelectedValue = Convert.ToInt32(CommonFunc.GetSingerTypeStr(0, 3, SongSingerTypeStr)) + 1;
                    SongAdd_EditSongSongName_TextBox.Text = song.SongName;
                    SongAdd_EditSongSongType_ComboBox.SelectedValue = (song.SongType == "") ? 1 : SongSongTypeList.IndexOf(song.SongType) + 2;
                    SongAdd_EditSongSpell_TextBox.Text = song.Spell;
                    SongAdd_EditSongWordCount_TextBox.Text = song.WordCount;
                    SongAdd_EditSongSrcPath_TextBox.Text = song.SrcPath;
                    SongAdd_EditSongTrack_ComboBox.SelectedValue = song.TrackValue;
                    SongAdd_EditSongVolume_TextBox.Text = song.Volume;
                    SongAdd_EditSongPlayCount_TextBox.Text = song.PlayCount;

                    Global.SongAddDataGridViewSelectList = new List<SongRecord>() { song };
                }
            }
        }

        private void SongAdd_DataGridView_Sorted(object sender, EventArgs e)
        {
            if (Global.SongAddDataGridViewRestoreSelectList.Count > 0)
            {
                SongAdd_DataGridView.ClearSelection();

                if (Global.SongAddDataGridViewRestoreCurrentRow != "")
                {
                    var query = from row in SongAdd_DataGridView.Rows.Cast<DataGridViewRow>()
                                where row.Cells["Song_SrcPath"].Value.Equals(Global.SongAddDataGridViewRestoreCurrentRow)
                                select row;

                    if (query.Count() > 0)
                    {
                        foreach (DataGridViewRow row in query)
                        {
                            SongAdd_DataGridView.CurrentCell = row.Cells[0];
                        }
                    }
                }

                foreach (string str in Global.SongAddDataGridViewRestoreSelectList)
                {
                    var query = from row in SongAdd_DataGridView.Rows.Cast<DataGridViewRow>()
                                where row.Cells["Song_SrcPath"].Value.Equals(str)
                                select row;

                    foreach (DataGridViewRow row in query)
                    {
                        row.Selected = true;
                    }
                }
                Global.SongAddDataGridViewRestoreSelectList.Clear();
            }
        }

        #endregion

    }
}

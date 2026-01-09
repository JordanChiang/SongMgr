using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Forms;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CrazyKTV_MediaKit.DirectShow.Controls;
using CrazyKTV_MediaKit.DirectShow.MediaPlayers;
using System.Data.OleDb;

namespace CrazyKTV_SongMgr
{
    public partial class DShowForm : Form
    {
        private string SongId;
        private string SongLang;
        private string SongSinger;
        private string SongSongName;
        private string SongTrack;
        private string SongVolume;
        private string SongReplayGain;
        private string SongFilePath;
        private string dvRowIndex;
        private string UpdateSongTrack;
        private string UpdateDataGridView;

        private MediaUriElement mediaUriElement;
        private System.Timers.Timer mouseClickTimer;
        private DateTime MediaPositionChangeTime;
        private bool sliderInit;
        private bool sliderDrag;

        /// <summary>
        /// The index of the monitor to use for fullscreen mode (0 = primary).
        /// </summary>
        public int FullScreenMonitorIndex { get; set; } = 0;

        /// <summary>
        /// If true, the form will attempt to go fullscreen when it loads.
        /// </summary>
        public bool StartInFullScreen { get; set; } = false;

        /// <summary>
        /// Gets or sets a value indicating whether hotkeys (ESC, F11, etc.) are enabled.
        /// </summary>
        public bool HotkeysEnabled { get; set; } = true;

        public DShowForm()
        {
            InitializeComponent();
            this.Load += DShowForm_Load;
            this.KeyPreview = true;
            this.KeyDown += DShowForm_KeyDown;
            SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;
        }

        public DShowForm(Form ParentForm, List<string> PlayerSongInfoList)
        {
            InitializeComponent();
            this.Load += DShowForm_Load;
            this.MouseWheel += DShowForm_MouseWheel;
            this.KeyPreview = true;
            this.KeyDown += DShowForm_KeyDown;
            SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;

            this.Owner = ParentForm;
            SongId = PlayerSongInfoList[0];
            SongLang = PlayerSongInfoList[1];
            SongSinger = PlayerSongInfoList[2];
            SongSongName = PlayerSongInfoList[3];
            SongTrack = PlayerSongInfoList[4];
            SongVolume = PlayerSongInfoList[5];
            SongReplayGain = PlayerSongInfoList[6];
            SongFilePath = PlayerSongInfoList[7];
            dvRowIndex = PlayerSongInfoList[8];
            UpdateDataGridView = PlayerSongInfoList[9];

            this.Text = "【" + SongLang + "】" + SongSinger + " - " + SongSongName;

            sliderInit = false;

            mediaUriElement = new MediaUriElement();
            mediaUriElement.BeginInit();
            elementHost.Child = mediaUriElement;
            mediaUriElement.EndInit();

            mediaUriElement.MediaUriPlayer.CodecsDirectory = System.Windows.Forms.Application.StartupPath + @"\Codec";
            mediaUriElement.VideoRenderer = (Global.MainCfgPlayerOutput == "1") ? CrazyKTV_MediaKit.DirectShow.MediaPlayers.VideoRendererType.VideoMixingRenderer9 : CrazyKTV_MediaKit.DirectShow.MediaPlayers.VideoRendererType.EnhancedVideoRenderer;
            mediaUriElement.DeeperColor = (Global.MainCfgPlayerOutput == "1") ? false : true;
            mediaUriElement.Stretch = System.Windows.Media.Stretch.Fill;
            mediaUriElement.EnableAudioCompressor = bool.Parse(Global.MainCfgPlayerEnableAudioCompressor);
            mediaUriElement.EnableAudioProcessor = bool.Parse(Global.MainCfgPlayerEnableAudioProcessor);

            mediaUriElement.MediaFailed += MediaUriElement_MediaFailed;
            mediaUriElement.MediaEnded += MediaUriElement_MediaEnded;
            mediaUriElement.MouseLeftButtonDown += mediaUriElement_MouseLeftButtonDown;
            mediaUriElement.MediaUriPlayer.MediaPositionChanged += MediaUriPlayer_MediaPositionChanged;
            mediaUriElement.MediaOpened += mediaUriElement_MediaOpened;
            MediaPositionChangeTime = DateTime.Now;

            // 隨選視訊
            if (Global.PlayerRandomVideoList.Count == 0)
            {
                string dir = System.Windows.Forms.Application.StartupPath + @"\Video";
                if (Directory.Exists(dir))
                {
                    Global.PlayerRandomVideoList.AddRange(Directory.GetFiles(dir));
                    if (Global.PlayerRandomVideoList.Count > 0)
                    {
                        Random rand = new Random(Guid.NewGuid().GetHashCode());
                        Global.PlayerRandomVideoList = Global.PlayerRandomVideoList.OrderBy(str => rand.Next()).ToList<string>();
                    }
                }
            }
            mediaUriElement.VideoSource = (Global.PlayerRandomVideoList.Count > 0) ? new Uri(Global.PlayerRandomVideoList[0]) : null;
            mediaUriElement.Source = new Uri(SongFilePath);

            mediaUriElement.Volume = Math.Round(Convert.ToDouble(Global.MainCfgPlayerDefaultVolume) / 100, 2);
            // 音量平衡
            int GainVolume = Convert.ToInt32(SongVolume);
            if (!string.IsNullOrEmpty(SongReplayGain))
            {
                int basevolume = 100;
                GainVolume = basevolume;

                double GainDB = Convert.ToDouble(SongReplayGain);
                GainVolume = Convert.ToInt32(basevolume * Math.Pow(10, GainDB / 20));

            }
            mediaUriElement.AudioAmplify = GainVolume;
            Player_CurrentVolValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = SongVolume + " %");

            Player_VolumeSlider.TrackBarValue = Convert.ToInt32(SongVolume);
            Player_VolumeSlider.ProgressBarValue = Player_VolumeSlider.TrackBarValue;
            Player_VolumeSlider.TrackBarValueChanged += Player_VolumeSlider_ValueChanged;

            NativeMethods.SystemSleepManagement.PreventSleep(true);
        }

        private void mediaUriElement_MediaOpened(object sender, RoutedEventArgs e)
        {
            string ChannelValue = string.Empty;

            Player_CurrentChannelValue_Label.Text = string.Empty;
            if (mediaUriElement.AudioStreams.Count == 1)
            {
                switch (SongTrack)
                {
                    case "1":
                        if (mediaUriElement.AudioChannel != 1) 
                            mediaUriElement.AudioChannel = 1;
                        ChannelValue = "1";
                        break;
                    case "2":
                        if (mediaUriElement.AudioChannel != 2) 
                            mediaUriElement.AudioChannel = 2;
                        ChannelValue = "2";
                        break;
                }
                ChannelValue = SongTrack;
                mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
            }
            else if (mediaUriElement.AudioStreams.Count >= 3)
            {
                mediaUriElement.AudioTrack = Convert.ToInt32(SongTrack);
                ChannelValue = SongTrack;
                //                if (ChannelValue == SongTrack)
                //                    Player_CurrentChannelValue_Label.Text = "伴唱";
                //                else
                //                    Player_CurrentChannelValue_Label.Text = String.Format("音軌 ({0})", mediaUriElement.AudioTrack);

            }
            else if (mediaUriElement.AudioStreams.Count > 1)
            {
                switch (SongTrack)
                {
                    case "1":
                        if (Global.SongMgrSongTrackMode == "True")
                        {
                            if (mediaUriElement.AudioStreams.IndexOf(mediaUriElement.AudioTrack) != mediaUriElement.AudioStreams[0])
                                mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
                        }
                        else
                        {
                            if (mediaUriElement.AudioStreams.IndexOf(mediaUriElement.AudioTrack) != mediaUriElement.AudioStreams[1])
                                mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[1];
                        }
                        break;
                    case "2":
                        if (Global.SongMgrSongTrackMode == "True")
                        {
                            if (mediaUriElement.AudioStreams.IndexOf(mediaUriElement.AudioTrack) != mediaUriElement.AudioStreams[1])
                                mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[1];
                        }
                        else
                        {
                            if (mediaUriElement.AudioStreams.IndexOf(mediaUriElement.AudioTrack) != mediaUriElement.AudioStreams[0])
                                mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
                        }
                        break;
                }
                ChannelValue = SongTrack;
            }
            // no text needed if we are processed
            if (Player_CurrentChannelValue_Label.Text == "")
                Player_CurrentChannelValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = (ChannelValue == SongTrack) ? "伴唱" : "人聲");
         
            if (mediaUriElement.MediaUriPlayer.IsAudioOnly && Global.PlayerRandomVideoList.Count > 0)
                Global.PlayerRandomVideoList.RemoveAt(0);
        }

        private void DShowForm_Load(object sender, EventArgs e)
        {
            if (StartInFullScreen)
            {
                // Use BeginInvoke to ensure the form is fully loaded and visible before toggling.
                this.BeginInvoke(new Action(() =>
                {
                    ToggleFullscreen();
                }));
            }
        }

        private void DShowForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (!HotkeysEnabled)
                return;

            // F11 to toggle fullscreen should always work.
            if (e.KeyCode == Keys.F11)
            {
                ToggleFullscreen();
                return;
            }

            // Other hotkeys (like ESC to close) only work when in fullscreen mode.
            if (this.FormBorderStyle == FormBorderStyle.None)
            {
                switch (e.KeyCode)
                {
                    // Terminate (close) the player window with ESC or Q.
                    case Keys.Escape:
                    case Keys.Q:
                        this.Close();
                        break;
                }
            }
        }

        private void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
        {
            // If we are in fullscreen mode, we might need to move to a different monitor.
            if (this.FormBorderStyle == FormBorderStyle.None)
            {
                Screen[] screens = Screen.AllScreens;
                Screen targetScreen;

                // Check if the originally requested screen index is still valid.
                if (FullScreenMonitorIndex >= 0 && FullScreenMonitorIndex < screens.Length)
                {
                    targetScreen = screens[FullScreenMonitorIndex];
                }
                else
                {
                    // If the monitor index is now invalid (e.g., screen disconnected), default to the primary screen.
                    targetScreen = Screen.PrimaryScreen;
                }

                // If the form is not on the target screen, move it.
                if (Screen.FromControl(this).DeviceName != targetScreen.DeviceName)
                {
                    this.Bounds = targetScreen.Bounds;
                }
            }
            RecreateMediaElement();
        }
        private void RecreateMediaElement()
        {
            if (mediaUriElement == null)
            {
                return;
            }

            // 1. Store playback state and all necessary properties.
            bool wasPlaying = (mediaUriElement.MediaUriPlayer.PlayerState == PlayerState.Playing);
            long currentPosition = mediaUriElement.MediaPosition;
            Uri currentVideoSource = mediaUriElement.VideoSource;
            Uri currentSource = mediaUriElement.Source;
            double currentVolume = mediaUriElement.Volume;
            int currentAudioTrack = mediaUriElement.AudioTrack;
            int currentAudioChannel = mediaUriElement.AudioChannel;
            int gainVolume = mediaUriElement.AudioAmplify;

            // 2. Detach and dispose the old media element.
            elementHost.Child = null;
            mediaUriElement.Close();
            mediaUriElement = null;

            // 3. Create a new MediaUriElement and re-initialize it.
            mediaUriElement = new MediaUriElement();
            mediaUriElement.BeginInit();
            elementHost.Child = mediaUriElement;
            mediaUriElement.EndInit();

            // 4. Re-apply all the initial settings from the constructor.
            mediaUriElement.MediaUriPlayer.CodecsDirectory = System.Windows.Forms.Application.StartupPath + @"\Codec";
            mediaUriElement.VideoRenderer = (Global.MainCfgPlayerOutput == "1") ? CrazyKTV_MediaKit.DirectShow.MediaPlayers.VideoRendererType.VideoMixingRenderer9 : CrazyKTV_MediaKit.DirectShow.MediaPlayers.VideoRendererType.EnhancedVideoRenderer;
            mediaUriElement.DeeperColor = (Global.MainCfgPlayerOutput == "1") ? false : true;
            mediaUriElement.Stretch = System.Windows.Media.Stretch.Fill;
            mediaUriElement.EnableAudioCompressor = bool.Parse(Global.MainCfgPlayerEnableAudioCompressor);
            mediaUriElement.EnableAudioProcessor = bool.Parse(Global.MainCfgPlayerEnableAudioProcessor);

            mediaUriElement.MediaFailed += MediaUriElement_MediaFailed;
            mediaUriElement.MediaEnded += MediaUriElement_MediaEnded;
            mediaUriElement.MouseLeftButtonDown += mediaUriElement_MouseLeftButtonDown;
            mediaUriElement.MediaUriPlayer.MediaPositionChanged += MediaUriPlayer_MediaPositionChanged;

            // 5. Restore media sources and properties
            mediaUriElement.VideoSource = currentVideoSource;
            mediaUriElement.Source = currentSource;
            mediaUriElement.Volume = currentVolume;
            mediaUriElement.AudioAmplify = gainVolume;

            // 6. Wait for it to open, then restore audio track, seek, and play
            Task.Run(() =>
            {
                // Spin until the player is in the opened state.
                SpinWait.SpinUntil(() => mediaUriElement.MediaUriPlayer.PlayerState == PlayerState.Opened, 5000); // 5 sec timeout

                if (mediaUriElement.MediaUriPlayer.PlayerState == PlayerState.Opened)
                {
                    mediaUriElement.Dispatcher.Invoke(() =>
                    {
                        mediaUriElement.AudioTrack = currentAudioTrack;
                        mediaUriElement.AudioChannel = currentAudioChannel;
                        mediaUriElement.MediaPosition = currentPosition;

                        if (wasPlaying)
                        {
                            mediaUriElement.Play();
                        }
                    });
                }
            });
        }

        private void DShowForm_MouseWheel(object sender, MouseEventArgs e)
        {
            if (e.Delta != 0)
            {
                if (e.Delta > 0)
                {
                    if (mediaUriElement.Volume <= 0.95)
                    {
                        mediaUriElement.Volume += 0.05;
                    }
                    else
                    {
                        mediaUriElement.Volume = 1.00;
                    }
                }
                else
                {
                    if (mediaUriElement.Volume >= 0.05)
                    {
                        mediaUriElement.Volume -= 0.05;
                    }
                    else
                    {
                        mediaUriElement.Volume = 0;
                    }
                }
                mediaUriElement.Volume = Math.Round(mediaUriElement.Volume, 2);
                
                int volPos = (int)(mediaUriElement.Volume * 100);
                if (Player_VolumeSlider.TrackBarValue != volPos)
                {
                   Player_VolumeSlider.TrackBarValue = volPos;
                   Player_VolumeSlider.ProgressBarValue = volPos;
                }
            }
        }

        private void MediaUriElement_MediaFailed(object sender, CrazyKTV_MediaKit.DirectShow.MediaPlayers.MediaFailedEventArgs e)
        {
            this.BeginInvokeIfRequired(form => form.Text = e.Message);
        }

        private void MediaUriPlayer_MediaPositionChanged(object sender, EventArgs e)
        {
            if (sliderDrag)
                return;

            if (!sliderInit)
            {
                this.Invoke((Action)delegate ()
                {
                    if (mediaUriElement.HasVideo)
                    {
                        if (mediaUriElement.NaturalVideoWidth != 0 && mediaUriElement.NaturalVideoHeight != 0)
                        {
                            Player_VideoSizeValue_Label.Text = mediaUriElement.NaturalVideoWidth + "x" + mediaUriElement.NaturalVideoHeight;
                        }
                    }

                    if (mediaUriElement.MediaDuration > 0)
                    {
                        Player_ProgressTrackBar.Maximum = ((int)mediaUriElement.MediaDuration < 0) ? (int)mediaUriElement.MediaDuration * -1 : (int)mediaUriElement.MediaDuration;
                        sliderInit = true;
                    }
                });
            }
            else
            {
                if ((DateTime.Now - MediaPositionChangeTime).TotalMilliseconds < 500) return;
                this.BeginInvoke(new Action(ChangeSlideValue), null);
                MediaPositionChangeTime = DateTime.Now;
            }
        }

        private void ChangeSlideValue()
        {
            if (sliderDrag)
                return;

            if (sliderInit)
            {
                double perc = (double)mediaUriElement.MediaPosition / mediaUriElement.MediaDuration;
                int newValue = (int)(Player_ProgressTrackBar.Maximum * perc);
                if (newValue - Player_ProgressTrackBar.ProgressBarValue < 500000) return;
                Player_ProgressTrackBar.TrackBarValue = newValue;
                Player_ProgressTrackBar.ProgressBarValue = newValue;
            }
        }

        private void Player_ProgressTrackBar_Click(object sender, EventArgs e)
        {
            if (!sliderInit)
                return;

            ChangeMediaPosition();
        }

        private void ChangeMediaPosition()
        {
            sliderDrag = true;
            double perc = (double)Player_ProgressTrackBar.TrackBarValue / Player_ProgressTrackBar.Maximum;
            mediaUriElement.MediaPosition = (long)(mediaUriElement.MediaDuration * perc);
            Player_ProgressTrackBar.ProgressBarValue = Player_ProgressTrackBar.TrackBarValue;
            sliderDrag = false;
        }

        private void MediaUriElement_MediaEnded(object sender, RoutedEventArgs e)
        {
            mediaUriElement.Stop();
            mediaUriElement.MediaPosition = 0;
            Player_ProgressTrackBar.TrackBarValue = 0;
            Player_ProgressTrackBar.ProgressBarValue = 0;
            Player_PlayControl_Button.Text = "播放";
        }

        private void Player_UpdateVolume_Button_Click(object sender, EventArgs e)
        {
            try
            {
                using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Global.CrazyktvDatabaseFile))
                {
                    conn.Open();
                    string sql = "UPDATE ktv_Song SET Song_Volume = @Volume WHERE Song_Id = @SongId";
                    using (System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@Volume", Player_VolumeSlider.TrackBarValue.ToString());
                        cmd.Parameters.AddWithValue("@SongId", SongId);
                        cmd.ExecuteNonQuery();
                    }
                }
                
                SongVolume = Player_VolumeSlider.TrackBarValue.ToString();
                Player_UpdateVolume_Button.Enabled = false;

                // Sync to MainForm DataGridView
                if (this.Owner != null)
                {
                    System.Reflection.FieldInfo field = this.Owner.GetType().GetField("SongQuery_DataGridView", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    if (field != null)
                    {
                        System.Windows.Forms.DataGridView grid = (System.Windows.Forms.DataGridView)field.GetValue(this.Owner);
                        if (grid != null && int.TryParse(dvRowIndex, out int rowIndex) && rowIndex >= 0 && rowIndex < grid.Rows.Count)
                        {
                            if (grid.Columns.Contains("Song_Volume"))
                            {
                                grid.Rows[rowIndex].Cells["Song_Volume"].Value = SongVolume;
                            }
                        }
                    }
                }

                this.Text = "【" + SongLang + "】" + SongSinger + " - " + SongSongName + " - 音量更新成功!";
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("音量更新失敗: " + ex.Message, "錯誤", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        private void Player_SwithChannel_Button_Click(object sender, EventArgs e)
        {
            string ChannelValue;

            if (mediaUriElement.AudioStreams.Count >= 3)
            {
                // lookup audio track 1 ~ n
                if (mediaUriElement.AudioTrack < mediaUriElement.AudioStreams.Count)
                {
                    mediaUriElement.AudioTrack++;
                }
                else
                {
                    mediaUriElement.AudioTrack = 1;
                }

                UpdateSongTrack = Convert.ToString(mediaUriElement.AudioTrack);
                ChannelValue = UpdateSongTrack;
                if (ChannelValue == SongTrack)
                    Player_CurrentChannelValue_Label.Text = "伴唱";
                else
                    Player_CurrentChannelValue_Label.Text = String.Format("音軌 ({0})", UpdateSongTrack);
            }
            else if (mediaUriElement.AudioStreams.Count > 1)
            {
                // 2 tracks only, toggle between them
                if (mediaUriElement.AudioTrack == mediaUriElement.AudioStreams[0])
                {
                    mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[1];
                    ChannelValue = "2";
                }
                else//  (mediaUriElement.AudioTrack == mediaUriElement.AudioStreams[1])
                {
                    mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
                    ChannelValue = "1";
                }

                if (Global.SongMgrSongTrackMode == "True")
                {
                    ChannelValue = (ChannelValue == "2") ? "2" : "1";
                }
                else
                {
                    ChannelValue = (ChannelValue == "1") ? "2" : "1";
                }
                UpdateSongTrack = ChannelValue;
            }
            else
            {
                if (mediaUriElement.AudioChannel == 1)
                {
                    mediaUriElement.AudioChannel = 2;
                    ChannelValue = "2";
                    UpdateSongTrack = "2";
                }
                else
                {
                    mediaUriElement.AudioChannel = 1;
                    ChannelValue = "1";
                    UpdateSongTrack = "1";
                }
            }

            if (mediaUriElement.AudioStreams.Count <= 2)
            { 
 //               if (Player_CurrentChannelValue_Label.Text == "")
                    Player_CurrentChannelValue_Label.Text = (ChannelValue == SongTrack) ? "伴唱" : "人聲";
            }

            Player_UpdateChannel_Button.Enabled = (Player_CurrentChannelValue_Label.Text == "人聲") ? true : false;

            // always enable the button if multiple audio track
            if (mediaUriElement.AudioStreams.Count > 2)
                Player_UpdateChannel_Button.Enabled = true;
        }

        private void Player_UpdateChannel_Button_Click(object sender, EventArgs e)
        {
            SongTrack = UpdateSongTrack;
            Player_UpdateChannel_Button.Enabled = false;
            Player_CurrentChannelValue_Label.Text = "伴唱";
            Global.PlayerUpdateSongValueList = new List<string>() { UpdateDataGridView, dvRowIndex, SongTrack };
        }

        private void Player_PlayControl_Button_Click(object sender, EventArgs e)
        {
            switch (((Button)sender).Text)
            {
                case "暫停播放":
                    mediaUriElement.Pause();
                    ((Button)sender).Text = "繼續播放";
                    break;
                case "繼續播放":
                    mediaUriElement.Play();
                    ((Button)sender).Text = "暫停播放";
                    break;
                case "播放":
                    mediaUriElement.Play();
                    ((Button)sender).Text = "暫停播放";
                    break;
            }
        }

        private void mediaUriElement_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (mouseClickTimer == null)
            {
                mouseClickTimer = new System.Timers.Timer
                {
                    Interval = SystemInformation.DoubleClickTime
                };
                mouseClickTimer.Elapsed += new System.Timers.ElapsedEventHandler(mouseClickTimer_Tick);
            }

            if (!mouseClickTimer.Enabled)
            {
                mouseClickTimer.Start();
            }
            else
            {
                mouseClickTimer.Stop();
                ToggleFullscreen();
            }
        }

        private void mouseClickTimer_Tick(object sender, EventArgs e)
        {
            mouseClickTimer.Stop();

            switch (Player_PlayControl_Button.Text)
            {
                case "暫停播放":
                    mediaUriElement.Dispatcher.Invoke(new Action(() => mediaUriElement.Pause()));
                    Player_PlayControl_Button.InvokeIfRequired<Button>(btn => btn.Text = "繼續播放");
                    break;
                case "繼續播放":
                    mediaUriElement.Dispatcher.Invoke(new Action(() => mediaUriElement.Play()));
                    Player_PlayControl_Button.InvokeIfRequired<Button>(btn => btn.Text = "暫停播放");
                    break;
            }
        }

        private FormWindowState winState;
        private System.Drawing.Point winLoc;
        private int winWidth;
        private int winHeight;
        private int eHostWidth;
        private int eHostHeight;
        
        private void ToggleFullscreen()
        {
            if (this.FormBorderStyle == FormBorderStyle.None)
            {
                this.WindowState = winState;

                this.Hide();
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.TopMost = false;
                this.Bounds = new System.Drawing.Rectangle(winLoc, new System.Drawing.Size(winWidth, winHeight));

                elementHost.Dock = DockStyle.None;
                elementHost.Location = new System.Drawing.Point(12, 12);
                elementHost.Width = eHostWidth;
                elementHost.Height = eHostHeight;
                elementHost.Anchor = AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Left;
                Cursor.Show();
                this.Show();
            }
            else
            {
                winState = this.WindowState;
                winLoc = this.Location;
                winWidth = this.Width;
                winHeight = this.Height;
                eHostWidth = elementHost.Width;
                eHostHeight = elementHost.Height;

                this.Hide();
                this.FormBorderStyle = FormBorderStyle.None;
                this.TopMost = true;

                Screen[] screens = Screen.AllScreens;
                Screen targetScreen;

                // Check if the requested screen index is valid
                if (FullScreenMonitorIndex >= 0 && FullScreenMonitorIndex < screens.Length)
                {
                    targetScreen = screens[FullScreenMonitorIndex];
                }
                else
                {
                    // If the monitor index is invalid (e.g., screen disconnected), default to primary.
                    targetScreen = Screen.PrimaryScreen;
                    FullScreenMonitorIndex = 0;
                }

                this.WindowState = FormWindowState.Normal; // Must be normal to change bounds
                this.Bounds = targetScreen.Bounds;

                elementHost.Dock = DockStyle.Fill;
                Cursor.Hide();
                this.Show();
            }
        }

        private void DShowForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SystemEvents.DisplaySettingsChanged -= SystemEvents_DisplaySettingsChanged;
            mediaUriElement.MediaUriPlayer.MediaPositionChanged -= MediaUriPlayer_MediaPositionChanged;
            mediaUriElement.Stop();
            mediaUriElement.Close();
            mediaUriElement.Source = null;
            mediaUriElement.VideoSource = null;

            if (mouseClickTimer != null)
                mouseClickTimer.Dispose();

            NativeMethods.SystemSleepManagement.ResotreSleep();
            if (this.Owner != null)
            {
                this.Owner.Show();
                this.Owner.TopMost = (Global.MainCfgAlwaysOnTop == "True") ? true : false;
            }
        }

        private void Player_VolumeSlider_ValueChanged(object sender, EventArgs e)
        {
            int val = Player_VolumeSlider.TrackBarValue;
            
            // Enforce step of 5
            int remainder = val % 5;
            if (remainder != 0)
            {
                if (remainder >= 3) val += (5 - remainder);
                else val -= remainder;

                // Clamp to valid range
                if (val > Player_VolumeSlider.Maximum) val = Player_VolumeSlider.Maximum;
                if (val < Player_VolumeSlider.Minimum) val = Player_VolumeSlider.Minimum;

                if (Player_VolumeSlider.TrackBarValue != val)
                {
                    Player_VolumeSlider.TrackBarValue = val;
                    return; // Return to let the new event handle the update
                }
            }

            double newVol = (double)val / 100.0;
            if (mediaUriElement.Volume != newVol)
            {
                mediaUriElement.Volume = newVol;
            }
            Player_VolumeSlider.ProgressBarValue = val;
            Player_CurrentVolValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = val.ToString() + " %");
            if(val.ToString() == SongVolume)
            {
                Player_UpdateVolume_Button.Enabled = false;
            }
            else
            {
                Player_UpdateVolume_Button.Enabled = true;
            }
        }
            
        private void Player_VolumeSlider_Click(object sender, EventArgs e)
        {
             // Handled by ValueChanged, but kept for Designer compatibility or click-specific logic if needed.
             // Currently ProgressTrackBar updates TrackBarValue on mouse moves/clicks which triggers ValueChanged.
        }

        private void DShowForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            mediaUriElement = null;
            GC.Collect();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    public partial class MainForm : Form
    {
        private void SongMonitor_CheckCurSong()
        {
            if (!Global.CrazyktvDatabaseStatus) return;

            if (Global.SongMgrSongAddMode != "4")
            {
                if (SongMgrCfg_TabControl.TabPages.IndexOf(SongMgrCfg_MonitorFolders_TabPage) >= 0)
                {
                    SongMgrCfg_MonitorFolders_TabPage.Hide();
                    SongMgrCfg_TabControl.TabPages.Remove(SongMgrCfg_MonitorFolders_TabPage);
                }
            }

            if (Global.SongMgrSongAddMode == "4" && Global.SongMgrEnableMonitorFolders == "True")
            {
                Global.TimerStartTime = DateTime.Now;
                DateTime TimerStartTime = DateTime.Now;
                Global.TotalList = new List<int>() { 0, 0, 0, 0, 0 };
                Global.MTotalList = new List<int>() { 0, 0, 0, 0 };
                Common_SwitchSetUI(false);

                var tasks = new List<Task>()
                {
                    Task.Factory.StartNew(() => SongMonitor_CheckCurSongTask())
                };

                Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                {
                    DateTime TimerEndTime = DateTime.Now;
                    this.BeginInvoke((Action)delegate()
                    {
                        SongQuery_QueryStatus_Label.Text = "總共加入 " + Global.TotalList[0] + " 首歌曲,忽略重複歌曲 " + Global.TotalList[1] + " 首,移除 " + Global.MTotalList[0] + " 首,共花費 " + (long)(TimerEndTime - TimerStartTime).TotalSeconds + " 秒完成監視。";
                        SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                        SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;

                        Common_SwitchSetUI(true);
                    });
                });
            }
        }

        private void SongMonitor_CheckCurSongTask()
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] Start execution");
                SpinWait.SpinUntil(() => Global.InitializeSongData == true);
                this.BeginInvoke((Action)delegate()
                {
                    SongQuery_QueryStatus_Label.Text = "正在檢查歌曲資料庫,請稍待...";
                    SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                    SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                });

                DataTable dt = new DataTable();
                string SongQuerySqlStr = "select Song_Id, Song_Lang, Song_FileName, Song_Path from ktv_Song order by Song_Id";
                dt = CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] Loaded DB song records count: {dt?.Rows.Count ?? 0}");

                List<string> RemoveSongIdList = new List<string>();
                HashSet<string> FileList = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string songId = row["Song_Id"]?.ToString() ?? "";
                        string songPath = row["Song_Path"]?.ToString() ?? "";
                        string songFile = row["Song_FileName"]?.ToString() ?? "";
                        string fullPath = Path.GetFullPath(Path.Combine(songPath, songFile));

                        if (!File.Exists(fullPath))
                        {
                            RemoveSongIdList.Add(songId + "|" + fullPath);
                        }
                        else
                        {
                            FileList.Add(fullPath);
                        }
                    }
                }
                if (dt != null)
                {
                    dt.Dispose();
                    dt = null;
                }

                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] DB check complete. RemoveCount: {RemoveSongIdList.Count}, ExistingFileCount: {FileList.Count}");

                OleDbConnection conn = CommonFunc.OleDbOpenConn(Global.CrazyktvDatabaseFile, "");
                OleDbCommand cmd = new OleDbCommand();
                string SongRemoveSqlStr = "delete from ktv_Song where Song_Id = @SongId";
                cmd = new OleDbCommand(SongRemoveSqlStr, conn);

                if (RemoveSongIdList.Count > 0)
                {
                    foreach (string str in RemoveSongIdList)
                    {
                        List<string> list = new List<string>(str.Split('|'));
                        cmd.Parameters.AddWithValue("@SongId", list[0]);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        Global.MTotalList[0]++;

                        this.BeginInvoke((Action)delegate()
                        {
                            SongQuery_QueryStatus_Label.Text = "正在移除第 " + Global.MTotalList[0] + " 首檔案不存在的歌曲資料,請稍待...";
                            SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                            SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                        });

                        lock (LockThis)
                        {
                            Global.SongLogDT.Rows.Add(Global.SongLogDT.NewRow());
                            Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][0] = "【歌庫監視】檔案不存在: " + list[0] + "|" + list[1];
                            Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][1] = Global.SongLogDT.Rows.Count;
                        }
                    }
                }
                conn.Close();

                this.BeginInvoke((Action)delegate()
                {
                    SongQuery_QueryStatus_Label.Text = "正在檢查歌庫監視目錄,請稍待...";
                    SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                    SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                });

                List<string> NewFileList = new List<string>();
                List<string> SupportFormat = new List<string>(Global.SongMgrSupportFormat.Split(';'));

                bool EnableSongMonitor = false;
                for (int i = 0; i < Global.SongMgrMonitorFoldersList.Count; i++)
                {
                    string mpath = Global.SongMgrMonitorFoldersList[i];
                    System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] Checking monitor folder [{i}]: {mpath}");
                    if (!string.IsNullOrEmpty(mpath) && Directory.Exists(mpath))
                    {
                        EnableSongMonitor = true;
                        DirectoryInfo dir = new DirectoryInfo(mpath);
                        FileInfo[] Files = dir.GetFiles("*", SearchOption.AllDirectories).Where(p => SupportFormat.Contains(p.Extension.ToLower())).ToArray();
                        foreach (FileInfo f in Files)
                        {
                            string fFullName = Path.GetFullPath(f.FullName);
                            if (!FileList.Contains(fFullName))
                            {
                                lock (LockThis)
                                {
                                    NewFileList.Add(fFullName);

                                    Global.SongLogDT.Rows.Add(Global.SongLogDT.NewRow());
                                    Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][0] = "【歌庫監視】偵測到新檔: " + fFullName;
                                    Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][1] = Global.SongLogDT.Rows.Count;
                                }
                            }
                        }

                        if (i < Global.SongMonitorWatcher.Length && !Global.SongMonitorWatcher[i].EnableRaisingEvents)
                        {
                            Global.SongMonitorWatcher[i].Path = mpath;
                            Global.SongMonitorWatcher[i].IncludeSubdirectories = true;
                            Global.SongMonitorWatcher[i].Filter = "*.*";
                            Global.SongMonitorWatcher[i].InternalBufferSize = 8192;
                            Global.SongMonitorWatcher[i].NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName;
                            Global.SongMonitorWatcher[i].Created += new FileSystemEventHandler(WatcherOnCreated);
                            Global.SongMonitorWatcher[i].Deleted += new FileSystemEventHandler(WatcherOnDeleted);
                            Global.SongMonitorWatcher[i].Renamed += new RenamedEventHandler(WatcherOnRenamed);
                            Global.SongMonitorWatcher[i].EnableRaisingEvents = true;
                            System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] FileSystemWatcher[{i}] enabled for {mpath}");
                        }
                    }
                }

                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] Folder scan complete. NewFilesFound: {NewFileList.Count}");

                if (NewFileList.Count > 0)
                {
                    SongAdd_SongAnalysisTask(NewFileList);
                    while (!SongAnalysis.SongAnalysisCompleted)
                    {
                        Thread.Sleep(500);
                    }

                    if (!SongAnalysis.SongAnalysisError)
                    {
                        SongAdd_SongAddTask();
                    }
                }
                else
                {
                    if (RemoveSongIdList.Count > 0)
                    {
                        this.BeginInvoke((Action)delegate()
                        {
                            Common_QueryAddSong(100);
                        });

                        Task.Factory.StartNew(() => Common_GetSongStatisticsTask());
                        Task.Factory.StartNew(() => Common_GetSingerStatisticsTask());
                        Task.Factory.StartNew(() => CommonFunc.GetRemainingSongIdCount((Global.SongMgrMaxDigitCode == "1") ? 5 : 6));
                    }
                }

                if (EnableSongMonitor)
                {
                    Global.SongMonitorDT = new DataTable();
                    SongQuerySqlStr = "select Song_Id, Song_Lang, Song_FileName, Song_Path from ktv_Song order by Song_Id";
                    Global.SongMonitorDT = CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");
                    System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] Global.SongMonitorDT initialized with {Global.SongMonitorDT?.Rows.Count ?? 0} rows");

                    this.BeginInvoke((Action)delegate()
                    {
                        if (!Global.SongMonitorTimer.Enabled)
                        {
                            Global.SongMonitorTimer.Tick += new EventHandler(SongMonitorTimer_Tick);
                            Global.SongMonitorTimer.Interval = 1000;
                            Global.SongMonitorTimer.Start();
                            System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask] Global.SongMonitorTimer started");
                        }
                    });
                }
                RemoveSongIdList.Clear();
                FileList.Clear();
                NewFileList.Clear();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_CheckCurSongTask EXCEPTION] {ex}");
            }
        }

        private void SongMonitor_SwitchSongMonitorWatcher()
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_SwitchSongMonitorWatcher] Start execution");
                if (!Global.CrazyktvDatabaseStatus) return;

                if (Global.SongMgrSongAddMode == "4" && Global.SongMgrEnableMonitorFolders == "True")
                {
                    bool EnableSongMonitor = false;
                    for (int i = 0; i < Global.SongMgrMonitorFoldersList.Count; i++)
                    {
                        string mpath = Global.SongMgrMonitorFoldersList[i];
                        if (!string.IsNullOrEmpty(mpath) && Directory.Exists(mpath))
                        {
                            EnableSongMonitor = true;
                            if (i < Global.SongMonitorWatcher.Length && !Global.SongMonitorWatcher[i].EnableRaisingEvents)
                            {
                                Global.SongMonitorWatcher[i].Path = mpath;
                                Global.SongMonitorWatcher[i].IncludeSubdirectories = true;
                                Global.SongMonitorWatcher[i].Filter = "*.*";
                                Global.SongMonitorWatcher[i].InternalBufferSize = 8192;
                                Global.SongMonitorWatcher[i].NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName;
                                Global.SongMonitorWatcher[i].Created += new FileSystemEventHandler(WatcherOnCreated);
                                Global.SongMonitorWatcher[i].Deleted += new FileSystemEventHandler(WatcherOnDeleted);
                                Global.SongMonitorWatcher[i].Renamed += new RenamedEventHandler(WatcherOnRenamed);
                                Global.SongMonitorWatcher[i].EnableRaisingEvents = true;
                                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_SwitchSongMonitorWatcher] Watcher[{i}] enabled for {mpath}");
                            }
                        }
                    }

                    if (EnableSongMonitor)
                    {
                        if (Global.SongMonitorDT != null)
                        {
                            Global.SongMonitorDT.Dispose();
                            Global.SongMonitorDT = null;
                        }

                        Global.SongMonitorDT = new DataTable();
                        string SongQuerySqlStr = "select Song_Id, Song_Lang, Song_FileName, Song_Path from ktv_Song order by Song_Id";
                        Global.SongMonitorDT = CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");

                        if (!Global.SongMonitorTimer.Enabled)
                        {
                            Global.SongMonitorTimer.Tick += new EventHandler(SongMonitorTimer_Tick);
                            Global.SongMonitorTimer.Interval = 1000;
                            Global.SongMonitorTimer.Start();
                        }
                    }
                }
                else
                {
                    foreach (FileSystemWatcher fwatcher in Global.SongMonitorWatcher)
                    {
                        if (fwatcher.EnableRaisingEvents)
                        {
                            fwatcher.Created -= new FileSystemEventHandler(WatcherOnCreated);
                            fwatcher.Deleted -= new FileSystemEventHandler(WatcherOnDeleted);
                            fwatcher.Renamed -= new RenamedEventHandler(WatcherOnRenamed);
                            fwatcher.EnableRaisingEvents = false;
                        }
                    }

                    if (Global.SongMonitorDT != null)
                    {
                        Global.SongMonitorDT.Dispose();
                        Global.SongMonitorDT = null;
                    }

                    if (Global.SongMonitorTimer.Enabled)
                    {
                        Global.SongMonitorTimer.Tick -= new EventHandler(SongMonitorTimer_Tick);
                        Global.SongMonitorTimer.Stop();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_SwitchSongMonitorWatcher EXCEPTION] {ex}");
            }
        }


        private static void WatcherOnCreated(object sender, FileSystemEventArgs e)
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][WatcherOnCreated] Path: {e.FullPath}");
                List<string> SupportFormat = new List<string>(Global.SongMgrSupportFormat.Split(';'));
                if (SupportFormat.Contains(Path.GetExtension(e.FullPath).ToLower()))
                {
                    if (Global.SongMonitorCreatedList.IndexOf(e.FullPath) < 0) Global.SongMonitorCreatedList.Add(e.FullPath);
                }
                else
                {
                    if (Directory.Exists(e.FullPath))
                    {
                        DirectoryInfo dir = new DirectoryInfo(e.FullPath);
                        FileInfo[] Files = dir.GetFiles("*", SearchOption.AllDirectories).Where(p => SupportFormat.Contains(p.Extension.ToLower())).ToArray();
                        foreach (FileInfo f in Files)
                        {
                            if (Global.SongMonitorCreatedList.IndexOf(f.FullName) < 0) Global.SongMonitorCreatedList.Add(f.FullName);
                        }
                    }
                }
                if (Global.SongMonitorCreatedList.Count > 0) Global.SongMonitorSTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][WatcherOnCreated EXCEPTION] Path: {e.FullPath}, Exception: {ex}");
            }
        }


        private static void WatcherOnDeleted(object sender, FileSystemEventArgs e)
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][WatcherOnDeleted] Path: {e.FullPath}");
                if (Global.SongMonitorDT == null) return;

                List<string> SupportFormat = new List<string>(Global.SongMgrSupportFormat.Split(';'));
                if (SupportFormat.Contains(Path.GetExtension(e.FullPath).ToLower()))
                {
                    string dirPath = (Path.GetDirectoryName(e.FullPath) ?? "").TrimEnd('\\');
                    string fileName = Path.GetFileName(e.FullPath);

                    var query = from row in Global.SongMonitorDT.AsEnumerable()
                                where (row.Field<string>("Song_Path") ?? "").TrimEnd('\\').Equals(dirPath, StringComparison.OrdinalIgnoreCase) &&
                                      (row.Field<string>("Song_FileName") ?? "").Equals(fileName, StringComparison.OrdinalIgnoreCase)
                                select row;

                    if (query.Count<DataRow>() > 0)
                    {
                        foreach (DataRow row in query)
                        {
                            if (Global.SongMonitorDeletedList.IndexOf(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"))) < 0) Global.SongMonitorDeletedList.Add(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")));
                            break;
                        }
                    }
                }
                else
                {
                    if (Path.GetExtension(e.FullPath) == "")
                    {
                        string folderPath = e.FullPath.TrimEnd('\\');
                        var query = from row in Global.SongMonitorDT.AsEnumerable()
                                    where (row.Field<string>("Song_Path") ?? "").TrimEnd('\\').Equals(folderPath, StringComparison.OrdinalIgnoreCase) ||
                                          (row.Field<string>("Song_Path") ?? "").StartsWith(folderPath + @"\", StringComparison.OrdinalIgnoreCase)
                                    select row;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow row in query)
                            {
                                if (!File.Exists(Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"))))
                                {
                                    if (Global.SongMonitorDeletedList.IndexOf(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"))) < 0) Global.SongMonitorDeletedList.Add(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")));
                                }
                            }
                        }
                    }
                }
                if (Global.SongMonitorDeletedList.Count > 0) Global.SongMonitorSTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][WatcherOnDeleted EXCEPTION] Path: {e.FullPath}, Exception: {ex}");
            }
        }


        private static void WatcherOnRenamed(object sender, RenamedEventArgs e)
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][WatcherOnRenamed] Old: {e.OldFullPath} -> New: {e.FullPath}");
                if (Global.SongMonitorDT == null) return;

                List<string> SupportFormat = new List<string>(Global.SongMgrSupportFormat.Split(';'));
                if (SupportFormat.Contains(Path.GetExtension(e.OldFullPath).ToLower()))
                {
                    string dirPath = (Path.GetDirectoryName(e.OldFullPath) ?? "").TrimEnd('\\');
                    string fileName = Path.GetFileName(e.OldFullPath);

                    var query = from row in Global.SongMonitorDT.AsEnumerable()
                                where (row.Field<string>("Song_Path") ?? "").TrimEnd('\\').Equals(dirPath, StringComparison.OrdinalIgnoreCase) &&
                                      (row.Field<string>("Song_FileName") ?? "").Equals(fileName, StringComparison.OrdinalIgnoreCase)
                                select row;

                    if (query.Count<DataRow>() > 0)
                    {
                        foreach (DataRow row in query)
                        {
                            if (Global.SongMonitorDeletedList.IndexOf(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"))) < 0) Global.SongMonitorDeletedList.Add(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")));
                            break;
                        }
                    }
                }
                else
                {
                    if (Path.GetExtension(e.OldFullPath) == "")
                    {
                        string folderPath = e.OldFullPath.TrimEnd('\\');
                        var query = from row in Global.SongMonitorDT.AsEnumerable()
                                    where (row.Field<string>("Song_Path") ?? "").TrimEnd('\\').Equals(folderPath, StringComparison.OrdinalIgnoreCase) ||
                                          (row.Field<string>("Song_Path") ?? "").StartsWith(folderPath + @"\", StringComparison.OrdinalIgnoreCase)
                                    select row;

                        if (query.Count<DataRow>() > 0)
                        {
                            foreach (DataRow row in query)
                            {
                                if (!File.Exists(Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"))))
                                {
                                    if (Global.SongMonitorDeletedList.IndexOf(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName"))) < 0) Global.SongMonitorDeletedList.Add(row["Song_Id"].ToString() + "|" + Path.Combine(row.Field<string>("Song_Path"), row.Field<string>("Song_FileName")));
                                }
                            }
                        }
                    }
                }

                if (SupportFormat.Contains(Path.GetExtension(e.FullPath).ToLower()))
                {
                    if (Global.SongMonitorCreatedList.IndexOf(e.FullPath) < 0) Global.SongMonitorCreatedList.Add(e.FullPath);
                }
                else
                {
                    if (Directory.Exists(e.FullPath))
                    {
                        DirectoryInfo dir = new DirectoryInfo(e.FullPath);
                        FileInfo[] Files = dir.GetFiles("*", SearchOption.AllDirectories).Where(p => SupportFormat.Contains(p.Extension.ToLower())).ToArray();
                        foreach (FileInfo f in Files)
                        {
                            if (Global.SongMonitorCreatedList.IndexOf(f.FullName) < 0) Global.SongMonitorCreatedList.Add(f.FullName);
                        }
                    }
                }
                if (Global.SongMonitorDeletedList.Count > 0 || Global.SongMonitorCreatedList.Count > 0) Global.SongMonitorSTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][WatcherOnRenamed EXCEPTION] Exception: {ex}");
            }
        }


        private void SongMonitorTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (Global.SongMonitorDeletedList.Count > 0 || Global.SongMonitorCreatedList.Count > 0)
                {
                    System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitorTimer_Tick] Triggered. DeletedList: {Global.SongMonitorDeletedList.Count}, CreatedList: {Global.SongMonitorCreatedList.Count}");
                    Global.SongMonitorETime = DateTime.Now;
                    if ((long)(Global.SongMonitorETime - Global.SongMonitorSTime).TotalSeconds >= 3)
                    {
                        Global.SongMonitorTimer.Stop();
                        Global.TimerStartTime = DateTime.Now;
                        DateTime TimerStartTime = DateTime.Now;
                        Global.TotalList = new List<int>() { 0, 0, 0, 0, 0 };
                        Global.MTotalList = new List<int>() { 0, 0, 0, 0 };
                        Common_SwitchSetUI(false);

                        SongQuery_QueryStatus_Label.Text = "正在同步檔案及資料庫,請稍待...";
                        SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                        SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;

                        var tasks = new List<Task>()
                        {
                            Task.Factory.StartNew(() => SongMonitor_UpdateSongDBTask())
                        };

                        Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
                        {
                            DateTime TimerEndTime = DateTime.Now;
                            this.BeginInvoke((Action)delegate()
                            {
                                SongQuery_QueryStatus_Label.Text = "總共加入 " + Global.TotalList[0] + " 首歌曲,忽略重複歌曲 " + Global.TotalList[1] + " 首,移除 " + Global.MTotalList[0] + " 首,共花費 " + (long)(TimerEndTime - TimerStartTime).TotalSeconds + " 秒完成監視。";
                                SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                                SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;

                                if (Global.SongMonitorDT != null)
                                {
                                    Global.SongMonitorDT.Dispose();
                                    Global.SongMonitorDT = null;
                                }

                                Global.SongMonitorDT = new DataTable();
                                string SongQuerySqlStr = "select Song_Id, Song_Lang, Song_FileName, Song_Path from ktv_Song order by Song_Id";
                                Global.SongMonitorDT = CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");

                                Common_SwitchSetUI(true);
                                Global.SongMonitorTimer.Start();
                            });
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitorTimer_Tick EXCEPTION] {ex}");
            }
        }


        private void SongMonitor_UpdateSongDBTask()
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_UpdateSongDBTask] Start execution");
                List<int> CountList = new List<int>() { Global.SongMonitorDeletedList.Count, Global.SongMonitorCreatedList.Count };
                if (Global.SongMonitorDeletedList.Count > 0)
                {
                    OleDbConnection conn = CommonFunc.OleDbOpenConn(Global.CrazyktvDatabaseFile, "");
                    OleDbCommand cmd = new OleDbCommand();
                    string SongRemoveSqlStr = "delete from ktv_Song where Song_Id = @SongId";
                    cmd = new OleDbCommand(SongRemoveSqlStr, conn);

                    foreach (string str in Global.SongMonitorDeletedList)
                    {
                        List<string> list = new List<string>(str.Split('|'));
                        cmd.Parameters.AddWithValue("@SongId", list[0]);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        Global.MTotalList[0]++;

                        this.BeginInvoke((Action)delegate()
                        {
                            SongQuery_QueryStatus_Label.Text = "正在移除第 " + Global.MTotalList[0] + " 首檔案不存在的歌曲資料,請稍待...";
                            SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                            SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                        });

                        lock (LockThis)
                        {
                            Global.SongLogDT.Rows.Add(Global.SongLogDT.NewRow());
                            Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][0] = "【歌庫監視】檔案已刪除: " + list[0] + "|" + list[1];
                            Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][1] = Global.SongLogDT.Rows.Count;
                        }
                    }
                    Global.SongMonitorDeletedList.Clear();
                    conn.Close();
                }

                if (Global.SongMonitorCreatedList.Count > 0)
                {
                    SongAdd_SongAnalysisTask(Global.SongMonitorCreatedList);
                    while (!SongAnalysis.SongAnalysisCompleted)
                    {
                        Thread.Sleep(500);
                    }

                    if (!SongAnalysis.SongAnalysisError)
                    {
                        SongAdd_SongAddTask();
                    }

                    this.BeginInvoke((Action)delegate()
                    {
                        SongQuery_QueryStatus_Label.Text = "正在將變更的歌曲資料寫入操作記錄,請稍待...";
                        SongAdd_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                        SongMgrCfg_Tooltip_Label.Text = SongQuery_QueryStatus_Label.Text;
                    });

                    foreach (string filepath in Global.SongMonitorCreatedList)
                    {
                        lock (LockThis)
                        {
                            Global.SongLogDT.Rows.Add(Global.SongLogDT.NewRow());
                            Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][0] = "【歌庫監視】偵測到新檔: " + filepath;
                            Global.SongLogDT.Rows[Global.SongLogDT.Rows.Count - 1][1] = Global.SongLogDT.Rows.Count;
                        }
                    }
                    Global.SongMonitorCreatedList.Clear();
                }

                if (CountList[0] > 0 && CountList[1] == 0)
                {
                    this.BeginInvoke((Action)delegate()
                    {
                        Common_QueryAddSong(100);
                    });

                    Task.Factory.StartNew(() => Common_GetSongStatisticsTask());
                    Task.Factory.StartNew(() => Common_GetSingerStatisticsTask());
                    Task.Factory.StartNew(() => CommonFunc.GetRemainingSongIdCount((Global.SongMgrMaxDigitCode == "1") ? 5 : 6));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][Thread {Thread.CurrentThread.ManagedThreadId}][SongMonitor_UpdateSongDBTask EXCEPTION] {ex}");
            }
        }





    }





    class SongMonitor
    {
    }
}

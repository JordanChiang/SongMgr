using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.EndToEnd
{
    /// <summary>
    /// The real CrazyKTV_SongMgr.exe in a sandbox folder: its own CrazySong.mdb (a copy of CrazySongEmpty.mdb),
    /// its own CrazyKTV_SongMgr.cfg, and a 歌庫監視 folder of tiny ffmpeg-made songs. Driven through UI Automation
    /// the way a user would, then checked by reading the sandbox's database and 設定檔.
    /// Needs an interactive desktop; set SONGMGR_SKIP_E2E=1 to skip these tests.
    /// </summary>
    internal sealed class AppSession : IDisposable
    {
        public static readonly List<string> Langs = new List<string> { "國語", "台語", "粵語", "日語", "英語", "客語", "原住民語", "韓語", "兒歌", "其它" };
        public const string Code5 = "10001,20001,30001,40001,50001,60001,70001,80001,90001,95001";
        public const string Code6 = "100001,200001,300001,400001,500001,600001,700001,800001,900001,950001";

        static readonly string[] AppFiles =
        {
            "CrazyKTV_SongMgr.exe", "CrazyKTV_SongMgr.dll", "CrazyKTV_SongMgr.runtimeconfig.json", "CrazyKTV_SongMgr.deps.json",
            "HtmlAgilityPack.dll", "Newtonsoft.Json.dll", "System.Data.OleDb.dll", "System.Data.SQLite.dll",
        };
        static readonly string[] AppFolders = { "runtimes", "SongMgr", "Tools" };

        public readonly string Root;
        public readonly string AppDir;
        public readonly string MonitorDir;
        public string Database => Path.Combine(AppDir, "CrazySong.mdb");
        public string SettingsPath => Path.Combine(AppDir, "CrazyKTV_SongMgr.cfg");

        Process process;
        AutomationElement main;
        readonly List<string> tabLog = new List<string>();

        public static string SkipReason =>
            Environment.GetEnvironmentVariable("SONGMGR_SKIP_E2E") == "1" ? "SONGMGR_SKIP_E2E=1" :
            !Environment.UserInteractive ? "no interactive desktop" : null;

        public AppSession(string name)
        {
            Root = TestEnv.NewTempDir("e2e_" + name);
            AppDir = Path.Combine(Root, "app");
            MonitorDir = Path.Combine(Root, "監視");
            Directory.CreateDirectory(AppDir);
            Directory.CreateDirectory(MonitorDir);

            string bin = AppContext.BaseDirectory;
            foreach (string f in AppFiles) File.Copy(Path.Combine(bin, f), Path.Combine(AppDir, f));
            foreach (string d in AppFolders) CopyDir(Path.Combine(bin, d), Path.Combine(AppDir, d));
            File.Copy(Path.Combine(TestEnv.FixtureDir, "CrazySongEmpty.mdb"), Database);
            File.SetAttributes(Database, FileAttributes.Normal);
        }

        static void CopyDir(string from, string to)
        {
            Directory.CreateDirectory(to);
            foreach (string f in Directory.GetFiles(from)) File.Copy(f, Path.Combine(to, Path.GetFileName(f)));
            foreach (string d in Directory.GetDirectories(from)) CopyDir(d, Path.Combine(to, Path.GetFileName(d)));
        }

        // ---------- setting up the 歌庫 ----------

        /// <summary>Writes the 設定檔. Automatic program and 歌庫版本 updates are always off: no network, no prompts.</summary>
        public void WriteSettings(int digits, params (string Name, string Value)[] more)
        {
            var settings = new List<(string, string)>
            {
                ("CrazyktvDatabaseFile", Database),
                ("SongMgrMaxDigitCode", digits == 5 ? "1" : "2"),
                ("SongMgrLangCode", digits == 5 ? Code5 : Code6),
                ("MainCfgEnableAutoUpdate", "False"),
                ("DBVerEnableDBVerUpdate", "False"),
            };
            foreach (var s in more)
            {
                settings.RemoveAll(x => x.Item1 == s.Name);
                settings.Add(s);
            }
            SettingsFile.Create(SettingsPath, settings);
        }

        public void WatchMonitorFolder(int digits)
        {
            WriteSettings(digits, ("SongMgrSongAddMode", "4"), ("SongMgrEnableMonitorFolders", "True"), ("SongMgrMonitorFolders", MonitorDir + ",,,,"));
        }

        public string Setting(string name) => SettingsFile.Load(SettingsPath)[name];

        /// <summary>A 3-second song file in the 歌庫監視 folder, at <paramref name="relative"/> (語系 folder + 歌手_歌名).</summary>
        public string MakeSongFile(string relative)
        {
            string path = Path.Combine(MonitorDir, relative + ".mpg");
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var psi = new ProcessStartInfo(Path.Combine(AppDir, "Tools", "ffmpeg.exe"),
                "-loglevel error -y -f lavfi -i testsrc=duration=3:size=320x240:rate=25 -f lavfi -i sine=frequency=440:duration=3 -ac 2 -c:v mpeg2video -c:a mp2 -shortest \"" + path + "\"")
            { UseShellExecute = false, CreateNoWindow = true };
            using (Process p = Process.Start(psi))
            {
                Assert.True(p.WaitForExit(60000) && p.ExitCode == 0, "ffmpeg could not make " + path);
            }
            return path;
        }

        public void InsertSong(string id, string lang, string singer, string name, string file = null)
        {
            file = file ?? Path.Combine(MonitorDir, lang, singer + "_" + name + ".mpg");
            Execute("insert into ktv_Song (Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState) values (@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s)",
                id, lang, "0", singer, name, "1", "", "100", name.Length.ToString(), "0", "0.5", "2024/01/02 03:04:05", Path.GetFileName(file), Path.GetDirectoryName(file) + @"\", "ㄍㄇ", "42", "12", "丨一", "0");
        }

        public void InsertFavorite(string userId, string songId) => Execute("insert into ktv_Favorite (User_Id, Song_Id) values (@a, @b)", userId, songId);

        public void Execute(string sql, params string[] values)
        {
            using (var conn = DbHelper.Access.OpenConn(Database))
            using (var cmd = new OleDbCommand(sql, conn))
            {
                foreach (string v in values) cmd.Parameters.AddWithValue("@p", v);
                cmd.ExecuteNonQuery();
            }
        }

        public List<(string Id, string Lang, string Singer, string Name)> Songs()
        {
            using (DataTable dt = Db.GetDataTable(Database, "select Song_Id, Song_Lang, Song_Singer, Song_SongName from ktv_Song order by Song_Id", ""))
            {
                return dt.AsEnumerable().Select(r => (r.Field<string>(0), r.Field<string>(1), r.Field<string>(2), r.Field<string>(3))).ToList();
            }
        }

        public List<string> Favorites()
        {
            using (DataTable dt = Db.GetDataTable(Database, "select User_Id, Song_Id from ktv_Favorite order by Song_Id", ""))
            {
                return dt.AsEnumerable().Select(r => r.Field<string>(0) + ":" + r.Field<string>(1)).ToList();
            }
        }

        // ---------- driving the program ----------

        public void Start()
        {
            process = Process.Start(new ProcessStartInfo(Path.Combine(AppDir, "CrazyKTV_SongMgr.exe")) { WorkingDirectory = AppDir, UseShellExecute = false });
            WaitUntil(() =>
            {
                // A machine without VMR9 warns before the main window opens.
                AnswerIfShown("顯示環境檢查", "確定");
                main = TopWindows().FirstOrDefault(w => w.Current.ClassName != "#32770" && w.Current.Name.StartsWith("CrazyKTV"));
                return main != null;
            }, 60, "the SongMgr window to open");
        }

        IEnumerable<AutomationElement> TopWindows()
        {
            var cond = new PropertyCondition(AutomationElement.ProcessIdProperty, process.Id);
            return AutomationElement.RootElement.FindAll(TreeScope.Children, cond).Cast<AutomationElement>();
        }

        /// <summary>
        /// Answers a MessageBox of SongMgr's by its title, with Win32 rather than UI Automation: while the button that
        /// opened it is still being invoked, UI Automation calls into SongMgr queue behind that one.
        /// </summary>
        bool AnswerIfShown(string title, string button)
        {
            int command = button.StartsWith("是") ? IDYES : button.StartsWith("否") ? IDNO : IDOK;
            IntPtr found = IntPtr.Zero;
            EnumWindows((hwnd, _) =>
            {
                GetWindowThreadProcessId(hwnd, out int pid);
                if (pid != process.Id || !IsWindowVisible(hwnd)) return true;
                var cls = new System.Text.StringBuilder(64);
                var text = new System.Text.StringBuilder(256);
                GetClassName(hwnd, cls, cls.Capacity);
                GetWindowText(hwnd, text, text.Capacity);
                if (cls.ToString() == "#32770" && text.ToString().Contains(title)) { found = hwnd; return false; }
                return true;
            }, IntPtr.Zero);
            if (found == IntPtr.Zero) return false;
            PostMessage(found, WM_COMMAND, new IntPtr(command), IntPtr.Zero);
            return true;
        }

        /// <summary>Waits for the MessageBox titled <paramref name="title"/> and presses <paramref name="button"/>.</summary>
        public void Answer(string title, string button) => WaitUntil(() => AnswerIfShown(title, button), 30, "the 「" + title + "」 dialog");

        static AutomationElement ByName(AutomationElement root, string name, ControlType type = null)
        {
            var all = root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name)).Cast<AutomationElement>().ToList();
            return all.FirstOrDefault(e => type == null || e.Current.ControlType == type) ?? (type == null ? null : all.FirstOrDefault());
        }

        AutomationElement ById(string id)
        {
            AutomationElement e = null;
            WaitUntil(() => (e = main.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id))) != null, 20, id);
            return e;
        }

        /// <summary>
        /// Selects a tab, presses a button or toggles a check box, by its caption. With <paramref name="onTab"/>, that
        /// tab is selected again until the control shows up: SongMgr switches tabs itself while it starts.
        /// </summary>
        public void Click(string name, string onTab = null)
        {
            try
            {
                ClickOrTimeout(name, onTab);
            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException(ex.Message + Environment.NewLine + Diagnose(), ex);
            }
        }

        /// <summary>The last tab moves, a screenshot (outside the sandbox, which is deleted) and the buttons and tabs on screen.</summary>
        string Diagnose()
        {
            string shot = Path.Combine(Path.GetTempPath(), "SongMgrTests", "e2e-" + Path.GetFileName(Root) + ".png");
            try
            {
                System.Windows.Rect r = main.Current.BoundingRectangle;
                using (var bmp = new System.Drawing.Bitmap((int)r.Width, (int)r.Height))
                using (var g = System.Drawing.Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen((int)r.X, (int)r.Y, 0, 0, bmp.Size);
                    bmp.Save(shot, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            catch (Exception ex)
            {
                shot = "(no screenshot: " + ex.Message + ")";
            }
            var seen = main.FindAll(TreeScope.Descendants, Condition.TrueCondition).Cast<AutomationElement>()
                .Where(x => x.Current.ControlType == ControlType.Button || x.Current.ControlType == ControlType.TabItem)
                .Select(x => x.Current.ControlType.ProgrammaticName.Replace("ControlType.", "") + ":" + x.Current.Name + (x.Current.IsEnabled ? "" : "(disabled)"));
            return "Tabs: " + string.Join(" / ", tabLog.Skip(Math.Max(0, tabLog.Count - 8))) + Environment.NewLine + "Screenshot: " + shot + Environment.NewLine + "On screen: " + string.Join(", ", seen);
        }

        void ClickOrTimeout(string name, string onTab)
        {
            AutomationElement e = null;
            WaitUntil(() =>
            {
                if (onTab != null) SelectTab(onTab);
                return (e = ByName(main, name, ControlType.Button) ?? ByName(main, name)) != null && e.Current.IsEnabled;
            }, 30, "「" + name + "」");
            if (e.Current.ControlType == ControlType.Button && e.Current.NativeWindowHandle != 0)
            {
                // Posted, not invoked: a Click handler that opens a MessageBox would hold an Invoke until it closes.
                PostMessage(new IntPtr(e.Current.NativeWindowHandle), BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            }
            else if (e.TryGetCurrentPattern(InvokePattern.Pattern, out object p)) ((InvokePattern)p).Invoke();
            else if (e.TryGetCurrentPattern(SelectionItemPattern.Pattern, out p)) ((SelectionItemPattern)p).Select();
            else if (e.TryGetCurrentPattern(TogglePattern.Pattern, out p)) ((TogglePattern)p).Toggle();
            else throw new InvalidOperationException("「" + name + "」 cannot be clicked");
        }

        public string Value(string id)
        {
            AutomationElement e = ById(id);
            return e.TryGetCurrentPattern(ValuePattern.Pattern, out object p) ? ((ValuePattern)p).Current.Value : e.Current.Name;
        }

        public void Choose(string comboId, string item)
        {
            AutomationElement combo = ById(comboId);
            var expand = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
            expand.Expand();
            AutomationElement option = null;
            WaitUntil(() => (option = combo.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, item)).Cast<AutomationElement>()
                .FirstOrDefault(o => o.Current.ControlType == ControlType.ListItem)) != null, 10, item + " in " + comboId);
            ((SelectionItemPattern)option.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            expand.Collapse();
        }

        /// <summary>Clicks the row of grid <paramref name="gridId"/> whose first column reads <paramref name="firstCell"/>.</summary>
        public void ClickRow(string gridId, string firstCell)
        {
            AutomationElement grid = ById(gridId);
            AutomationElement cell = null;
            WaitUntil(() =>
            {
                foreach (AutomationElement row in grid.FindAll(TreeScope.Children, Condition.TrueCondition))
                {
                    var cells = row.FindAll(TreeScope.Children, Condition.TrueCondition);
                    if (cells.Count > 1 && cells[1].TryGetCurrentPattern(ValuePattern.Pattern, out object v) && ((ValuePattern)v).Current.Value == firstCell)
                    {
                        cell = cells[1];
                        return true;
                    }
                }
                return false;
            }, 30, "row " + firstCell + " in " + gridId);

            // UI Automation has no way to select a DataGridView row, so click it like a user.
            ClickAt(cell);
        }

        // MainTabControl's first tabs, in MainForm.Designer.cs order; SongMgr never hides these. (UI Automation lists
        // only the selected tab and those after it, so it cannot tell a tab's index.)
        static readonly string[] MainTabs = { "歌庫查詢", "加歌頁面", "歌手管理", "歌庫設定", "歌庫維護", "錢櫃資料", "程式設定" };

        /// <summary>
        /// Moves the main tab control one tab towards <paramref name="name"/> with Ctrl+PageDown / Ctrl+PageUp, as a
        /// user can; the caller repeats until it gets there. UI Automation's TabItem cannot be used to pick a tab here.
        /// </summary>
        void SelectTab(string name)
        {
            AutomationElement tabs = ById("MainTabControl");
            int index = Array.IndexOf(MainTabs, name);
            if (index < 0) throw new InvalidOperationException("not a fixed main tab: " + name);
            int current = (int)SendMessage(new IntPtr(tabs.Current.NativeWindowHandle), TCM_GETCURSEL, IntPtr.Zero, IntPtr.Zero);
            tabLog.Add("tab " + current + " -> " + index);
            if (current == index) return;

            tabs.SetFocus();
            byte key = (byte)(current < index ? VK_NEXT : VK_PRIOR);
            keybd_event(VK_CONTROL, 0, 0, IntPtr.Zero);
            keybd_event(key, 0, 0, IntPtr.Zero);
            keybd_event(key, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
        }

        void ClickAt(AutomationElement target)
        {
            IntPtr previous = SetThreadDpiAwarenessContext(new IntPtr(-4)); // per-monitor v2: same pixels as the bounding rectangle
            try
            {
                SetForegroundWindow(new IntPtr(main.Current.NativeWindowHandle));
                System.Windows.Rect r = target.Current.BoundingRectangle;
                SetCursorPos((int)(r.X + r.Width / 2), (int)(r.Y + r.Height / 2));
                Thread.Sleep(150);
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
            }
            finally
            {
                SetThreadDpiAwarenessContext(previous);
            }
        }

        public static void WaitUntil(Func<bool> done, int seconds, string what)
        {
            var sw = Stopwatch.StartNew();
            while (true)
            {
                try
                {
                    if (done()) return;
                }
                catch (ElementNotAvailableException) { }
                catch (COMException) { } // SongMgr's UI thread was busy; ask again
                if (sw.Elapsed.TotalSeconds > seconds) throw new TimeoutException("Gave up waiting for " + what + " after " + seconds + " s");
                Thread.Sleep(250);
            }
        }

        public void Dispose()
        {
            try
            {
                if (process != null && !process.HasExited)
                {
                    process.CloseMainWindow();
                    if (!process.WaitForExit(10000)) process.Kill();
                    process.WaitForExit();
                }
            }
            finally
            {
                process?.Dispose();
                try { TestEnv.DeleteDir(Root); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        const int MOUSEEVENTF_LEFTDOWN = 0x02, MOUSEEVENTF_LEFTUP = 0x04, TCM_GETCURSEL = 0x130B, KEYEVENTF_KEYUP = 0x2;
        const byte VK_CONTROL = 0x11, VK_PRIOR = 0x21, VK_NEXT = 0x22;
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
        [DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr hWnd, out int pid);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder name, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int max);
        const int WM_COMMAND = 0x111, BM_CLICK = 0xF5, IDOK = 1, IDYES = 6, IDNO = 7;
        [DllImport("user32.dll")] static extern void keybd_event(byte key, byte scan, int flags, IntPtr extra);
        [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")] static extern void mouse_event(int flags, int dx, int dy, int data, IntPtr extra);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    }
}

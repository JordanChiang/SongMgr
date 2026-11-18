using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Compare
{
    /// <summary>Runs only when tools\compare-with-release.ps1 sets the 比對 folders; skipped in a normal test run.</summary>
    public sealed class CompareFactAttribute : FactAttribute
    {
        public CompareFactAttribute(string variable = "SONGMGR_COMPARE_OUT")
        {
            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable(variable))) Skip = variable + " is not set (tools\\compare-with-release.ps1)";
        }
    }

    /// <summary>
    /// 新舊版比對: runs the real MainForm (off screen) against a copy of a real 歌庫 and writes what a user would see:
    /// the start-up labels, every 歌庫查詢 / 異常查詢 / 我的最愛 list, the 歌曲分析 list, the 加歌 result and the 歌庫
    /// afterwards. tools\compare-with-release.ps1 builds this folder against a released version as well, then
    /// <see cref="CompareReport"/> diffs the two outputs. Drives the form through its own handlers by reflection, so it
    /// builds against any version that has the same control and handler names.
    /// SONGMGR_COMPARE_DATA: the folder <see cref="CompareData"/> prepared; SONGMGR_COMPARE_OUT: where this version
    /// writes; SONGMGR_COMPARE_CFG: the 歌庫設定檔 (CrazyKTV_SongMgr.cfg) to start with.
    /// </summary>
    [Trait("Category", "Compare")]
    public class LibraryCompare
    {
        static string Data => Environment.GetEnvironmentVariable("SONGMGR_COMPARE_DATA");
        static string Out => Environment.GetEnvironmentVariable("SONGMGR_COMPARE_OUT");
        static string Cfg => Environment.GetEnvironmentVariable("SONGMGR_COMPARE_CFG");

        [CompareFact]
        public void Queries()
        {
            RunSta(() => new Driver("query", Path.Combine(Data, "query.mdb")).Queries());
        }

        /// <summary>加歌 from SONGMGR_COMPARE_ADD, one of the folders <see cref="CompareData"/> made.</summary>
        [CompareFact]
        public void AddSongs()
        {
            string folder = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_ADD") ?? "add_new";
            RunSta(() => new Driver(folder, Path.Combine(Data, "add.mdb")).AddSongs(Path.Combine(Data, folder)));
        }

        static void RunSta(Action body)
        {
            Exception error = null;
            var t = new Thread(() => { try { body(); } catch (Exception ex) { error = ex; } });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            if (error != null) throw new Exception("compare run failed", error);
        }

        sealed class Driver
        {
            readonly string work;
            readonly StreamWriter w;
            readonly string root = Data;
            MainForm form;
            volatile bool watching = true;

            public Driver(string name, string database)
            {
                work = Path.Combine(Out, name + "_work");
                if (Directory.Exists(work)) Directory.Delete(work, true);
                Directory.CreateDirectory(Path.Combine(work, "SongMgr"));
                w = new StreamWriter(Path.Combine(Out, name + ".txt"), false, new UTF8Encoding(false)) { NewLine = "\n" };

                string db = Path.Combine(work, "CrazySong.mdb");
                File.Copy(database, db);
                string mgr = Path.Combine(work, "SongMgr", "CrazySongMgr.mdb");
                File.Copy(Path.Combine(AppContext.BaseDirectory, "SongMgr", "CrazySongMgr.mdb"), mgr);
                string cfg = Path.Combine(work, "CrazyKTV_SongMgr.cfg");
                File.Copy(Cfg, cfg);
                string ktvCfg = Path.Combine(Path.GetDirectoryName(Cfg), "CrazyKTV.cfg");
                if (File.Exists(ktvCfg)) File.Copy(ktvCfg, Path.Combine(work, "CrazyKTV.cfg"));

                foreach (var (k, v) in new[]
                {
                    ("CrazyktvDatabaseFile", db),
                    ("MainCfgEnableAutoUpdate", "False"),
                    ("SongMgrEnableMonitorFolders", "False"),
                })
                {
                    CommonFunc.SaveConfigXmlFile(cfg, k, v);
                }

                SetGlobal("SongMgrCfgFile", cfg);
                SetGlobal("CrazyktvCfgFile", Path.Combine(work, "CrazyKTV.cfg"));
                SetGlobal("CrazyktvSongMgrDatabaseFile", mgr);
                // Only versions before 資料庫結構更新 (2.11.4 and older) keep a 歌庫版本 file.
                if (typeof(Global).GetField("CrazyktvSongDBUpdateFile", GlobalFlags) != null)
                {
                    SetGlobal("CrazyktvSongDBUpdateFile", Path.Combine(work, "SongMgr", "UpdateDB.xml"));
                }
            }

            // ---------- reflection helpers ----------

            // A Global member may be a field or, for a 歌庫設定, a property; a missing one throws rather than being skipped.
            const BindingFlags GlobalFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
            static void SetGlobal(string name, object value)
            {
                FieldInfo field = typeof(Global).GetField(name, GlobalFlags);
                if (field != null) { field.SetValue(null, value); return; }
                (typeof(Global).GetProperty(name, GlobalFlags) ?? throw new MissingMemberException(nameof(Global), name)).SetValue(null, value);
            }
            static object GetGlobal(string name)
            {
                FieldInfo field = typeof(Global).GetField(name, GlobalFlags);
                if (field != null) return field.GetValue(null);
                return (typeof(Global).GetProperty(name, GlobalFlags) ?? throw new MissingMemberException(nameof(Global), name)).GetValue(null);
            }

            T C<T>(string name) where T : class =>
                (T)typeof(MainForm).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public).GetValue(form);

            void Call(string method, params object[] args)
            {
                MethodInfo m = typeof(MainForm).GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                if (m == null) throw new MissingMethodException(method);
                m.Invoke(form, args);
            }

            // ---------- progress ----------

            static string Only => Environment.GetEnvironmentVariable("SONGMGR_COMPARE_ONLY");
            static bool Wanted(string part) => Only == null || Only.Split(',').Contains(part);

            volatile string step = "";
            void Step(string s)
            {
                step = s;
                File.AppendAllText(Path.Combine(Out, "progress.txt"), DateTime.Now.ToString("HH:mm:ss.fff") + " " + s + Environment.NewLine);
            }

            // ---------- pumping ----------

            void Pump(Func<bool> done, int timeoutSec, string what)
            {
                DateTime end = DateTime.Now.AddSeconds(timeoutSec);
                while (!done())
                {
                    Application.DoEvents();
                    Thread.Sleep(15);
                    if (DateTime.Now > end)
                    {
                        w.WriteLine("!! timeout waiting for " + what);
                        w.Flush();
                        return;
                    }
                }
            }

            void Settle(int ms)
            {
                DateTime end = DateTime.Now.AddMilliseconds(ms);
                while (DateTime.Now < end) { Application.DoEvents(); Thread.Sleep(15); }
            }

            /// <summary>Waits until no label on the form has changed for <paramref name="quietMs"/>.</summary>
            void Quiet(int quietMs, int timeoutSec)
            {
                string last = LabelState();
                DateTime since = DateTime.Now, end = DateTime.Now.AddSeconds(timeoutSec);
                while (DateTime.Now < end)
                {
                    Application.DoEvents();
                    Thread.Sleep(50);
                    string now = LabelState();
                    if (now != last) { last = now; since = DateTime.Now; }
                    else if ((DateTime.Now - since).TotalMilliseconds >= quietMs) return;
                }
                w.WriteLine("!! labels never went quiet");
            }

            string LabelState() => string.Join("\u0001", AllControls(form).OfType<Label>().Select(l => l.Text));

            static IEnumerable<Control> AllControls(Control c)
            {
                foreach (Control child in c.Controls)
                {
                    yield return child;
                    foreach (Control d in AllControls(child)) yield return d;
                }
            }

            // ---------- dialogs ----------

            delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
            [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
            [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr l);
            [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);
            [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
            [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
            [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);

            static string Text(IntPtr h) { var sb = new StringBuilder(4096); GetWindowText(h, sb, sb.Capacity); return sb.ToString(); }
            static string Class(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, sb.Capacity); return sb.ToString(); }

            readonly List<string> dialogs = new List<string>();

            /// <summary>Answers every message box: Yes / OK, logging what it said.</summary>
            void WatchDialogs()
            {
                uint me = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                var handled = new HashSet<IntPtr>();
                while (watching)
                {
                    EnumWindows((h, l) =>
                    {
                        GetWindowThreadProcessId(h, out uint pid);
                        if (pid == me && IsWindowVisible(h) && Class(h) == "#32770" && handled.Add(h))
                        {
                            var texts = new List<string>();
                            var buttons = new List<(int Id, string Text)>();
                            EnumChildWindows(h, (c, l2) =>
                            {
                                string cls = Class(c);
                                if (cls == "Static" && Text(c) != "") texts.Add(Text(c));
                                if (cls == "Button") buttons.Add((GetDlgCtrlID(c), Text(c)));
                                return true;
                            }, IntPtr.Zero);
                            int answer = buttons.Any(b => b.Id == 6) ? 6 : buttons.Any(b => b.Id == 1) ? 1 : buttons.Count > 0 ? buttons[0].Id : 2;
                            lock (dialogs) dialogs.Add("[" + Text(h) + "] " + string.Join(" / ", texts).Replace("\r", "").Replace("\n", " ") + " {" + string.Join(",", buttons.Select(b => b.Text)) + "} -> " + answer);
                            PostMessage(h, 0x0111, (IntPtr)answer, IntPtr.Zero);
                        }
                        return true;
                    }, IntPtr.Zero);
                    Thread.Sleep(200);
                }
            }

            void FlushDialogs()
            {
                lock (dialogs)
                {
                    foreach (string d in dialogs) w.WriteLine("DIALOG " + d);
                    dialogs.Clear();
                }
            }

            // ---------- start ----------

            void Start()
            {
                new Thread(WatchDialogs) { IsBackground = true }.Start();
                form = new MainForm
                {
                    StartPosition = FormStartPosition.Manual,
                    Location = new Point(-30000, -30000),
                    ShowInTaskbar = false,
                };
                form.Show();
                var query = C<Button>("SongQuery_Query_Button");
                Pump(() => GetGlobal("CrazyktvDatabaseStatus") is bool b && b && query.Enabled, 600, "startup");
                Quiet(4000, 600);

                w.WriteLine("## 啟動");
                FlushDialogs();
                foreach (Control c in AllControls(form).Where(c => c is Label || c is TextBox || c is ComboBox || c is CheckBox || c is RadioButton || c is Button)
                    .OrderBy(c => c.Name, StringComparer.Ordinal))
                {
                    if (c.Name == "") continue;
                    string state = c is CheckBox cb ? " checked=" + cb.Checked : c is RadioButton rb ? " checked=" + rb.Checked : "";
                    w.WriteLine(c.GetType().Name + " " + c.Name + " = " + Clean(c.Text) + state + (c.Enabled ? "" : " [disabled]"));
                }
                w.Flush();
            }

            /// <summary>Paths that differ between the two runs become &lt;DATA&gt; and &lt;OUT&gt;.</summary>
            string Clean(string s) => (s ?? "").Replace(root, "<DATA>").Replace(Out, "<OUT>").Replace("\r", "").Replace("\n", "\\n");

            void Finish()
            {
                watching = false;
                form.Close();
                form.Dispose();
                w.Flush();
                w.Dispose();
            }

            // ---------- dumping ----------

            void DumpTable(DataTable dt, IEnumerable<string> visibleHeaders)
            {
                if (dt == null) { w.WriteLine("(no list)"); return; }
                if (visibleHeaders != null) w.WriteLine("COLUMNS " + string.Join(",", visibleHeaders));
                var cols = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
                w.WriteLine("FIELDS " + string.Join(",", cols));
                w.WriteLine("ROWS " + dt.DefaultView.Count);
                foreach (DataRowView r in dt.DefaultView)
                {
                    w.WriteLine("R " + string.Join("|", cols.Select(c => Value(r[c]))));
                }
            }

            string Value(object v)
            {
                switch (v)
                {
                    case null: return "";
                    case DBNull _: return "";
                    case DateTime d: return d.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture);
                    case float f: return f.ToString(CultureInfo.InvariantCulture);
                    case double f: return f.ToString(CultureInfo.InvariantCulture);
                    default: return Clean(Convert.ToString(v, CultureInfo.InvariantCulture));
                }
            }

            void DumpGrid(DataGridView grid)
            {
                var headers = grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).Select(c => c.HeaderText);
                DumpTable(grid.DataSource as DataTable, headers);
            }

            void DumpLogs()
            {
                foreach (string name in new[] { "SongLog", "FailureSongLog", "DuplicateSongLog" })
                {
                    object log = GetGlobal(name);
                    DataTable t = (DataTable)log.GetType().GetProperty("Table").GetValue(log);
                    w.WriteLine("LOG " + name + " " + t.Rows.Count);
                    foreach (DataRow r in t.Rows) w.WriteLine("L " + Clean(r[0].ToString()));
                }
            }

            void DumpDatabase(string sql)
            {
                w.WriteLine("SQL " + sql);
                using (DataTable dt = Db.GetDataTable((string)GetGlobal("CrazyktvDatabaseFile"), sql, ""))
                {
                    DumpTable(dt, null);
                }
            }

            // ---------- 歌庫查詢 ----------

            void Select(Control c)
            {
                for (Control p = c; p != null; p = p.Parent)
                {
                    if (p is TabPage page && page.Parent is TabControl tabs) tabs.SelectedTab = page;
                }
                Settle(100);
            }

            void WaitQuery()
            {
                var button = C<Button>("SongQuery_Query_Button");
                var status = C<Label>("SongQuery_QueryStatus_Label");
                Settle(100);
                Pump(() => button.Enabled && !status.Text.StartsWith("正在") && status.Text != "查詢中,請稍待...", 900, "query");
                Settle(300);
            }

            void Result(string title)
            {
                w.WriteLine("## " + title);
                FlushDialogs();
                w.WriteLine("STATUS " + Clean(C<Label>("SongQuery_QueryStatus_Label").Text));
                DumpGrid(C<DataGridView>("SongQuery_DataGridView"));
                w.Flush();
            }

            void Query(int type, string text, string fuzzy)
            {
                if (!Wanted(type.ToString())) return;
                var typeBox = C<ComboBox>("SongQuery_QueryType_ComboBox");
                Step("query " + type + " " + text + " select");
                Select(typeBox);
                Step("query " + type + " set type");
                typeBox.SelectedValue = type;
                Step("query " + type + " settle");
                Settle(100);
                // Picking 全部 / 新進 / 合唱 starts a query by itself: let it finish first.
                WaitQuery();
                Step("query " + type + " set text");
                C<TextBox>("SongQuery_QueryValue_TextBox").Text = text;
                SetGlobal("SongQueryFuzzyQuery", fuzzy);
                Step("query " + type + " click");
                Call("SongQuery_Query_Button_Click", null, EventArgs.Empty);
                Step("query " + type + " wait");
                WaitQuery();
                Step("query " + type + " dump");
                Result("查詢 " + typeBox.Text + " 『" + text + "』 fuzzy=" + fuzzy);
            }

            void QueryEachValue(int type)
            {
                if (!Wanted(type.ToString())) return;
                var typeBox = C<ComboBox>("SongQuery_QueryType_ComboBox");
                Select(typeBox);
                typeBox.SelectedValue = type;
                Settle(100);
                var values = C<ComboBox>("SongQuery_QueryValue_ComboBox");
                int n = values.Items.Count;
                for (int i = 0; i < n; i++)
                {
                    typeBox.SelectedValue = type;
                    Settle(50);
                    values.SelectedIndex = i;
                    Settle(50);
                    WaitQuery();
                    string shown = values.Text;
                    Call("SongQuery_Query_Button_Click", null, EventArgs.Empty);
                    WaitQuery();
                    Result("查詢 " + typeBox.Text + " 『" + shown + "』");
                }
            }

            public void Queries()
            {
                Start();
                string fuzzyDefault = (string)GetGlobal("SongQueryFuzzyQuery");

                // 一般查詢的值: 常見字、大小寫、SQL 特殊字元、頭尾空白; 查不到的也要兩版相同.
                string[] names = { "愛", "love", "LOVE", "我的", "Fallin", "'", "[", "(瑞影)", "%", "_", "愛 ", "一", "天" };
                string[] singers = { "周杰倫", "江蕙", "張學友", "tank", "Tank", "Ella", "翻拍英", "boss", "alicia keys", "Michael孫", "楊潔儀", "張" };
                foreach (string fuzzy in new[] { "True", "False" })
                {
                    foreach (string s in names) Query(1, s, fuzzy);
                    foreach (string s in singers) Query(2, s, fuzzy);
                }
                SetGlobal("SongQueryFuzzyQuery", fuzzyDefault);
                foreach (string s in new[] { "123", "0001", "950", "95001", "2" }) Query(3, s, fuzzyDefault);
                foreach (int t in new[] { 4, 5, 6 }) { Query(t, "", fuzzyDefault); Query(t, "愛", fuzzyDefault); }
                QueryEachValue(7);
                QueryEachValue(8);
                QueryEachValue(9);
                foreach (string s in new[] { "", "1" }) Query(10, s, fuzzyDefault);

                // 異常查詢: every item; the handler runs on SelectedIndexChanged.
                if (Wanted("ex")) {
                var ex = C<ComboBox>("SongQuery_ExceptionalQuery_ComboBox");
                Select(ex);
                int count = ex.Items.Count;
                foreach (int i in Enumerable.Range(1, count - 1).Concat(new[] { 0 }))
                {
                    Step("ex " + i);
                    ex.SelectedIndex = i;
                    WaitQuery();
                    Result("異常查詢 " + ex.Text);
                }
                }

                // 我的最愛: needs the combo focused, as when a user picks from it.
                if (Wanted("fav")) {
                var fav = C<ComboBox>("SongQuery_FavoriteQuery_ComboBox");
                Select(fav);
                var main = C<TabControl>("MainTabControl");
                main.SelectedTab = main.TabPages.Cast<TabPage>().First(p => p.Name == "SongQuery_TabPage");
                Settle(200);
                int favCount = fav.Items.Count;
                w.WriteLine("## 我的最愛清單 " + favCount);
                for (int i = 0; i < favCount; i++) w.WriteLine("U " + fav.GetItemText(fav.Items[i]));
                for (int i = 0; i < favCount; i++)
                {
                    fav.Focus();
                    Settle(50);
                    if (fav.SelectedIndex == i) { fav.SelectedIndex = (i + 1) % favCount; WaitQuery(); fav.Focus(); }
                    fav.SelectedIndex = i;
                    WaitQuery();
                    Result("我的最愛 " + fav.Text + " focused=" + fav.Focused);
                }
                }

                Finish();
            }

            // ---------- 加歌 ----------

            public void AddSongs(string folder)
            {
                Start();
                var grid = C<DataGridView>("SongAdd_DataGridView");
                var tip = C<Label>("SongAdd_Tooltip_Label");
                var add = C<Button>("SongAdd_Add_Button");
                Select(grid);
                var main = C<TabControl>("MainTabControl");
                main.SelectedTab = main.TabPages.Cast<TabPage>().First(p => p.Name == "SongAdd_TabPage");
                Settle(200);

                Type argsType = typeof(MainForm).Assembly.GetTypes().First(t => t.Name == "ElevatedDragDropArgs");
                object args = Activator.CreateInstance(argsType);
                argsType.GetProperty("HWnd").SetValue(args, grid.Handle);
                argsType.GetProperty("Files").SetValue(args, new List<string> { folder });

                // A fresh log, as the 加歌 tab shows only this run.
                foreach (string name in new[] { "SongLog", "FailureSongLog", "DuplicateSongLog" })
                {
                    object log = GetGlobal(name);
                    log.GetType().GetMethod("Clear").Invoke(log, null);
                }

                DateTime t0 = DateTime.Now;
                Call("SongAdd_ElevatedDragDrop", null, args);
                Pump(() => grid.DataSource != null && !tip.Text.StartsWith("正在分析") && tip.Text != "", 3600, "analysis");
                Settle(500);
                w.WriteLine("## 歌曲分析 (" + (int)(DateTime.Now - t0).TotalSeconds + "s)");
                FlushDialogs();
                w.WriteLine("TIP " + System.Text.RegularExpressions.Regex.Replace(Clean(tip.Text), @"共花費 \d+ 秒", "共花費 N 秒"));
                w.WriteLine("ADD enabled=" + add.Enabled + " text=" + add.Text);
                DumpGrid(grid);
                DumpLogs();
                w.Flush();

                if (add.Enabled)
                {
                    t0 = DateTime.Now;
                    Call("SongAdd_Add_Button_Click", null, EventArgs.Empty);
                    Pump(() => tip.Text.Contains("完成加歌") || tip.Text.Contains("完成更新"), 1800, "add");
                    Quiet(3000, 600);
                    w.WriteLine("## 加入歌庫 (" + (int)(DateTime.Now - t0).TotalSeconds + "s)");
                    FlushDialogs();
                    w.WriteLine("TIP " + System.Text.RegularExpressions.Regex.Replace(Clean(tip.Text), @"共花費 \d+ 秒", "共花費 N 秒"));
                    w.WriteLine("ADD enabled=" + add.Enabled + " text=" + add.Text);
                    DumpGrid(grid);
                    DumpLogs();
                    w.Flush();
                }

                w.WriteLine("## 歌庫 ktv_Song");
                DumpDatabase("select Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState, Song_ReplayGain from ktv_Song order by Song_Id");
                w.WriteLine("## 歌庫 ktv_Singer");
                DumpDatabase("select Singer_Id, Singer_Name, Singer_Type, Singer_Spell, Singer_Strokes, Singer_SpellNum, Singer_PenStyle from ktv_Singer order by Singer_Name, Singer_Id");
                w.WriteLine("## 歌庫 ktv_Favorite");
                DumpDatabase("select User_Id, Song_Id from ktv_Favorite order by User_Id, Song_Id");
                Finish();
            }
        }
    }
}

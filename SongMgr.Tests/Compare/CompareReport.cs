using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace SongMgr.Tests.Compare
{
    /// <summary>
    /// 新舊版比對的報告: reads what <see cref="LibraryCompare"/> wrote for the released version (SONGMGR_COMPARE_BASE)
    /// and for this one (SONGMGR_COMPARE_NEW), section by section, and writes what differs to SONGMGR_COMPARE_REPORT.
    /// Rows are compared on the fields both versions have; 加歌日期 of songs just analysed or added is the time of the
    /// run and is left out. Same rows in another order is reported as such.
    /// </summary>
    [Trait("Category", "Compare")]
    public class CompareReport
    {
        const int Examples = 10;

        sealed class Section
        {
            public readonly List<string> Other = new List<string>();
            public List<string> Fields = new List<string>();
            public string Columns;
            public readonly List<string> Rows = new List<string>();
            public readonly Dictionary<string, List<string>> Logs = new Dictionary<string, List<string>>();
            public readonly List<string> LogOrder = new List<string>();
        }

        [CompareFact("SONGMGR_COMPARE_REPORT")]
        public void Write()
        {
            string baseDir = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_BASE");
            string newDir = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_NEW");
            string report = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_REPORT");

            var sb = new StringBuilder();
            int same = 0, all = 0;
            foreach (string file in Directory.GetFiles(baseDir, "*.txt").Select(Path.GetFileName).Where(f => f != "progress.txt").OrderBy(f => f, StringComparer.Ordinal))
            {
                string other = Path.Combine(newDir, file);
                sb.Append("######## ").Append(file).Append("\r\n");
                if (!File.Exists(other))
                {
                    sb.Append("  only the released version wrote this\r\n");
                    continue;
                }
                var a = Parse(Path.Combine(baseDir, file));
                var b = Parse(other);
                foreach (string title in a.Keys.Concat(b.Keys.Where(k => !a.ContainsKey(k))))
                {
                    all++;
                    if (!a.TryGetValue(title, out Section x) || !b.TryGetValue(title, out Section y))
                    {
                        sb.Append("### ").Append(title).Append(": only in ").Append(x == null ? "新版" : "舊版").Append("\r\n");
                        continue;
                    }
                    List<string> notes = Compare(title, x, y);
                    if (notes.Count == 0) { same++; continue; }
                    sb.Append("### ").Append(title).Append("\r\n");
                    foreach (string n in notes) sb.Append(n).Append("\r\n");
                }
            }
            sb.Insert(0, "=== " + same + " / " + all + " sections identical\r\n");
            File.WriteAllText(report, sb.ToString(), new UTF8Encoding(false));
        }

        static List<string> Compare(string title, Section a, Section b)
        {
            var notes = new List<string>();
            foreach (string x in Minus(a.Other, b.Other)) notes.Add("  舊版: " + x);
            foreach (string x in Minus(b.Other, a.Other)) notes.Add("  新版: " + x);
            if (a.Columns != b.Columns)
            {
                notes.Add("  欄位 舊版: " + (a.Columns ?? "(無清單)"));
                notes.Add("  欄位 新版: " + (b.Columns ?? "(無清單)"));
            }
            var common = a.Fields.Where(b.Fields.Contains).ToList();
            if (title.Contains("歌曲分析") || title.Contains("加入歌庫")) common.Remove("Song_CreatDate");
            List<string> ra = Project(a.Rows, a.Fields, common), rb = Project(b.Rows, b.Fields, common);
            if (!ra.SequenceEqual(rb))
            {
                List<string> onlyA = Minus(ra, rb), onlyB = Minus(rb, ra);
                notes.Add($"  rows: 舊版 {ra.Count}, 新版 {rb.Count}; 只在舊版 {onlyA.Count}, 只在新版 {onlyB.Count}" +
                    (onlyA.Count == 0 && onlyB.Count == 0 ? " (同樣的列, 順序不同)" : ""));
                if (common.Count > 0) notes.Add("  fields: " + string.Join(",", common));
                notes.AddRange(onlyA.Take(Examples).Select(r => "  - " + r));
                notes.AddRange(onlyB.Take(Examples).Select(r => "  + " + r));
                if (onlyA.Count == 0 && onlyB.Count == 0)
                {
                    int i = Enumerable.Range(0, ra.Count).First(k => ra[k] != rb[k]);
                    notes.Add($"  第一個順序不同的列 {i}: {ra[i]}  <>  {rb[i]}");
                }
            }
            foreach (string name in a.LogOrder.Concat(b.LogOrder).Distinct())
            {
                List<string> la = a.Logs.TryGetValue(name, out var l1) ? l1 : new List<string>();
                List<string> lb = b.Logs.TryGetValue(name, out var l2) ? l2 : new List<string>();
                List<string> onlyA = Minus(la, lb), onlyB = Minus(lb, la);
                if (onlyA.Count == 0 && onlyB.Count == 0) continue;
                notes.Add($"  {name}: 舊版 {la.Count}, 新版 {lb.Count}; 只在舊版 {onlyA.Count}, 只在新版 {onlyB.Count}");
                notes.AddRange(onlyA.Take(Examples).Select(r => "  - " + r));
                notes.AddRange(onlyB.Take(Examples).Select(r => "  + " + r));
            }
            return notes;
        }

        /// <summary>The items of <paramref name="a"/> left after taking out one of each item of <paramref name="b"/>.</summary>
        static List<string> Minus(List<string> a, List<string> b)
        {
            var counts = new Dictionary<string, int>();
            foreach (string x in b) counts[x] = counts.TryGetValue(x, out int n) ? n + 1 : 1;
            var left = new List<string>();
            foreach (string x in a)
            {
                if (counts.TryGetValue(x, out int n) && n > 0) counts[x] = n - 1;
                else left.Add(x);
            }
            return left;
        }

        static List<string> Project(List<string> rows, List<string> fields, List<string> keep)
        {
            int[] idx = keep.Select(f => fields.IndexOf(f)).ToArray();
            return rows.Select(r =>
            {
                string[] parts = r.Split('|');
                return parts.Length == fields.Count ? string.Join("|", idx.Select(i => parts[i])) : r;
            }).ToList();
        }

        static Dictionary<string, Section> Parse(string path)
        {
            var sections = new Dictionary<string, Section>();
            var order = new List<string>();
            Section cur = null;
            string log = null;
            foreach (string line in File.ReadLines(path, Encoding.UTF8))
            {
                if (line.StartsWith("## "))
                {
                    string title = line.Substring(3);
                    // 「歌曲分析 (12s)」: how long it took is not a difference.
                    int p = title.LastIndexOf(" (", StringComparison.Ordinal);
                    if (p > 0 && title.EndsWith("s)") && title.Substring(p + 2, title.Length - p - 4).All(char.IsDigit)) title = title.Substring(0, p);
                    string key = title;
                    for (int n = 2; sections.ContainsKey(key); n++) key = title + " #" + n;
                    cur = new Section();
                    sections[key] = cur;
                    log = null;
                    continue;
                }
                if (cur == null) continue;
                if (line.StartsWith("R ")) cur.Rows.Add(line.Substring(2));
                else if (line.StartsWith("FIELDS ")) cur.Fields = line.Substring(7).Split(',').ToList();
                else if (line.StartsWith("COLUMNS ")) cur.Columns = line.Substring(8);
                else if (line.StartsWith("ROWS ")) { }
                else if (line.StartsWith("LOG "))
                {
                    log = line.Substring(4).Split(' ')[0];
                    cur.Logs[log] = new List<string>();
                    cur.LogOrder.Add(log);
                }
                else if (line.StartsWith("L ") && log != null) cur.Logs[log].Add(line.Substring(2));
                else cur.Other.Add(line);
            }
            return sections;
        }
    }
}

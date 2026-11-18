using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 歌曲編號配置: hands out 歌曲編號 within one operation. Built from the 語系編號區段, the digit count and the
    /// 歌曲編號 already in use; <see cref="Take"/> hands out a 缺號 first, then the next number past the highest.
    /// Build one per operation (加歌 batch, 歌庫查詢 edit, 編碼位數校正...) and drop it afterwards: it does not see
    /// songs added by anything else. Safe to call from several threads.
    /// </summary>
    public sealed class SongIdAllocator
    {
        readonly object sync = new object();
        readonly List<string> langs;
        readonly int[] starts;
        readonly int[] ends;
        readonly int digits;
        readonly string format;
        readonly HashSet<string> used;
        readonly int[] max;
        readonly SortedSet<int>[] gaps;

        /// <param name="langs">The 語系 in 語系編號區段 order (<see cref="Global.CrazyktvSongLangList"/>).</param>
        /// <param name="startIds">Each 語系's first 歌曲編號; a 語系's range ends just before the next one's start, the last at 99999 / 999999.</param>
        /// <param name="digitCount">5 or 6.</param>
        /// <param name="usedIds">Every Song_Id already in the 歌庫, as stored.</param>
        public SongIdAllocator(IEnumerable<string> langs, IEnumerable<int> startIds, int digitCount, IEnumerable<string> usedIds)
        {
            if (digitCount != 5 && digitCount != 6) throw new ArgumentOutOfRangeException(nameof(digitCount), "歌曲編號只有 5 或 6 位數");
            this.langs = langs.ToList();
            starts = startIds.ToArray();
            if (starts.Length < this.langs.Count) throw new ArgumentException("每個語系都要有起始編號", nameof(startIds));

            digits = digitCount;
            format = "D" + digitCount;
            int top = digitCount == 5 ? 99999 : 999999;
            used = new HashSet<string>(usedIds);

            int n = this.langs.Count;
            ends = new int[n];
            max = new int[n];
            gaps = new SortedSet<int>[n];
            for (int i = 0; i < n; i++)
            {
                ends[i] = (i < n - 1) ? starts[i + 1] - 1 : top;
                max[i] = starts[i] - 1;
                gaps[i] = new SortedSet<int>();
            }

            var numbers = new List<int>();
            foreach (string id in used)
            {
                if (id.Length == digits && int.TryParse(id, out int value)) numbers.Add(value);
            }

            // Each 語系 on its own, so a mistyped 語系編號區段 that overlaps another gives the same numbers as before.
            for (int i = 0; i < n; i++)
            {
                foreach (int value in numbers)
                {
                    if (value >= starts[i] && value <= ends[i] && value > max[i]) max[i] = value;
                }
                for (int value = starts[i]; value < max[i]; value++)
                {
                    if (!used.Contains(value.ToString(format))) gaps[i].Add(value);
                }
            }
        }

        /// <summary>The 歌庫's own allocator: the current 語系編號區段 and digit count, and every Song_Id in ktv_Song.</summary>
        public static SongIdAllocator FromLibrary()
        {
            var ids = new List<string>();
            using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, "select Song_Id from ktv_Song", ""))
            {
                foreach (DataRow row in dt.Rows) ids.Add(row["Song_Id"].ToString());
            }
            return FromSettings(ids);
        }

        /// <summary>The current 語系編號區段 and digit count, with <paramref name="usedIds"/> as the numbers in use.</summary>
        public static SongIdAllocator FromSettings(IEnumerable<string> usedIds)
        {
            IEnumerable<int> startIds = Regex.Split(Global.SongMgrLangCode, ",", RegexOptions.None).Select(s => Convert.ToInt32(s));
            return new SongIdAllocator(Global.CrazyktvSongLangList, startIds, Global.SongMgrMaxDigitCode == "1" ? 5 : 6, usedIds);
        }

        /// <summary>
        /// Whether <paramref name="lang"/> is a 語系 of this allocator. <see cref="Take"/>, <see cref="Peek"/>,
        /// <see cref="TryClaim"/> and <see cref="Remaining"/> throw for any other, 「未知」 included.
        /// </summary>
        public bool HasLang(string lang) => langs.Contains(lang);

        /// <summary>The number <see cref="Take"/> would hand out next, without taking it.</summary>
        public string Peek(string lang)
        {
            lock (sync)
            {
                int i = IndexOf(lang);
                int value = gaps[i].Count > 0 ? gaps[i].Min : max[i] + 1;
                return value.ToString(format);
            }
        }

        /// <summary>Hands out the lowest 缺號 of <paramref name="lang"/>, or else the number after its highest.</summary>
        public string Take(string lang)
        {
            lock (sync)
            {
                int i = IndexOf(lang);
                int value;
                if (gaps[i].Count > 0)
                {
                    value = gaps[i].Min;
                    gaps[i].Remove(value);
                }
                else
                {
                    value = ++max[i];
                }
                string id = value.ToString(format);
                used.Add(id);
                return id;
            }
        }

        /// <summary>
        /// Claims a 自訂編號 for <paramref name="lang"/>. False, and nothing changes, when the number has the wrong digit
        /// count, is in use, or lies outside the 語系's range. A 6-digit 歌庫 also accepts 000000-099999 for any 語系.
        /// </summary>
        public bool TryClaim(string id, string lang)
        {
            lock (sync)
            {
                int i = IndexOf(lang);
                if (id == null || id.Length != digits || used.Contains(id) || !int.TryParse(id, out int value)) return false;

                bool legacyRange = digits == 6 && value >= 0 && value <= 99999;
                if (!legacyRange && (value < starts[i] || value > ends[i])) return false;

                if (legacyRange)
                {
                    int owner = RangeOf(value);
                    if (owner >= 0) gaps[owner].Remove(value);
                }
                else
                {
                    gaps[i].Remove(value);
                    if (value > max[i])
                    {
                        for (int v = max[i] + 1; v < value; v++)
                        {
                            if (!used.Contains(v.ToString(format))) gaps[i].Add(v);
                        }
                        max[i] = value;
                    }
                }
                used.Add(id);
                return true;
            }
        }

        /// <summary>How many 歌曲編號 of <paramref name="lang"/>'s range are not in use.</summary>
        public int Remaining(string lang)
        {
            lock (sync)
            {
                int i = IndexOf(lang);
                int inUse = 0;
                foreach (string id in used)
                {
                    if (id.Length == digits && int.TryParse(id, out int value) && value >= starts[i] && value <= ends[i]) inUse++;
                }
                return ends[i] - starts[i] + 1 - inUse;
            }
        }

        int IndexOf(string lang)
        {
            int i = langs.IndexOf(lang);
            if (i < 0) throw new ArgumentException("不是歌庫的語系: " + lang, nameof(lang));
            return i;
        }

        int RangeOf(int value)
        {
            for (int i = 0; i < langs.Count; i++)
            {
                if (value >= starts[i] && value <= ends[i]) return i;
            }
            return -1;
        }
    }
}

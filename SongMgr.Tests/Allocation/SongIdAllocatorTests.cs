using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Allocation
{
    /// <summary>
    /// 歌曲編號配置: <see cref="SongIdAllocator"/> on its own, and side by side with the 歌曲編號 code it replaced
    /// (<see cref="LegacySongId"/>). Deliberate change: a 語系 with no songs no longer hands out its start number twice.
    /// </summary>
    public class SongIdAllocatorTests
    {
        static readonly List<string> Langs = new List<string> { "國語", "台語", "粵語", "日語", "英語", "客語", "原住民語", "韓語", "兒歌", "其它" };
        const string Code5 = "10001,20001,30001,40001,50001,60001,70001,80001,90001,95001";
        const string Code6 = "100001,200001,300001,400001,500001,600001,700001,800001,900001,950001";
        // A 6-digit 語系編號區段 whose first 語系 starts below 100000, so a 自訂編號 in 000000-099999 falls inside it.
        const string Code6Low = "1,50001,100001,200001,300001,400001,500001,600001,700001,800001";

        static List<int> Starts(string code) => code.Split(',').Select(int.Parse).ToList();

        static SongIdAllocator Allocator(string code, int digits, params string[] used) => new SongIdAllocator(Langs, Starts(code), digits, used);

        // ---------- side by side with the old code ----------

        public static IEnumerable<object[]> Libraries()
        {
            foreach (int seed in Enumerable.Range(1, 40))
            {
                yield return new object[] { seed, Code5, 5 };
                yield return new object[] { seed, Code6, 6 };
                yield return new object[] { seed, Code6Low, 6 };
            }
        }

        [Theory]
        [MemberData(nameof(Libraries))]
        public void Matches_the_old_code_operation_for_operation(int seed, string code, int digits)
        {
            var rnd = new Random(seed);
            List<int> starts = Starts(code);
            string fmt = "D" + digits;

            // Every 語系 has at least one song, so the empty-語系 fix does not come into play.
            var library = new List<string>();
            for (int i = 0; i < Langs.Count; i++)
            {
                library.Add((starts[i] + rnd.Next(0, 5)).ToString(fmt));
                for (int k = 0; k < 40; k++)
                {
                    if (rnd.Next(3) == 0) library.Add((starts[i] + k).ToString(fmt));
                }
            }
            // Numbers of the other digit count, as in a 歌庫 that needs 編碼位數校正.
            for (int k = 0; k < 5; k++) library.Add((starts[0] + k).ToString(digits == 5 ? "D6" : "D5"));
            library = library.Distinct().ToList();

            var legacy = new LegacySongId(new List<string>(Langs), code, digits, library);
            legacy.Load();
            legacy.GetRemainingSongIdCount(digits);
            var module = new SongIdAllocator(Langs, starts, digits, library);

            for (int i = 0; i < Langs.Count; i++)
            {
                Assert.Equal(legacy.RemainingSongIdCountList[i], module.Remaining(Langs[i]));
            }

            for (int step = 0; step < 120; step++)
            {
                string lang = Langs[rnd.Next(Langs.Count)];
                int li = Langs.IndexOf(lang);
                string where = $"seed {seed}, step {step}, {lang}";

                switch (rnd.Next(4))
                {
                    case 0:
                        Assert.True(legacy.GetNextSongId(lang, false) == module.Peek(lang), "Peek " + where);
                        break;
                    case 1:
                    {
                        string id = legacy.GetNextSongId(lang, true);
                        legacy.SongDataIdList.Add(id);
                        Assert.True(id == module.Take(lang), "Take " + where);
                        break;
                    }
                    default:
                    {
                        string custom = CustomId(rnd, starts, li, digits);
                        var (id, accepted) = legacy.AddSongId(custom, lang);
                        legacy.SongDataIdList.Add(id);
                        bool claimed = module.TryClaim(custom, lang);
                        string moduleId = claimed ? custom : module.Take(lang);
                        Assert.True(accepted == claimed, $"TryClaim {custom} {where}");
                        Assert.True(id == moduleId, $"自訂編號 {custom} {where}: {id} vs {moduleId}");
                        break;
                    }
                }
            }
        }

        static string CustomId(Random rnd, List<int> starts, int li, int digits)
        {
            string fmt = "D" + digits;
            switch (rnd.Next(6))
            {
                case 0: return (starts[li] + rnd.Next(0, 70)).ToString(digits == 5 ? "D6" : "D5"); // wrong digit count
                case 1: return (starts[(li + 1) % starts.Count] + rnd.Next(0, 20)).ToString(fmt); // another 語系's range
                case 2: return digits == 6 ? rnd.Next(0, 100).ToString("D6") : (starts[li] + rnd.Next(0, 70)).ToString(fmt);
                default: return (starts[li] + rnd.Next(0, 70)).ToString(fmt);
            }
        }

        [Fact]
        public void 轉換為6位數_gives_the_same_numbers_as_before()
        {
            var rnd = new Random(7);
            var songs = Enumerable.Range(0, 200).Select(k => ((k + 1).ToString("D5"), Langs[rnd.Next(Langs.Count)])).ToList();

            List<string> before = LegacySongId.ConvertTo6(Langs, Code6, songs);
            var module = new SongIdAllocator(Langs, Starts(Code6), 6, Enumerable.Empty<string>());
            List<string> after = songs.Select(s => module.Take(s.Item2)).ToList();

            Assert.Equal(before, after);
        }

        // ---------- the deliberate change ----------

        [Fact]
        public void Empty_語系_hands_out_its_start_number_once()
        {
            var library = new List<string> { "100001", "100002" }; // only 國語 has songs

            var legacy = new LegacySongId(new List<string>(Langs), Code6, 6, library);
            legacy.Load();
            Assert.Equal("200001", legacy.GetNextSongId("台語", true));
            Assert.Equal("200001", legacy.GetNextSongId("台語", true)); // the old duplicate

            var module = Allocator(Code6, 6, library.ToArray());
            Assert.Equal("200001", module.Take("台語"));
            Assert.Equal("200002", module.Take("台語"));
        }

        // ---------- on its own ----------

        [Fact]
        public void Take_fills_缺號_lowest_first_then_goes_past_the_highest()
        {
            var a = Allocator(Code6, 6, "100001", "100003", "100006");
            Assert.Equal(new[] { "100002", "100004", "100005", "100007", "100008" }, Enumerable.Range(0, 5).Select(_ => a.Take("國語")));
        }

        [Fact]
        public void Peek_does_not_take()
        {
            var a = Allocator(Code5, 5, "10001", "10003");
            Assert.Equal("10002", a.Peek("國語"));
            Assert.Equal("10002", a.Peek("國語"));
            Assert.Equal("10002", a.Take("國語"));
            Assert.Equal("10004", a.Peek("國語"));
        }

        [Fact]
        public void Numbers_of_the_other_digit_count_are_ignored()
        {
            var a = Allocator(Code6, 6, "10001", "10002", "100001");
            Assert.Equal(100000 - 1, a.Remaining("國語")); // 100001-200000, only 100001 counts
            Assert.Equal("100002", a.Take("國語"));
        }

        [Fact]
        public void Each_語系_stays_in_its_own_range()
        {
            var a = Allocator(Code5, 5, "10001", "19999");
            Assert.Equal("10002", a.Take("國語"));
            Assert.Equal("20001", a.Take("台語"));
            Assert.Equal("95001", a.Take("其它"));
        }

        [Theory]
        [InlineData("10005", "國語", true)]
        [InlineData("19999", "國語", true)]
        [InlineData("20001", "國語", false)]  // 台語's range
        [InlineData("10001", "國語", false)]  // in use
        [InlineData("100005", "國語", false)] // 6 digits in a 5-digit 歌庫
        [InlineData("1000", "國語", false)]
        [InlineData("1000a", "國語", false)]
        [InlineData("99999", "其它", true)]  // the last 語系 runs to 99999
        [InlineData("", "國語", false)]
        public void TryClaim_rules_in_a_5_digit_歌庫(string id, string lang, bool expected)
        {
            var a = Allocator(Code5, 5, "10001");
            Assert.Equal(expected, a.TryClaim(id, lang));
        }

        [Theory]
        [InlineData("012345", "國語", true)]  // 000000-099999 is open to every 語系 of a 6-digit 歌庫
        [InlineData("012345", "其它", true)]
        [InlineData("999999", "其它", true)]
        [InlineData("200001", "國語", false)]
        [InlineData("10005", "國語", false)]
        public void TryClaim_rules_in_a_6_digit_歌庫(string id, string lang, bool expected)
        {
            var a = Allocator(Code6, 6, "100001");
            Assert.Equal(expected, a.TryClaim(id, lang));
        }

        [Fact]
        public void A_claimed_number_is_not_handed_out_again()
        {
            var a = Allocator(Code6, 6, "100001", "100003");
            Assert.True(a.TryClaim("100002", "國語"));
            Assert.False(a.TryClaim("100002", "國語"));
            Assert.Equal("100004", a.Take("國語"));
        }

        [Fact]
        public void Claiming_past_the_highest_leaves_the_numbers_between_as_缺號()
        {
            var a = Allocator(Code6, 6, "100001");
            Assert.True(a.TryClaim("100005", "國語"));
            Assert.Equal(new[] { "100002", "100003", "100004", "100006" }, Enumerable.Range(0, 4).Select(_ => a.Take("國語")));
        }

        [Fact]
        public void Remaining_counts_the_unused_numbers_of_the_range()
        {
            var a = Allocator(Code5, 5, "10001", "10002", "100001");
            Assert.Equal(9998, a.Remaining("國語"));
            Assert.Equal(4999, a.Remaining("其它"));
            a.Take("國語");
            Assert.Equal(9997, a.Remaining("國語"));
        }

        [Fact]
        public void Unknown_語系_throws()
        {
            var a = Allocator(Code6, 6);
            Assert.Throws<ArgumentException>(() => a.Take("未知"));
        }

        [Fact]
        public void Parallel_takes_never_repeat()
        {
            var a = Allocator(Code6, 6, "100001", "100005", "100009");
            var ids = new ConcurrentBag<string>();
            Parallel.For(0, 2000, _ => ids.Add(a.Take("國語")));
            Assert.Equal(2000, ids.Distinct().Count());
        }

        // ---------- reading the 歌庫 ----------

        [Fact]
        public void FromLibrary_reads_the_Song_Id_of_an_Access_and_a_SQLite_歌庫()
        {
            var saved = TestEnv.Capture();
            string root = TestEnv.NewTempDir("allocator");
            try
            {
                TestEnv.UseScratchDatabases(root);
                Global.CrazyktvSongLangList = new List<string>(Langs);
                Global.SongMgrLangCode = Code6;
                Global.SongMgrMaxDigitCode = "2";
                using (var conn = DbHelper.Access.OpenConn(Global.CrazyktvDatabaseFile))
                {
                    foreach (string id in new[] { "100001", "100003", "200001" })
                    {
                        using (var cmd = new OleDbCommand("insert into ktv_Song (Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState) values (@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s)", conn))
                        {
                            foreach (string v in new[] { id, "國語", "0", "歌手" + id, "歌名" + id, "1", "", "100", "2", "0", "12.5", "2024/01/02 03:04:05", id + ".mpg", @"D:\KTV\", "ㄍㄇ", "42", "12", "丨一", "0" })
                                cmd.Parameters.AddWithValue("@p", v);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                void Check()
                {
                    var a = SongIdAllocator.FromLibrary();
                    Assert.Equal("100002", a.Take("國語"));
                    Assert.Equal("100004", a.Take("國語"));
                    Assert.Equal("200002", a.Peek("台語"));
                    Assert.Equal(100000 - 1, a.Remaining("台語"));
                }

                Check();

                string db = Path.Combine(root, "CrazySong.db");
                var results = SqliteExport.Run(new[] { new SqliteExport.Job(Global.CrazyktvDatabaseFile, db, SqliteExport.CrazySongDBVer) });
                Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());
                Global.CrazyktvDatabaseFile = db;
                Check();
            }
            finally
            {
                System.Data.SQLite.SQLiteConnection.ClearAllPools();
                TestEnv.Restore(saved);
                TestEnv.DeleteDir(root);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;
using Song = CrazyKTV_SongMgr.CashboxIdPlan.Song;

namespace SongMgr.Tests.CashboxTests
{
    /// <summary>
    /// 套用錢櫃編號: the plan of moves, played out step by step so no two songs ever share a number, and the writes
    /// on a scratch 歌庫 with its 我的最愛.
    /// </summary>
    public class CashboxIdPlanTests
    {
        static readonly List<string> Langs = new List<string> { "國語", "台語", "粵語", "日語", "英語", "客語", "原住民語", "韓語", "兒歌", "其它" };
        static readonly List<int> Code6 = new List<int> { 100001, 200001, 300001, 400001, 500001, 600001, 700001, 800001, 900001, 950001 };

        static SongIdAllocator Allocator(IEnumerable<string> used) => new SongIdAllocator(Langs, Code6, 6, used);

        static Song S(string id, string name = null) => new Song(id, id.StartsWith("2") ? "台語" : "國語", "歌手" + id, name ?? "歌" + id);

        /// <summary>Plays the moves out; every move must start where the song is and land on a free number.</summary>
        static Dictionary<Song, string> PlayOut(IReadOnlyList<Song> library, CashboxIdPlan.Result plan)
        {
            var where = library.ToDictionary(s => s, s => s.Id);
            foreach (var m in plan.Moves)
            {
                Assert.Equal(where[m.Song], m.From);
                Assert.False(where.Values.Contains(m.To), $"{m.Song} moves onto {m.To}, which is taken");
                where[m.Song] = m.To;
            }
            Assert.Equal(where.Count, where.Values.Distinct().Count());
            return where;
        }

        [Fact]
        public void A_free_錢櫃_number_is_taken_directly()
        {
            Song a = S("100005");
            var library = new[] { a, S("100001") };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "100002") }, Allocator);

            Assert.Equal("100002", PlayOut(library, plan)[a]);
            Assert.Single(plan.Moves);
        }

        [Fact]
        public void The_song_on_the_number_gets_a_new_one_first()
        {
            Song a = S("100005"), b = S("100001");
            var library = new[] { a, b };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "100001") }, Allocator);

            var end = PlayOut(library, plan);
            Assert.Equal("100001", end[a]);
            Assert.Equal("100002", end[b]); // the first 缺號 of 國語
            Assert.Equal(CashboxIdPlan.Reason.MadeRoom, plan.Moves.First().Why);
        }

        [Fact]
        public void A_song_with_its_own_錢櫃_number_is_not_moved_twice()
        {
            // The old code moved b out of the way to a new number, then looked for b under its old number and moved a instead.
            Song a = S("100005"), b = S("100001");
            var library = new[] { a, b };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "100001"), (b, "100003") }, Allocator);

            var end = PlayOut(library, plan);
            Assert.Equal("100001", end[a]);
            Assert.Equal("100003", end[b]);
            Assert.All(plan.Moves, m => Assert.Equal(CashboxIdPlan.Reason.Cashbox, m.Why));
        }

        [Fact]
        public void Two_songs_swap_through_a_spare_number()
        {
            Song a = S("100001"), b = S("100002");
            var library = new[] { a, b };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "100002"), (b, "100001") }, Allocator);

            var end = PlayOut(library, plan);
            Assert.Equal("100002", end[a]);
            Assert.Equal("100001", end[b]);
            Assert.Contains(plan.Moves, m => m.Why == CashboxIdPlan.Reason.Spare);
        }

        [Fact]
        public void A_new_number_is_never_one_of_the_錢櫃_numbers()
        {
            // 100002 is a 缺號, but c takes it: b has to go elsewhere.
            Song a = S("100010"), b = S("100001"), c = S("100011");
            var library = new[] { a, b, c };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "100001"), (c, "100002") }, Allocator);

            var end = PlayOut(library, plan);
            Assert.Equal("100001", end[a]);
            Assert.Equal("100002", end[c]);
            Assert.DoesNotContain(end[b], new[] { "100001", "100002", "100010", "100011" });
        }

        [Fact]
        public void A_made_room_song_keeps_its_own_語系()
        {
            Song a = S("100005"), b = S("200001");
            var library = new[] { a, b };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "200001") }, Allocator);

            Assert.Equal("200002", PlayOut(library, plan)[b]);
        }

        [Fact]
        public void One_song_per_錢櫃_number_and_one_number_per_song()
        {
            Song a = S("100005"), c = S("100006"), d = S("100007");
            var library = new[] { a, c, d };
            var plan = CashboxIdPlan.Plan(library, new[] { (c, "100001"), (a, "100001"), (d, "100002"), (d, "100003") }, Allocator);

            var end = PlayOut(library, plan);
            Assert.Equal("100001", end[a]); // lowest Song_Id first
            Assert.Equal("100006", end[c]);
            Assert.Equal("100002", end[d]); // lowest 錢櫃 number first
            Assert.Equal(2, plan.Skipped.Count);
        }

        [Fact]
        public void A_song_already_on_its_number_keeps_it()
        {
            Song a = S("100001"), b = S("100005");
            var library = new[] { a, b };
            var plan = CashboxIdPlan.Plan(library, new[] { (a, "100001"), (b, "100001") }, Allocator);

            Assert.Empty(plan.Moves);
            Assert.Single(plan.Skipped);
        }

        // ---------- writing it ----------

        [Fact]
        public void Apply_moves_the_songs_and_their_我的最愛_or_nothing_at_all()
        {
            var saved = TestEnv.Capture();
            string root = TestEnv.NewTempDir("cashbox");
            try
            {
                TestEnv.UseScratchDatabases(root);
                string db = Global.CrazyktvDatabaseFile;
                Song a = S("100001"), b = S("100002"), c = S("100009");
                var library = new[] { a, b, c };
                foreach (Song s in library) InsertSong(db, s);
                Execute(db, "insert into ktv_Favorite (User_Id, Song_Id) values ('0001', '100001')");
                Execute(db, "insert into ktv_Favorite (User_Id, Song_Id) values ('0002', '100009')");

                // a and b swap, c takes a free number.
                var plan = CashboxIdPlan.Plan(library, new[] { (a, "100002"), (b, "100001"), (c, "100003") }, Allocator);
                using (var store = SongLibraryStore.OpenFile(db))
                {
                    Assert.True(CashboxIdPlan.Apply(store, plan.Moves, out Exception error), error?.ToString());
                }
                Assert.Equal(new[] { "100001:" + b.Name, "100002:" + a.Name, "100003:" + c.Name }, Songs(db));
                Assert.Equal(new[] { "0001:100002", "0002:100003" }, Favorites(db));

                // A plan made from an out-of-date list changes nothing: c's move would go through, b is not on 100008.
                Song cNow = new Song("100003", c.Lang, c.Singer, c.Name), bStale = new Song("100008", b.Lang, b.Singer, b.Name);
                var stale = CashboxIdPlan.Plan(new[] { cNow, bStale }, new[] { (cNow, "100005"), (bStale, "100006") }, Allocator);
                Assert.Equal(new[] { cNow, bStale }, stale.Moves.Select(m => m.Song));
                using (var store = SongLibraryStore.OpenFile(db))
                {
                    Assert.False(CashboxIdPlan.Apply(store, stale.Moves, out Exception error));
                    Assert.NotNull(error);
                }
                Assert.Equal(new[] { "100001:" + b.Name, "100002:" + a.Name, "100003:" + c.Name }, Songs(db));
                Assert.Equal(new[] { "0001:100002", "0002:100003" }, Favorites(db));
            }
            finally
            {
                TestEnv.Restore(saved);
                TestEnv.DeleteDir(root);
            }
        }

        static void InsertSong(string db, Song s)
        {
            Execute(db, "insert into ktv_Song (Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState) values (@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s)",
                s.Id, s.Lang, "0", s.Singer, s.Name, "1", "", "100", "2", "0", "0.5", "2024/01/02 03:04:05", s.Name + ".mpg", @"D:\KTV\", "ㄍㄇ", "42", "12", "丨一", "0");
        }

        static void Execute(string db, string sql, params string[] values)
        {
            using (var conn = DbHelper.Access.OpenConn(db))
            using (var cmd = new OleDbCommand(sql, conn))
            {
                foreach (string v in values) cmd.Parameters.AddWithValue("@p", v);
                cmd.ExecuteNonQuery();
            }
        }

        static List<string> Songs(string db) =>
            Db.GetDataTable(db, "select Song_Id, Song_SongName from ktv_Song order by Song_Id", "").AsEnumerable().Select(r => r[0] + ":" + r[1]).ToList();

        static List<string> Favorites(string db) =>
            Db.GetDataTable(db, "select User_Id, Song_Id from ktv_Favorite order by User_Id", "").AsEnumerable().Select(r => r[0] + ":" + r[1]).ToList();
    }
}

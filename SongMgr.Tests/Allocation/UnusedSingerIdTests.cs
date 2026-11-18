using System;
using System.Collections.Generic;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Allocation
{
    /// <summary>
    /// <see cref="CommonFunc.GetUnusedSingerId"/>: the 歌手編號 below the highest one that no singer uses, which
    /// 新增歌手, 加歌, 歌庫轉換 and 重建歌庫歌手 hand out (first entry first) before going past the highest.
    /// </summary>
    public class UnusedSingerIdTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("singerid");
        readonly Dictionary<string, object> saved = TestEnv.Capture();

        public UnusedSingerIdTests()
        {
            TestEnv.UseScratchDatabases(root);
            using (var conn = Db.OpenConn(Global.CrazyktvDatabaseFile))
            using (var cmd = Db.Command("delete * from ktv_Singer", conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(root);
        }

        static void AddSingers(params int[] ids)
        {
            using (SongLibraryStore store = SongLibraryStore.Open(SingerTable.Library))
            {
                foreach (int id in ids) store.InsertSinger(SingerTable.Library, id.ToString(), "歌手" + id, "0");
            }
        }

        [Theory]
        [InlineData(3000)]
        [InlineData(300000)]
        public void Lists_every_unused_number_below_the_highest_once_lowest_first(int highest)
        {
            AddSingers(1, 5, highest);

            List<string> unused = CommonFunc.GetUnusedSingerId("ktv_Singer", Global.CrazyktvDatabaseFile);

            List<string> expected = Enumerable.Range(2, highest - 2).Where(i => i != 5).Select(i => i.ToString()).ToList();
            Assert.Equal(expected.Count, unused.Count);
            Assert.DoesNotContain(null, unused);
            Assert.True(expected.SequenceEqual(unused), "缺號 should come lowest first; got " + string.Join(",", unused.Take(10)) + "...");
        }

        [Fact]
        public void No_singers_or_no_gaps_gives_an_empty_list()
        {
            Assert.Empty(CommonFunc.GetUnusedSingerId("ktv_Singer", Global.CrazyktvDatabaseFile));

            AddSingers(1, 2, 3);
            Assert.Empty(CommonFunc.GetUnusedSingerId("ktv_Singer", Global.CrazyktvDatabaseFile));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Jobs
{
    /// <summary>變更歌曲路徑的作業本體, 直接對測試歌庫 (CrazySongEmpty.mdb 的複本) 執行, 不經過畫面.</summary>
    public class SongPathChangeTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string root = TestEnv.NewTempDir("pathchange");

        public SongPathChangeTests()
        {
            TestEnv.UseScratchDatabases(root);
            TestEnv.ResetLogs();
        }

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(root);
        }

        static void Seed(params (string Id, string Path)[] songs)
        {
            using (SongLibraryStore store = SongLibraryStore.Open())
            {
                foreach (var s in songs)
                {
                    store.InsertSong(new SongRecord
                    {
                        Id = s.Id, Lang = "國語", SingerType = "0", Singer = "歌手" + s.Id, SongName = "歌" + s.Id, Track = "1",
                        Volume = "100", WordCount = "1", PlayCount = "0", MB = "1", CreatDate = "2024/01/02 03:04:05",
                        FileName = s.Id + ".mpg", Path = s.Path, Spell = "ㄍ", SpellNum = "4", SongStroke = "10", PenStyle = "丨", PlayState = "0",
                    }, SongColumns.AllButReplayGain);
                }
            }
        }

        static Dictionary<string, string> Paths()
        {
            return Db.GetDataTable(Global.CrazyktvDatabaseFile, "select Song_Id, Song_Path from ktv_Song", "")
                .AsEnumerable().ToDictionary(r => r.Field<string>("Song_Id"), r => r.Field<string>("Song_Path"));
        }

        static DataTable Songs() => Db.GetDataTable(Global.CrazyktvDatabaseFile, "select Song_Id, Song_Path from ktv_Song order by Song_Id", "");

        static SongPathChange.Change Change(string id, string path) => new SongPathChange.Change(id, path);

        // 超過 Song_Path 欄位長度, Access 會拒絕寫入
        static readonly string TooLong = @"E:\" + new string('長', 300) + @"\";

        [Theory]
        [InlineData(@"d:/ktv//國語", @"d:\ktv\國語\")]
        [InlineData(@"  D:\KTV  ", @"D:\KTV\")]
        [InlineData(@"\\nas\ktv", @"\\nas\ktv\")]
        [InlineData(" ", "")]
        public void NormalizeDirectory(string input, string expected)
        {
            Assert.Equal(expected, SongPathChange.NormalizeDirectory(input));
        }

        [Fact]
        public void Plan_ReplacesTheSourceFolderIgnoringCase_AndSkipsOtherSongs()
        {
            Seed(("000001", @"D:\KTV\國語\"), ("000002", @"d:\ktv\台語\老歌\"), ("000003", @"E:\Other\"));

            var changes = SongPathChange.Plan(Songs(), @"D:\KTV", @"F:\新歌庫\", JobContext.ForTest());

            Assert.Equal(new[] { "000001=F:\\新歌庫\\國語\\", "000002=F:\\新歌庫\\台語\\老歌\\" }, changes.Select(c => c.SongId + "=" + c.Path));
        }

        [Fact]
        public void Apply_WritesAllBatches_AndCountsSuccesses()
        {
            Seed(("000001", @"D:\A\"), ("000002", @"D:\A\"), ("000003", @"D:\A\"), ("000004", @"D:\A\"), ("000005", @"D:\A\"));
            var ctx = JobContext.ForTest();

            using (SongLibraryStore store = SongLibraryStore.Open())
            {
                SongPathChange.Apply(store, SongPathChange.Plan(Songs(), @"D:\A\", @"E:\B\", ctx), ctx, batchSize: 2);
            }

            Assert.Equal(5, ctx.Counts[0]);
            Assert.Equal(0, ctx.Counts[1]);
            Assert.All(Paths().Values, p => Assert.Equal(@"E:\B\", p));
        }

        [Fact]
        public void Apply_ErrorInLastBatch_RollsItBack_AndCountsItAsFailed()
        {
            Seed(("000001", @"D:\A\"), ("000002", @"D:\A\"), ("000003", @"D:\A\"), ("000004", @"D:\A\"), ("000005", @"D:\A\"));
            var ctx = JobContext.ForTest();
            var changes = new[]
            {
                Change("000001", @"E:\B\"), Change("000002", @"E:\B\"),  // 第 1 批: Commit
                Change("000003", @"E:\B\"), Change("000004", TooLong),    // 第 2 批: 有一筆失敗, 仍 Commit
                Change("000005", @"E:\B\"),                               // 最後一批: 因為有錯而 Rollback
            };

            using (SongLibraryStore store = SongLibraryStore.Open())
            {
                SongPathChange.Apply(store, changes, ctx, batchSize: 2);
            }

            var paths = Paths();
            Assert.Equal(@"E:\B\", paths["000001"]);
            Assert.Equal(@"E:\B\", paths["000002"]);
            Assert.Equal(@"E:\B\", paths["000003"]);
            Assert.Equal(@"D:\A\", paths["000004"]);
            Assert.Equal(@"D:\A\", paths["000005"]);
            Assert.Equal(3, ctx.Counts[0]);  // 只算真的寫進歌庫的
            Assert.Equal(2, ctx.Counts[1]);  // 寫入失敗 1 首 + 被 Rollback 1 首
            Assert.Contains(TestEnv.SongLog(), l => l.Contains("最後一批 1 首歌曲的變更已取消"));
        }

        [Fact]
        public void Apply_NothingToChange_LeavesNoOpenTransaction()
        {
            Seed(("000001", @"D:\A\"));
            var ctx = JobContext.ForTest();

            using (SongLibraryStore store = SongLibraryStore.Open())
            {
                SongPathChange.Apply(store, new SongPathChange.Change[0], ctx);
                store.UpdateSong(new SongRecord { Id = "000001", Path = @"Z:\" }, SongColumns.Path);
            }

            Assert.Equal(@"Z:\", Paths()["000001"]);
            Assert.Equal(0, ctx.Counts[0] + ctx.Counts[1]);
        }

        [Fact]
        public void Apply_ReportsProgressForEverySong()
        {
            Seed(("000001", @"D:\A\"), ("000002", @"D:\A\"));
            var reports = new List<string>();
            var ctx = JobContext.ForTest(reports.Add);

            using (SongLibraryStore store = SongLibraryStore.Open())
            {
                SongPathChange.Apply(store, new[] { Change("000001", @"E:\"), Change("000002", @"E:\") }, ctx);
            }

            Assert.Equal("正在轉換第 2 首歌曲的歌曲路徑資料,請稍待...", reports.Last());
        }
    }
}

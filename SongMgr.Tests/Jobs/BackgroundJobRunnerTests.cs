using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Jobs
{
    /// <summary>
    /// <see cref="BackgroundJobRunner"/> 不需要 WinForms: 鎖定畫面換成記錄事件, UI 執行緒換成測試自己抽的佇列.
    /// </summary>
    public class BackgroundJobRunnerTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly FakeUi ui = new FakeUi();

        public BackgroundJobRunnerTests()
        {
            TestEnv.ResetLogs();
        }

        public void Dispose() => TestEnv.Restore(saved);

        BackgroundJobRunner NewRunner(Func<DateTime> clock = null) => new BackgroundJobRunner(ui.SetLocked, ui.Post, clock);

        [Fact]
        public void Success_LocksRunsShowsSummaryCompletesThenUnlocks()
        {
            var runner = NewRunner();
            bool done = false;

            bool started = runner.TryStart(new BackgroundJob
            {
                Name = "測試",
                Status = ui.Show,
                Work = ctx => { ctx.Counts.Increment(0, 3); ctx.Counts.Increment(1); },
                Summary = (c, secs) => "成功 " + c[0] + " 失敗 " + c[1],
                Completed = () => { ui.Events.Add("completed"); done = true; },
            });

            Assert.True(started);
            Assert.True(runner.IsBusy);
            ui.PumpUntil(() => !runner.IsBusy);

            Assert.True(done);
            Assert.Equal(new[] { "lock", "status:成功 3 失敗 1", "completed", "unlock" }, ui.Events);
        }

        [Fact]
        public void WorkThrows_LogsShowsFailureAndStillUnlocks()
        {
            var runner = NewRunner();
            bool summaryCalled = false;

            runner.TryStart(new BackgroundJob
            {
                Name = "變更歌曲路徑",
                Status = ui.Show,
                Work = ctx => throw new InvalidOperationException("磁碟已滿"),
                Summary = (c, secs) => { summaryCalled = true; return ""; },
                Completed = () => ui.Events.Add("completed"),
            });
            ui.PumpUntil(() => !runner.IsBusy);

            Assert.False(summaryCalled);
            Assert.Equal(new[] { "lock", "status:變更歌曲路徑作業失敗: 磁碟已滿", "completed", "unlock" }, ui.Events);
            Assert.Contains(TestEnv.SongLog(), l => l.EndsWith("【變更歌曲路徑】發生錯誤: 磁碟已滿"));
        }

        [Fact]
        public void CompletedThrows_StillUnlocks()
        {
            var runner = NewRunner();

            runner.TryStart(new BackgroundJob { Name = "收尾", Work = ctx => { }, Completed = () => throw new Exception("壞掉") });
            ui.PumpUntil(() => !runner.IsBusy);

            Assert.Equal("unlock", ui.Events.Last());
            Assert.Contains(TestEnv.SongLog(), l => l.EndsWith("【收尾】發生錯誤: 壞掉"));
        }

        [Fact]
        public void SecondJobWhileRunning_IsRejected()
        {
            var runner = NewRunner();
            var release = new ManualResetEventSlim();
            bool secondRan = false;

            runner.TryStart(new BackgroundJob { Name = "第一件", Work = ctx => release.Wait(5000) });
            bool second = runner.TryStart(new BackgroundJob { Name = "第二件", Work = ctx => secondRan = true });
            release.Set();
            ui.PumpUntil(() => !runner.IsBusy);

            Assert.False(second);
            Assert.False(secondRan);
            Assert.Equal(1, ui.Events.Count(e => e == "lock"));
        }

        [Fact]
        public void LegacyLock_BlocksJobs_AndUnlockRunsWaitingActions()
        {
            var runner = NewRunner();
            var ran = new List<string>();

            runner.SetBusy(true);   // 還沒遷移的程式呼叫 Common_SwitchSetUI(false)
            Assert.False(runner.TryStart(new BackgroundJob { Name = "監視", Work = ctx => { } }));
            runner.RunWhenIdle(() => ran.Add("更新"));
            ui.PumpPending();
            Assert.Empty(ran);

            runner.SetBusy(false);  // Common_SwitchSetUI(true)
            ui.PumpPending();
            Assert.Equal(new[] { "更新" }, ran);
        }

        [Fact]
        public void RunWhenIdle_WhenIdle_RunsImmediately()
        {
            var runner = NewRunner();
            bool ran = false;
            runner.RunWhenIdle(() => ran = true);
            Assert.True(ran);
        }

        [Fact]
        public void RunWhenIdle_DuringJob_RunsAfterUnlock()
        {
            var runner = NewRunner();
            var release = new ManualResetEventSlim();

            runner.TryStart(new BackgroundJob { Name = "工作", Work = ctx => release.Wait(5000) });
            runner.RunWhenIdle(() => ui.Events.Add("idle"));
            release.Set();
            ui.PumpUntil(() => ui.Events.Contains("idle"));

            Assert.Equal(new[] { "lock", "unlock", "idle" }, ui.Events);
        }

        [Fact]
        public void RunWhenIdle_ActionStartingAJob_LeavesTheRestWaiting()
        {
            var runner = NewRunner();
            var release = new ManualResetEventSlim();
            var ran = new List<string>();

            runner.TryStart(new BackgroundJob { Name = "工作", Work = ctx => { } });
            runner.RunWhenIdle(() => { ran.Add("a"); runner.TryStart(new BackgroundJob { Name = "下一件", Work = c => release.Wait(5000) }); });
            runner.RunWhenIdle(() => ran.Add("b"));
            ui.PumpUntil(() => ran.Count > 0);
            ui.PumpPending();
            Assert.Equal(new[] { "a" }, ran);

            release.Set();
            ui.PumpUntil(() => ran.Count == 2);
            Assert.Equal(new[] { "a", "b" }, ran);
        }

        [Fact]
        public void Summary_GetsElapsedSecondsFromClock()
        {
            DateTime now = new DateTime(2026, 1, 1, 12, 0, 0);
            var runner = NewRunner(() => now);
            long seen = -1;

            runner.TryStart(new BackgroundJob { Name = "計時", Work = ctx => now = now.AddSeconds(7.6), Summary = (c, secs) => { seen = secs; return ""; } });
            ui.PumpUntil(() => !runner.IsBusy);

            Assert.Equal(7, seen);
        }

        [Fact]
        public void LegacyHelpers_WritingGlobalTotalList_LandInTheJobsCounts()
        {
            var runner = NewRunner();
            Global.TotalList = new List<int> { 9, 9, 9, 9 };
            int seen = -1;

            runner.TryStart(new BackgroundJob
            {
                Name = "舊函式",
                Work = ctx => { Global.TotalList[2]++; Global.TotalList[4] += 2; },
                Summary = (c, secs) => { seen = c[2] * 10 + c[4]; return ""; },
            });
            ui.PumpUntil(() => !runner.IsBusy);

            Assert.Equal(12, seen);
        }

        [Fact]
        public void CountsCopyFrom_AdoptsAReplacedLegacyList()
        {
            var counts = JobContext.ForTest().Counts;
            counts.CopyFrom(new List<int> { 4, 5, 6 });
            Assert.Equal(new[] { 4, 5, 6, 0, 0 }, Enumerable.Range(0, JobCounts.Size).Select(i => counts[i]));
        }

        [Fact]
        public void Progress_IsThrottled_AndShowsTheLatestText()
        {
            var runner = NewRunner();
            var reported = new ManualResetEventSlim();
            var release = new ManualResetEventSlim();

            runner.TryStart(new BackgroundJob
            {
                Name = "進度",
                Status = ui.Show,
                Work = ctx =>
                {
                    for (int i = 1; i <= 2000; i++) ctx.Report("第 " + i + " 首");
                    reported.Set();
                    release.Wait(5000);
                },
                Summary = (c, secs) => "完成",
            });

            Assert.True(reported.Wait(5000));
            // 合併後排著的最後一筆在節流間隔後才送到 UI 佇列
            ui.PumpUntil(() => ui.Events.Contains("status:第 2000 首"));
            release.Set();
            ui.PumpUntil(() => !runner.IsBusy);

            var statuses = ui.Events.Where(e => e.StartsWith("status:")).ToList();
            Assert.InRange(statuses.Count, 2, 4);
            Assert.Equal("status:第 2000 首", statuses[statuses.Count - 2]);
            Assert.Equal("status:完成", statuses.Last());
        }

        [Fact]
        public void Progress_ScheduledAfterTheJobEnds_DoesNotOverwriteTheSummary()
        {
            var runner = NewRunner();

            runner.TryStart(new BackgroundJob
            {
                Name = "進度",
                Status = ui.Show,
                Work = ctx => { ctx.Report("一"); ctx.Report("二"); },
                Summary = (c, secs) => "完成",
            });
            ui.PumpUntil(() => !runner.IsBusy);
            Thread.Sleep(BackgroundJobRunner.ProgressInterval * 3);
            ui.PumpPending();

            Assert.Equal("status:完成", ui.Events.Last(e => e.StartsWith("status:")));
        }

        /// <summary>假的畫面: 記錄鎖定與狀態文字; Post 進佇列, 由測試執行緒 (當作 UI 執行緒) 抽出執行.</summary>
        sealed class FakeUi
        {
            readonly BlockingCollection<Action> queue = new BlockingCollection<Action>();
            public readonly List<string> Events = new List<string>();

            public void SetLocked(bool locked) => Events.Add(locked ? "lock" : "unlock");
            public void Show(string text) => Events.Add("status:" + text);
            public void Post(Action action) => queue.Add(action);

            public void PumpUntil(Func<bool> done, int timeoutMs = 5000)
            {
                var sw = Stopwatch.StartNew();
                while (!done())
                {
                    if (sw.ElapsedMilliseconds > timeoutMs) throw new TimeoutException("UI 佇列等不到預期的狀態");
                    if (queue.TryTake(out Action action, 50)) action();
                }
            }

            public void PumpPending()
            {
                while (queue.TryTake(out Action action)) action();
            }
        }
    }
}

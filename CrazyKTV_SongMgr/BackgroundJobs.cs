using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 一件背景工作: 作業本體在背景執行緒跑, 只碰 <see cref="JobContext"/>; 其餘 (鎖定畫面, 互斥, 計時, 例外, 摘要) 由 <see cref="BackgroundJobRunner"/> 負責.
    /// </summary>
    sealed class BackgroundJob
    {
        /// <summary>作業名稱, 例外記錄用: 「【變更歌曲路徑】發生錯誤: …」.</summary>
        public string Name = "";

        /// <summary>進度與摘要顯示在哪 (UI 執行緒). 可為 null.</summary>
        public Action<string> Status;

        /// <summary>作業本體 (背景執行緒).</summary>
        public Action<JobContext> Work;

        /// <summary>完成摘要: 計數與花費秒數 → 顯示文字. 作業失敗時不呼叫. 可為 null.</summary>
        public Func<JobCounts, long, string> Summary;

        /// <summary>收尾 (UI 執行緒), 成功或失敗都會呼叫, 之後才解鎖畫面. 可為 null.</summary>
        public Action Completed;
    }

    /// <summary>作業本體拿到的東西: 自己的計數, 以及回報進度的出口.</summary>
    sealed class JobContext
    {
        private readonly Action<string> report;

        internal JobContext(JobCounts counts, Action<string> report)
        {
            Counts = counts;
            this.report = report;
        }

        public JobCounts Counts { get; }

        /// <summary>回報進度文字. 可以在每一筆資料呼叫, 畫面更新頻率由 runner 控制.</summary>
        public void Report(string text) => report(text);

        /// <summary>給測試用: 不經過 runner 的 context, 進度交給 <paramref name="report"/>.</summary>
        public static JobContext ForTest(Action<string> report = null) => new JobContext(new JobCounts(), report ?? (_ => { }));
    }

    /// <summary>
    /// 一件背景工作的計數, 以編號存取 (對應原本的 <c>Global.TotalList[n]</c>), 各編號的意義由作業自己決定.
    /// </summary>
    sealed class JobCounts
    {
        public const int Size = 5;

        // 遷移期間 runner 讓 Global.TotalList 指向這份 List, 還沒改的共用函式照樣能累加
        internal readonly List<int> Values = new List<int>(new int[Size]);

        public int this[int index]
        {
            get { lock (Values) return Values[index]; }
        }

        public void Increment(int index, int by = 1)
        {
            lock (Values) Values[index] += by;
        }

        /// <summary>遷移期間用: 還沒改的程式換掉了 Global.TotalList 時, 從那份讀回計數.</summary>
        internal void CopyFrom(IReadOnlyList<int> legacy)
        {
            if (ReferenceEquals(legacy, Values)) return;
            lock (Values)
            {
                for (int i = 0; i < Size; i++) Values[i] = i < legacy.Count ? legacy[i] : 0;
            }
        }
    }

    /// <summary>
    /// 背景工作的執行者: 一次只跑一件工作, 期間鎖定畫面; 計時, 接住例外並記錄, 顯示摘要, 最後一定解鎖.
    /// 還沒遷移的程式直接呼叫 Common_SwitchSetUI, 也透過 <see cref="SetBusy"/> 共用同一個「忙碌中」.
    /// </summary>
    sealed class BackgroundJobRunner
    {
        /// <summary>進度最多每隔這麼久更新一次畫面.</summary>
        public static readonly TimeSpan ProgressInterval = TimeSpan.FromMilliseconds(100);

        private readonly Action<bool> setUiLocked;
        private readonly Action<Action> post;
        private readonly Func<DateTime> clock;
        private readonly Queue<Action> idleActions = new Queue<Action>();
        private volatile bool busy;

        /// <param name="setUiLocked">true 鎖定畫面, false 解鎖 (UI 執行緒).</param>
        /// <param name="post">把動作排到 UI 執行緒執行.</param>
        /// <param name="clock">計時用, 測試可替換.</param>
        public BackgroundJobRunner(Action<bool> setUiLocked, Action<Action> post, Func<DateTime> clock = null)
        {
            this.setUiLocked = setUiLocked;
            this.post = post;
            this.clock = clock ?? (() => DateTime.Now);
        }

        /// <summary>有背景工作在進行 (runner 的工作, 或還沒遷移的程式鎖定了畫面).</summary>
        public bool IsBusy => busy;

        /// <summary>
        /// 還沒遷移的程式鎖定 / 解鎖畫面時呼叫 (Common_SwitchSetUI). 由忙碌轉為閒置時, 執行 <see cref="RunWhenIdle"/> 排著的動作.
        /// </summary>
        public void SetBusy(bool value)
        {
            bool wasBusy = busy;
            busy = value;
            if (wasBusy && !value) post(DrainIdle);
        }

        /// <summary>
        /// 開始一件工作 (UI 執行緒). 已有工作在進行時不開始, 傳回 false; 由計時器觸發的工作就等下一次再試.
        /// </summary>
        public bool TryStart(BackgroundJob job)
        {
            if (busy) return false;

            SetBusy(true);
            setUiLocked(true);

            JobCounts counts = new JobCounts();
            Global.TotalList = counts.Values;
            ProgressThrottle progress = new ProgressThrottle(job.Status, post, clock);
            JobContext ctx = new JobContext(counts, progress.Report);
            DateTime start = clock();

            Task.Run(() =>
            {
                Exception error = null;
                try
                {
                    job.Work(ctx);
                }
                catch (Exception ex)
                {
                    error = ex;
                    Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][BackgroundJob {job.Name} EXCEPTION] {ex}");
                    Global.SongLog.Add("【" + job.Name + "】發生錯誤: " + ex.Message);
                }
                long seconds = (long)(clock() - start).TotalSeconds;
                post(() => Finish(job, counts, seconds, error, progress));
            });
            return true;
        }

        /// <summary>沒有背景工作時立即執行; 否則等工作結束 (畫面解鎖) 後才執行.</summary>
        public void RunWhenIdle(Action action)
        {
            if (!busy)
            {
                action();
                return;
            }
            lock (idleActions) idleActions.Enqueue(action);
        }

        private void Finish(BackgroundJob job, JobCounts counts, long seconds, Exception error, ProgressThrottle progress)
        {
            progress.Stop();
            try
            {
                string text = null;
                if (error != null)
                {
                    text = job.Name + "作業失敗: " + error.Message;
                }
                else if (job.Summary != null)
                {
                    text = job.Summary(counts, seconds);
                }
                if (text != null) job.Status?.Invoke(text);
                job.Completed?.Invoke();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][BackgroundJob {job.Name} COMPLETION EXCEPTION] {ex}");
                Global.SongLog.Add("【" + job.Name + "】發生錯誤: " + ex.Message);
            }
            finally
            {
                setUiLocked(false);
                SetBusy(false);
            }
        }

        private void DrainIdle()
        {
            while (!busy)
            {
                Action next;
                lock (idleActions)
                {
                    if (idleActions.Count == 0) return;
                    next = idleActions.Dequeue();
                }
                next();
            }
        }

        /// <summary>
        /// 合併進度更新: 同時最多一筆排著等畫面更新, 相隔至少 <see cref="ProgressInterval"/>, 顯示的永遠是最新的一筆.
        /// 工作結束後排著的更新不再顯示, 以免蓋掉摘要.
        /// </summary>
        private sealed class ProgressThrottle
        {
            private readonly Action<string> status;
            private readonly Action<Action> post;
            private readonly Func<DateTime> clock;
            private string latest;
            private int scheduled;
            private long lastShownTicks = DateTime.MinValue.Ticks;
            private volatile bool stopped;

            public ProgressThrottle(Action<string> status, Action<Action> post, Func<DateTime> clock)
            {
                this.status = status;
                this.post = post;
                this.clock = clock;
            }

            public void Report(string text)
            {
                if (status == null || stopped) return;
                Volatile.Write(ref latest, text);
                if (Interlocked.CompareExchange(ref scheduled, 1, 0) != 0) return;

                TimeSpan wait = ProgressInterval - (clock() - new DateTime(Interlocked.Read(ref lastShownTicks)));
                if (wait <= TimeSpan.Zero) post(Show);
                else Task.Delay(wait).ContinueWith(_ => post(Show));
            }

            public void Stop() => stopped = true;

            private void Show()
            {
                Interlocked.Exchange(ref lastShownTicks, clock().Ticks);
                Volatile.Write(ref scheduled, 0);
                if (!stopped) status(Volatile.Read(ref latest));
            }
        }
    }
}

using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>下載更新檔時顯示進度, 可取消. DialogResult.OK 表示更新檔已備妥 (PackageRoot).</summary>
    public class AppUpdateProgressForm : Form
    {
        private readonly AppUpdater.UpdatePlan plan;
        private readonly string appDir;
        private readonly CancellationTokenSource cts = new CancellationTokenSource();
        private readonly Label status = new Label();
        private readonly ProgressBar bar = new ProgressBar();
        private readonly Button cancel = new Button();

        public string PackageRoot { get; private set; }
        public Exception Error { get; private set; }

        public AppUpdateProgressForm(AppUpdater.UpdatePlan plan, string appDir)
        {
            this.plan = plan;
            this.appDir = appDir;

            Text = "下載 CrazyKTV 加歌程式 " + plan.Release.Tag;
            Font = new Font("微軟正黑體", 10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(420, 120);

            status.SetBounds(16, 14, 388, 24);
            status.Text = "正在下載更新檔...";
            bar.SetBounds(16, 42, 388, 22);
            bar.Maximum = 1000;
            cancel.SetBounds(314, 76, 90, 30);
            cancel.Text = "取消";
            cancel.Click += (s, e) => { cancel.Enabled = false; cts.Cancel(); };

            Controls.AddRange(new Control[] { status, bar, cancel });
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Task.Run(() =>
            {
                try
                {
                    PackageRoot = AppUpdater.Prepare(plan, appDir, ReportProgress, cts.Token);
                }
                catch (Exception ex)
                {
                    Error = ex;
                }
                BeginInvoke((Action)(() =>
                {
                    DialogResult = PackageRoot != null ? DialogResult.OK : DialogResult.Cancel;
                    Close();
                }));
            });
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 使用者按 X: 取消下載, 等背景工作收尾後才真正關閉
            if (PackageRoot == null && Error == null && e.CloseReason == CloseReason.UserClosing)
            {
                cts.Cancel();
                cancel.Enabled = false;
                e.Cancel = true;
                return;
            }
            base.OnFormClosing(e);
        }

        private long lastReport;

        private void ReportProgress(long done, long total)
        {
            long now = Environment.TickCount64;
            if (now - lastReport < 100 && done != total) return;
            lastReport = now;
            BeginInvoke((Action)(() =>
            {
                if (total > 0)
                {
                    bar.Value = (int)Math.Min(1000, done * 1000 / total);
                    status.Text = "正在下載更新檔... " + AppUpdater.FormatMB(done) + " / " + AppUpdater.FormatMB(total);
                }
                else
                {
                    status.Text = "正在下載更新檔... " + AppUpdater.FormatMB(done);
                }
                if (total > 0 && done >= total) status.Text = "正在解壓縮並檢查更新檔...";
            }));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) cts.Dispose();
            base.Dispose(disposing);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using CrazyKTV_SongMgr;
using Xunit;

[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace SongMgr.Tests.Support
{
    /// <summary>
    /// Shared setup for tests that touch the static <see cref="Global"/> state:
    /// temp folders, scratch copies of the bundled .mdb fixtures, and log readers.
    /// </summary>
    internal static class TestEnv
    {
        static TestEnv()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static string FixtureDir => Path.Combine(AppContext.BaseDirectory, "Fixtures");

        public static string NewTempDir(string prefix)
        {
            string dir = Path.Combine(Path.GetTempPath(), "SongMgrTests", prefix + "_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static void DeleteDir(string dir)
        {
            if (!Directory.Exists(dir)) return;
            foreach (string f in Directory.GetFiles(dir, "*", SearchOption.AllDirectories))
            {
                File.SetAttributes(f, FileAttributes.Normal);
            }
            Directory.Delete(dir, true);
        }

        /// <summary>Points Global at scratch copies of CrazySongEmpty.mdb / CrazySongMgr.mdb in <paramref name="dir"/>.</summary>
        public static void UseScratchDatabases(string dir)
        {
            string song = Path.Combine(dir, "CrazySong.mdb");
            string mgr = Path.Combine(dir, "CrazySongMgr.mdb");
            File.Copy(Path.Combine(FixtureDir, "CrazySongEmpty.mdb"), song, true);
            File.Copy(Path.Combine(FixtureDir, "CrazySongMgr.mdb"), mgr, true);
            File.SetAttributes(song, FileAttributes.Normal);
            File.SetAttributes(mgr, FileAttributes.Normal);
            Global.CrazyktvDatabaseFile = song;
            Global.CrazyktvSongMgrDatabaseFile = mgr;
            Global.PhoneticsWordList = new List<string>();
        }

        /// <summary>Resets the operation log, failure and duplicate tables to empty.</summary>
        public static void ResetLogs()
        {
            Global.SongLog.Reset();
            Global.FailureSongLog.Reset();
            Global.DuplicateSongLog.Reset();
        }

        public static List<string> SongLog() => Read(Global.SongLog.Table);
        public static List<string> FailureLog() => Read(Global.FailureSongLog.Table);
        public static List<string> DuplicateLog() => Read(Global.DuplicateSongLog.Table);

        static List<string> Read(DataTable dt)
        {
            return dt.AsEnumerable().Select(r => r[1] + ":" + r[0]).ToList();
        }

        /// <summary>Snapshot of every Global field a test may mutate, restored by <see cref="Restore"/>.</summary>
        public static Dictionary<string, object> Capture()
        {
            return typeof(Global).GetFields().Where(f => f.IsStatic && !f.IsInitOnly)
                .ToDictionary(f => f.Name, f => f.GetValue(null));
        }

        public static void Restore(Dictionary<string, object> state)
        {
            foreach (var f in typeof(Global).GetFields().Where(f => f.IsStatic && !f.IsInitOnly))
            {
                f.SetValue(null, state[f.Name]);
            }
        }
    }
}

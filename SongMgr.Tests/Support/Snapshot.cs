using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using Xunit;

namespace SongMgr.Tests.Support
{
    /// <summary>
    /// Golden-master assertions. Snapshots under SongMgr.Tests/Snapshots were recorded from the
    /// code as it was before the refactor; a refactor passes only if it reproduces them exactly.
    /// Set SONGMGR_RECORD_SNAPSHOTS=1 to (re)record — only ever do that on unrefactored code.
    /// </summary>
    internal static class Snapshot
    {
        public static void Match(string name, string actual, [CallerFilePath] string callerFile = "")
        {
            string dir = Path.Combine(FindTestsRoot(callerFile), "Snapshots");
            string file = Path.Combine(dir, name + ".txt");
            actual = actual.Replace("\r\n", "\n");

            if (Environment.GetEnvironmentVariable("SONGMGR_RECORD_SNAPSHOTS") == "1")
            {
                Directory.CreateDirectory(dir);
                File.WriteAllText(file, actual, new UTF8Encoding(false));
                return;
            }

            Assert.True(File.Exists(file), "Missing snapshot " + file + ". Record it on the pre-refactor code.");
            string expected = File.ReadAllText(file, Encoding.UTF8).Replace("\r\n", "\n");
            if (expected == actual) return;

            string[] e = expected.Split('\n');
            string[] a = actual.Split('\n');
            int n = Math.Min(e.Length, a.Length);
            for (int i = 0; i < n; i++)
            {
                if (e[i] != a[i])
                {
                    Assert.Fail($"Snapshot {name} differs at line {i + 1}:\n  expected: {e[i]}\n  actual:   {a[i]}");
                }
            }
            Assert.Fail($"Snapshot {name} differs in length: expected {e.Length} lines, actual {a.Length} lines.");
        }

        static string FindTestsRoot(string callerFile)
        {
            string dir = Path.GetDirectoryName(callerFile);
            while (dir != null && !File.Exists(Path.Combine(dir, "SongMgr.Tests.csproj")))
            {
                dir = Path.GetDirectoryName(dir);
            }
            return dir ?? throw new InvalidOperationException("Cannot locate SongMgr.Tests.csproj from " + callerFile);
        }
    }
}

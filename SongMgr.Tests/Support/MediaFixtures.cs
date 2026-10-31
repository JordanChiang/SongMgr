using System;
using System.Diagnostics;
using System.IO;

namespace SongMgr.Tests.Support
{
    /// <summary>
    /// Builds small deterministic media files with the bundled ffmpeg: PCM audio whose
    /// left/right channels and streams have different loudness, so track selection shows up in the gain.
    /// </summary>
    internal static class MediaFixtures
    {
        public static string FFmpeg => Path.Combine(AppContext.BaseDirectory, "Tools", "ffmpeg.exe");

        static string Tone(double left, double right, int freq) =>
            $"-f lavfi -i \"aevalsrc={left.ToString(System.Globalization.CultureInfo.InvariantCulture)}*sin({freq}*2*PI*t)|{right.ToString(System.Globalization.CultureInfo.InvariantCulture)}*sin({freq}*2*PI*t):s=8000:d=4\"";

        /// <summary>Creates fixtures with 1, 2 and 3 audio streams plus an empty (corrupt) file.</summary>
        public static (string one, string two, string three, string empty) Create(string dir)
        {
            string one = Path.Combine(dir, "one_stream.mkv");
            string two = Path.Combine(dir, "two_streams.mkv");
            string three = Path.Combine(dir, "three_streams.mkv");
            string empty = Path.Combine(dir, "empty.mpg");

            Run($"{Tone(0.05, 0.6, 440)} -c:a pcm_s16le \"{one}\"");
            Run($"{Tone(0.1, 0.1, 440)} {Tone(0.7, 0.7, 660)} -map 0:a -map 1:a -c:a pcm_s16le \"{two}\"");
            Run($"{Tone(0.1, 0.1, 440)} {Tone(0.4, 0.4, 550)} {Tone(0.8, 0.8, 660)} -map 0:a -map 1:a -map 2:a -c:a pcm_s16le \"{three}\"");
            File.WriteAllBytes(empty, new byte[0]);
            return (one, two, three, empty);
        }

        static void Run(string args)
        {
            using (Process p = new Process())
            {
                p.StartInfo.FileName = FFmpeg;
                p.StartInfo.Arguments = "-y -v error " + args;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardError = true;
                p.Start();
                string err = p.StandardError.ReadToEnd();
                p.WaitForExit();
                if (p.ExitCode != 0) throw new InvalidOperationException("ffmpeg failed: " + err);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Media
{
    /// <summary>
    /// Golden master for ffprobe/ffmpeg use: audio-track detection, duration,
    /// replay-gain measurement per track, and the 歌曲音量 formula with each caller's clamp.
    /// </summary>
    public class ProbeSnapshotTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string dir = TestEnv.NewTempDir("media");

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(dir);
        }

        [Fact]
        public void Probe_matches_pre_refactor_snapshot()
        {
            var f = MediaFixtures.Create(dir);
            var files = new[] { ("one", f.one), ("two", f.two), ("three", f.three), ("empty", f.empty) };
            StringBuilder sb = new StringBuilder();

            foreach (string trackMode in new[] { "False", "True" })
            {
                Global.SongMgrSongTrackMode = trackMode;
                foreach (var (name, file) in files)
                {
                    sb.Append($"tracks {trackMode} {name} => {Probe.DetectTrack(file)}\n");
                }
            }

            foreach (var (name, file) in files)
            {
                sb.Append($"duration {name} => {Probe.Duration(file).ToString("F3", CultureInfo.InvariantCulture)}\n");
            }

            foreach (string trackMode in new[] { "False", "True" })
            {
                Global.SongMgrSongTrackMode = trackMode;
                foreach (var (name, file) in files)
                foreach (int track in new[] { 1, 2, 3 })
                foreach (string seek in new[] { "", "-ss 1 -t 2" })
                {
                    var v = Probe.Volume(file, track, seek);
                    sb.Append($"gain {trackMode} {name} t{track} [{seek}] => {v.success} {v.gain.ToString(CultureInfo.InvariantCulture)} {v.error}\n");
                }
            }

            Snapshot.Match("media-probe", sb.ToString());
        }

        [Fact]
        public void Volume_formula_matches_pre_refactor_snapshot()
        {
            StringBuilder sb = new StringBuilder();
            foreach (int basevolume in new[] { 30, 50, 75, 100 })
            foreach (double gain in new[] { -20.0, -9.37, -3.0, -0.01, 0.0, 0.5, 2.25, 6.0, 12.0, 25.5 })
            {
                sb.Append($"{basevolume} {gain.ToString(CultureInfo.InvariantCulture)} => {Probe.Formula(basevolume, gain)}\n");
            }
            Snapshot.Match("media-volume-formula", sb.ToString());
        }

        /// <summary>Thin adapter over today's entry points, so the same snapshot checks the refactored module.</summary>
        static class Probe
        {
            public static string DetectTrack(string file) => MediaProbe.DetectTrack(file);

            public static double Duration(string file) => MediaProbe.Duration(file);

            public static (bool success, double gain, string error) Volume(string file, int track, string seek)
            {
                MediaProbe.SongVolumeValue v = MediaProbe.MeasureGain(file, track, seek);
                return (v.Success, v.GainDB, v.ErrorMessage);
            }

            public static string Formula(int basevolume, double gain)
            {
                string raw = MediaProbe.SongVolume(basevolume, gain);
                string add = MediaProbe.SongVolume(basevolume, gain, VolumeRange.Library);
                string norm = MediaProbe.SongVolume(basevolume, gain, VolumeRange.Normalize);
                return raw + " add=" + add + " norm=" + norm;
            }
        }
    }
}

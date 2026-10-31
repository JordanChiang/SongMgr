using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>The 歌曲音量 range a feature accepts; values outside are pulled to the nearest end.</summary>
    struct VolumeRange
    {
        public int Min { get; }
        public int Max { get; }

        public VolumeRange(int min, int max) { Min = min; Max = max; }

        /// <summary>加歌 音量偵測 and 歌庫維護 音量平衡.</summary>
        public static readonly VolumeRange Library = new VolumeRange(10, 100);

        /// <summary>音量正規化 suggestions.</summary>
        public static readonly VolumeRange Normalize = new VolumeRange(5, 100);
    }

    /// <summary>
    /// Media probe: everything the program asks ffprobe / ffmpeg about a song file — how many audio
    /// streams it has, which 聲道 is the 伴唱, how long it is, its 重播增益 per track, and whether it decodes
    /// cleanly — plus the 歌曲音量 that follows from a 重播增益. Process start-up and output parsing live here only.
    /// </summary>
    static class MediaProbe
    {
        private static string FFmpegPath => Application.StartupPath + @"\Tools\ffmpeg.exe";
        private static string FFprobePath => Application.StartupPath + @"\Tools\ffprobe.exe";

        public class SongVolumeValue
        {
            public double GainDB { get; set; }
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
        }

        #region --- Running the tools ---

        /// <summary>Runs a tool to completion and returns the lines it wrote to stdout (or stderr), or null when it is missing.</summary>
        private static List<string> Run(string exe, string arguments, bool readStandardError)
        {
            if (!File.Exists(exe)) return null;

            using (Process p = new Process())
            {
                p.StartInfo.FileName = exe;
                p.StartInfo.Arguments = arguments;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardOutput = !readStandardError;
                p.StartInfo.RedirectStandardError = readStandardError;
                if (readStandardError) p.StartInfo.StandardErrorEncoding = Encoding.UTF8; else p.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                p.Start();

                StreamReader reader = readStandardError ? p.StandardError : p.StandardOutput;
                List<string> lines = new List<string>();
                string line;
                while ((line = reader.ReadLine()) != null) lines.Add(line);
                p.WaitForExit();
                return lines;
            }
        }

        #endregion

        #region --- Streams, 聲道, duration ---

        /// <summary>Number of audio streams in the file (0 when ffprobe is missing or cannot read it).</summary>
        public static int AudioStreamCount(string file)
        {
            List<string> lines = Run(FFprobePath, (char)34 + file + (char)34 + " -show_entries stream=index,codec_type -of compact=p=0 -v 0", false);
            if (lines == null) return 0;

            int AudioCount = 0;
            Regex r = new Regex(@"\w+=audio");
            foreach (string line in lines)
            {
                if (r.IsMatch(line)) AudioCount += 1;
            }
            return AudioCount;
        }

        /// <summary>
        /// The 歌曲聲道 to store for the file: 1 stream → 1, 2 streams → 1 (or 2 with 雙音軌 track mode),
        /// 3 streams → 3, otherwise 0.
        /// </summary>
        public static string DetectTrack(string file)
        {
            switch (AudioStreamCount(file))
            {
                case 1:
                    return "1";
                case 2:
                    return (Global.SongMgrSongTrackMode == "True") ? "2" : "1";
                case 3:
                    return "3";
                default:
                    return "0";
            }
        }

        /// <summary>Returns file duration in seconds via ffprobe, or 0 on failure.</summary>
        public static double Duration(string file)
        {
            double duration = 0;
            string args = string.Format("-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 \"{0}\"", file);
            List<string> lines = Run(FFprobePath, args, false);
            if (lines != null)
            {
                double.TryParse(string.Join("\n", lines).Trim(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out duration);
            }
            return duration;
        }

        /// <summary>True when ffmpeg decodes the whole file without reporting an error.</summary>
        public static bool VerifyFile(string file)
        {
            List<string> lines = Run(FFmpegPath, string.Format("-y -i \"{0}\" -v error -f null /dev/null", file), true);
            if (lines == null) return true;
            foreach (string line in lines)
            {
                if (!string.IsNullOrEmpty(line)) return false;
            }
            return true;
        }

        #endregion

        #region --- 重播增益 and 歌曲音量 ---

        /// <summary>
        /// 重播增益 of the 伴唱 for <paramref name="track"/>, picking the stream / channel the player would play:
        ///   1 stream  – Track 1 left ch, Track 2 right ch (SongMgrSongTrackMode has no effect)
        ///   2 streams – Track 1: stream[1] normal / stream[0] TrackMode=True
        ///               Track 2: stream[0] normal / stream[1] TrackMode=True
        ///   ≥3 streams – stream index = SongTrack value (player uses AudioTrack = SongTrack)
        /// <paramref name="seekArgs"/>: optional input-seek args, e.g. "-ss 30 -t 60" for fast mode.
        /// </summary>
        public static SongVolumeValue MeasureGain(string file, int track = 1, string seekArgs = "")
        {
            SongVolumeValue result = new SongVolumeValue { GainDB = 0, Success = false, ErrorMessage = "" };

            if (!File.Exists(FFmpegPath))
            {
                result.ErrorMessage = "找不到 FFmpeg 執行檔";
                return result;
            }

            try
            {
                bool trackMode = (Global.SongMgrSongTrackMode == "True");

                string afFilter;
                string mapArg;

                if (track <= 2)
                {
                    if (track == 1)
                    {
                        if (trackMode)
                        {
                            // Accompaniment on stream[0] / left channel
                            mapArg   = "";
                            afFilter = "pan=mono|c0=c0,replaygain";
                        }
                        else
                        {
                            // Accompaniment on stream[1] / right channel
                            mapArg   = "-map 0:a:1";
                            afFilter = "replaygain";
                        }
                    }
                    else // track == 2
                    {
                        if (trackMode)
                        {
                            mapArg   = "-map 0:a:1";
                            afFilter = "replaygain";
                        }
                        else
                        {
                            mapArg   = "-map 0:a:0";
                            afFilter = "replaygain";
                        }
                    }
                }
                else
                {
                    // 3+ streams: player uses AudioTrack = SongTrack (1-based track number).
                    mapArg   = string.Format("-map 0:a:{0}", track - 1);
                    afFilter = "replaygain";
                }

                // seekArgs go BEFORE -i for fast input seeking (e.g. "-ss 30 -t 60")
                string args = string.Format(
                    "{0} -i \"{1}\" {2} -af \"{3}\" -vn -sn -dn -f null /dev/null",
                    seekArgs, file, mapArg, afFilter).TrimStart();

                List<string> lines = Run(FFmpegPath, args, true);
                if (lines == null)
                {
                    result.ErrorMessage = "FFmpeg 執行失敗";
                    return result;
                }

                double GainDB = 0;
                bool matched = false;
                Regex gainline = new Regex(@"\[Parsed_replaygain_\d+.+?\] track_gain =");
                foreach (string line in lines)
                {
                    if (gainline.IsMatch(line))
                    {
                        GainDB = Convert.ToDouble(Regex.Replace(line, @"\[.+?\]|track_gain =|dB|/s", ""));
                        matched = true;
                    }
                }

                if (matched)
                {
                    result.GainDB = Math.Round(GainDB, 2);
                    result.Success = true;
                }
                else
                {
                    result.ErrorMessage = "無法讀取重播增益 (FFmpeg 分析失敗或無音軌)";
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }

        /// <summary>The 歌曲音量 that brings a song with this 重播增益 to <paramref name="baseVolume"/> (not clamped).</summary>
        public static string SongVolume(int baseVolume, double gainDB)
        {
            return Convert.ToInt32(baseVolume * Math.Pow(10, gainDB / 20)).ToString();
        }

        /// <summary><see cref="SongVolume(int, double)"/> pulled into <paramref name="range"/>.</summary>
        public static string SongVolume(int baseVolume, double gainDB, VolumeRange range)
        {
            return Clamp(Convert.ToInt32(SongVolume(baseVolume, gainDB)), range).ToString();
        }

        public static int Clamp(int volume, VolumeRange range)
        {
            return Math.Max(range.Min, Math.Min(range.Max, volume));
        }

        #endregion
    }
}

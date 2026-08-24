using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    class FFmpeg
    {
        private static string FFmpegPath  = Application.StartupPath + @"\Tools\ffmpeg.exe";
        private static string FFprobePath = Application.StartupPath + @"\Tools\ffprobe.exe";

        private static StreamReader RunFFmpeg(string fileName, string arguments)
        {
            if (!File.Exists(FFmpegPath)) return null;

            StreamReader sr;
            using (Process p = new Process())
            {
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.FileName = fileName;
                p.StartInfo.Arguments = arguments;
                p.Start();
                sr = p.StandardError;
            }
            return sr;
        }

        // Reads stdout (used for ffprobe queries)
        private static StreamReader RunFFprobe(string arguments)
        {
            if (!File.Exists(FFprobePath)) return null;

            StreamReader sr;
            using (Process p = new Process())
            {
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.FileName = FFprobePath;
                p.StartInfo.Arguments = arguments;
                p.Start();
                sr = p.StandardOutput;
            }
            return sr;
        }

        /// <summary>Returns file duration in seconds via ffprobe, or 0 on failure.</summary>
        public static double GetFileDuration(string file)
        {
            double duration = 0;
            string args = string.Format(
                "-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 \"{0}\"",
                file);
            using (StreamReader sr = RunFFprobe(args))
            {
                if (sr != null)
                    double.TryParse(sr.ReadToEnd().Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out duration);
            }
            return duration;
        }

        public class SongVolumeValue
        {
            public double GainDB { get; set; }
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
        }

        public static SongVolumeValue GetSongVolume(string file)
        {
            return GetSongVolume(file, 1);
        }

        /// <summary>
        /// Analyse accompaniment-track volume using Song_Track from DB.
        ///   Track 1  – single-stream, left channel only  (左伴右唱)
        ///   Track 2  – single/multi-stream, stream 0, both channels (1伴2唱)
        ///   Track 3+ – multi-stream, stream (track-1), both channels
        /// </summary>
        /// <summary>
        /// Analyse accompaniment-track volume using Song_Track from DB.
        /// Mirrors mediaUriElement_MediaOpened() channel/stream selection:
        ///   1 stream  – Track 1 left ch, Track 2 right ch (SongMgrSongTrackMode has no effect)
        ///   2 streams – Track 1: stream[1] normal / stream[0] TrackMode=True
        ///               Track 2: stream[0] normal / stream[1] TrackMode=True
        ///   ≥3 streams – stream index = SongTrack value (player uses AudioTrack = SongTrack)
        /// seekArgs: optional input-seek args, e.g. "-ss 30 -t 60" for fast mode.
        /// </summary>
        public static SongVolumeValue GetSongVolume(string file, int track, string seekArgs = "")
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

                // Build audio filter chain mirroring mediaUriElement_MediaOpened:
                string afFilter;
                string mapArg;

                if (track <= 2)
                {
                    if (track == 1)
                    {
                        if (trackMode)
                        {
                            // Accompaniment on stream[0] / left channel
                            mapArg   = "";                                      // use default (first) stream
                            afFilter = "pan=mono|c0=c0,replaygain";             // left channel for 1-stream
                        }
                        else
                        {
                            // Accompaniment on stream[1] / right channel
                            mapArg   = "-map 0:a:1";                            // try stream[1] for 2-stream
                            afFilter = "replaygain";
                        }
                    }
                    else // track == 2
                    {
                        if (trackMode)
                        {
                            // Accompaniment on stream[1] / right channel
                            mapArg   = "-map 0:a:1";
                            afFilter = "replaygain";
                        }
                        else
                        {
                            // Accompaniment on stream[0] / left channel
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

                using (StreamReader sr = RunFFmpeg(FFmpegPath, args))
                {
                    if (sr == null)
                    {
                        result.ErrorMessage = "FFmpeg 執行失敗";
                        return result;
                    }

                    double GainDB = 0;
                    bool matched = false;
                    Regex gainline = new Regex(@"\[Parsed_replaygain_\d+.+?\] track_gain =");

                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        if (line != null && gainline.IsMatch(line))
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
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }

        public static string CalSongVolume(int basevolume, double GainDB)
        {
            string SongVolume = Convert.ToInt32(basevolume * Math.Pow(10, GainDB / 20)).ToString();
            return SongVolume;
        }


        public static bool VerifyFile(string file)
        {
            bool result = true;

            using (StreamReader sr = RunFFmpeg(FFmpegPath, string.Format("-y -i \"{0}\" -v error -f null /dev/null", file)))
            {
                if (sr != null)
                {
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        if (!string.IsNullOrEmpty(line)) result = false;
                    }
                }
            }
            return result;
        }
    }
}

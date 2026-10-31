using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Analysis
{
    /// <summary>
    /// Golden master for 歌曲分析: a corpus of real-world file layouts analysed under several
    /// settings profiles against the bundled singer database.
    /// </summary>
    public class AnalysisSnapshotTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string root = TestEnv.NewTempDir("analysis");

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(root);
        }

        internal static readonly string[] Corpus =
        {
            @"周杰倫_晴天.mpg",
            @"周杰倫-晴天-國語-VR.mpg",
            @"晴天_周杰倫.mkv",
            @"01.周杰倫_稻香.mpg",
            @"000123_周杰倫_簡單愛.mpg",
            @"000124_簡單愛_周杰倫.mpg",
            @"周杰倫_晴天(國語_VL).mpg",
            @"周杰倫_晴天【國語】.mpg",
            @"周杰倫_晴天(演唱會).mpg",
            @"周杰倫_晴天[合唱版].mpg",
            @"周杰倫&蔡依林_布拉格廣場.mpg",
            @"周杰倫+蔡依林-布拉格廣場-合唱.mpg",
            @"S.H.E_Super Star.mpg",
            @"梁靜茹_兒歌.mpg",
            @"梁靜茹_兒歌_兒歌.mpg",
            @"莫文蔚_台.mpg",
            @"莫文蔚-台-台語.mpg",
            @"台灣女孩_羅百吉&寶貝.mpg",
            @"羅百吉&寶貝_台灣女孩.mpg",
            @"Maroon 5_sugar_英語.mp4",
            @"maroon 5 - this love.mp4",
            @"不存在的歌手_某首歌.mpg",
            @"單一名稱.mpg",
            @"國語\男歌星\周杰倫\晴天.mpg",
            @"國語\女歌星\蔡依林\日不落.mpg",
            @"台語\男\江蕙\家後.mpg",
            @"粵語\合唱\張學友&陳慧嫻\相思風雨中.mpg",
            @"英語\外國團體\Maroon 5\Sugar.mp4",
            @"日語\不明\歌手甲_歌曲乙_日語.mpg",
            @"新增資料夾\周杰倫\周杰倫_晴天.mpg",
            @"新增資料夾\國\某歌手_某歌_VR.mpg",
            @"周杰倫_晴天_V0.mpg",
            @"周杰倫_晴天_r.mpg",
            @"周杰倫_晴天(l).mpg",
            @"周杰倫_晴天_男.mpg",
            @"周杰倫_晴天_外男_台.mpg",
            @"张学友_吻别.mpg",
            @"张学友_吻别_国语.mpg",
            @"周杰倫_晴天 (Live).mpg",
            @"周杰倫 _ 晴天 _ 國語.mpg",
            @"F.I.R._我們的愛.mpg",
            @"F.I.R. - 我們的愛 - 國語.mpg",
            @"Super Star-S.H.E.mpg",
            @"周杰倫_THE SHOW MUST GO ON.mpg",
            @"the beatles_hey jude_英.mp4",
        };

        internal sealed class Profile
        {
            public string Name;
            public Action Apply;
        }

        internal static IEnumerable<Profile> Profiles()
        {
            foreach (string mode in new[] { "1", "2", "3", "4", "5", "6" })
            {
                yield return new Profile
                {
                    Name = "ident" + mode,
                    Apply = () => { Global.SongAddSongIdentificationMode = mode; }
                };
            }
            yield return new Profile
            {
                Name = "ident1-tc-eng-types",
                Apply = () =>
                {
                    Global.SongAddSongIdentificationMode = "1";
                    Global.SongAddEnableConvToTC = "True";
                    Global.SongAddEngSongNameFormat = "True";
                    Global.SongMgrSongType = "演唱會,MTV,Live";
                    Global.SongAddDefaultSongType = "2";
                    Global.SongAddSpecialStr = "S.H.E|F.I.R.";
                }
            };
            yield return new Profile
            {
                Name = "ident1-monitor-defaults",
                Apply = () =>
                {
                    Global.SongAddSongIdentificationMode = "1";
                    Global.SongMgrSongAddMode = "4";
                    Global.SongAddDefaultSongLang = "11";
                    Global.SongAddDefaultSingerType = "3";
                    Global.SongAddDefaultSongTrack = "6";
                    Global.SongAddEnableVolumeDetect = "True";
                    Global.SongMaintenanceReplayGainVolume = "50";
                }
            };
        }

        internal static void ResetAnalysisSettings()
        {
            Global.SongAddSongIdentificationMode = "1";
            Global.SongAddEnableConvToTC = "False";
            Global.SongAddEngSongNameFormat = "False";
            Global.SongMgrSongType = "";
            Global.SongAddDefaultSongType = "1";
            Global.SongAddSpecialStr = "";
            Global.SongMgrSongAddMode = "3";
            Global.SongAddDefaultSongLang = "1";
            Global.SongAddDefaultSingerType = "8";
            Global.SongAddDefaultSongTrack = "1";
            Global.SongAddDefaultSongVolume = "100";
            Global.SongAddEnableVolumeDetect = "False";
            Global.SongMgrSongTrackMode = "False";
        }

        internal static List<string> CreateCorpusFiles(string root)
        {
            List<string> files = new List<string>();
            foreach (string rel in Corpus)
            {
                string path = Path.Combine(root, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllBytes(path, new byte[0]);
                files.Add(path);
            }
            return files;
        }

        internal static string RowText(DataRow row, string root)
        {
            string[] cols =
            {
                "Song_AddStatus", "Song_Id", "Song_Lang", "Song_SingerType", "Song_Singer", "Song_SongName",
                "Song_Track", "Song_SongType", "Song_Volume", "Song_WordCount", "Song_PlayCount", "Song_MB",
                "Song_Spell", "Song_SpellNum", "Song_SongStroke", "Song_PenStyle", "Song_PlayState",
                "Song_SrcPath", "Song_SortIndex", "Song_ReplayGain"
            };
            return string.Join("|", cols.Select(c => row[c].ToString())).Replace(root, "<ROOT>");
        }

        [Fact]
        public void Analysis_matches_pre_refactor_snapshot()
        {
            TestEnv.UseScratchDatabases(root);
            List<string> files = CreateCorpusFiles(Path.Combine(root, "in"));

            StringBuilder sb = new StringBuilder();
            foreach (Profile profile in Profiles())
            {
                ResetAnalysisSettings();
                profile.Apply();
                TestEnv.ResetLogs();
                Global.TotalList = new List<int>() { 0, 0, 0, 0, 0 };

                sb.Append("## ").Append(profile.Name).Append('\n');
                foreach (string line in AnalyseAll(files, root))
                {
                    sb.Append(line).Append('\n');
                }
            }
            Snapshot.Match("analysis", sb.ToString());
        }

        /// <summary>Runs one analysis pass the way 加歌 does and returns the rows, log and counters as text.</summary>
        internal static List<string> AnalyseAll(List<string> files, string root)
        {
            List<string> lines = new List<string>();
            SongAnalysis.CreateSongDataTable();
            foreach (string file in files)
            {
                SongAnalysis.SongInfoAnalysis(file);
            }
            foreach (DataRow row in SongAnalysis.SongAnalysisDT.Rows)
            {
                lines.Add(RowText(row, root));
            }
            foreach (string log in TestEnv.SongLog())
            {
                lines.Add("LOG " + log.Replace(root, "<ROOT>"));
            }
            lines.Add("DUP " + Global.TotalList[0]);
            SongAnalysis.DisposeSongDataTable();
            return lines;
        }

        /// <summary>Same flattening SongAdd_SongAnalysisTask does for the keyword lists.</summary>
        internal static List<string> KeyWords(List<string> source)
        {
            List<string> list = new List<string>();
            foreach (string str in source)
            {
                list.AddRange(str.Split(','));
            }
            return list;
        }
    }
}

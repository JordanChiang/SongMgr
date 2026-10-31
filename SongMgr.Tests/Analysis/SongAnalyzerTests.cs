using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Analysis
{
    /// <summary>
    /// 歌曲分析 through its own interface: a file, the run's settings, and a singer lookup.
    /// The lookup here is built in memory — the second adapter next to <see cref="SingerLookup.LoadFromDatabase"/>.
    /// </summary>
    public class SongAnalyzerTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string root = TestEnv.NewTempDir("analyzer");

        public SongAnalyzerTests()
        {
            // 注音 (拼音 / 筆劃) still comes from CrazySongMgr.mdb.
            TestEnv.UseScratchDatabases(root);
        }

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(root);
        }

        static SingerLookup Singers() => new SingerLookup(
            new List<string> { "周杰倫", "蔡依林", "S.H.E", "張學友", "陳慧嫻", "Maroon 5" },
            new List<string> { "0", "1", "2", "0", "1", "6" });

        static SongAnalysisSettings Settings(Action<SongAnalysisSettings> change = null)
        {
            SongAnalysisSettings s = SongAnalysisSettings.FromGlobal();
            s.IdentificationMode = "1";
            s.ConvertToTraditional = false;
            s.SpecialStr = "";
            s.SongTypes = "";
            s.DefaultSongLang = "1";
            s.DefaultSingerType = "8";
            s.DefaultSongTrack = "1";
            s.DefaultSongType = "1";
            s.DefaultSongVolume = "100";
            s.DetectVolume = false;
            s.TitleCaseEnglish = false;
            change?.Invoke(s);
            return s;
        }

        string File(string relative)
        {
            string path = Path.Combine(root, "in", relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            System.IO.File.WriteAllBytes(path, new byte[0]);
            return path;
        }

        [Theory]
        [InlineData(@"周杰倫_晴天.mpg", "國語", "0", "周杰倫", "晴天", "1", "")]
        [InlineData(@"晴天_周杰倫.mpg", "國語", "0", "周杰倫", "晴天", "1", "")]
        [InlineData(@"周杰倫-晴天-台語-VR.mpg", "台語", "0", "周杰倫", "晴天", "1", "")]
        [InlineData(@"周杰倫_晴天(國語_VL).mpg", "國語", "0", "周杰倫", "晴天(國語_VL)", "2", "")]
        [InlineData(@"周杰倫&蔡依林_布拉格廣場.mpg", "國語", "3", "周杰倫&蔡依林", "布拉格廣場", "1", "")]
        [InlineData(@"周杰倫+蔡依林_布拉格廣場.mpg", "國語", "3", "周杰倫&蔡依林", "布拉格廣場", "1", "")]
        [InlineData(@"粵語\合唱\張學友&陳慧嫻\相思風雨中.mpg", "粵語", "3", "張學友&陳慧嫻", "相思風雨中", "1", "")]
        [InlineData(@"英語\Maroon 5\Sugar.mp4", "英語", "6", "Maroon 5", "Sugar", "1", "")]
        [InlineData(@"沒聽過_某首歌.mpg", "國語", "7", "沒聽過", "某首歌", "1", "")]
        public void Reads_lang_singer_name_and_track_from_the_file_layout(string file, string lang, string singerType, string singer, string name, string track, string songType)
        {
            SongAnalysisResult r = SongAnalyzer.Analyze(File(file), Settings(), Singers());

            Assert.Equal(new[] { lang, singerType, singer, name, track, songType }, new[] { r.Lang, r.SingerType, r.Singer, r.SongName, r.Track, r.SongType });
            Assert.Equal(Path.Combine(root, "in", file), r.SrcPath);
        }

        [Fact]
        public void Simplified_chinese_is_converted_before_matching_singers()
        {
            // 2.10.11: on .NET 8 the 簡轉繁 call threw / returned garbage.
            SongAnalysisResult r = SongAnalyzer.Analyze(File(@"张学友_吻别.mpg"), Settings(s => s.ConvertToTraditional = true), Singers());

            Assert.Equal("張學友", r.Singer);
            Assert.Equal("吻別", r.SongName);
            Assert.Equal("0", r.SingerType);
        }

        [Fact]
        public void Unknown_language_marks_the_song_as_not_addable()
        {
            SongAnalysisResult r = SongAnalyzer.Analyze(File(@"周杰倫_晴天.mpg"), Settings(s => s.DefaultSongLang = "11"), Singers());

            Assert.Equal("未知", r.Lang);
            Assert.Equal("語系類別必須有值才能加歌!", r.AddStatus);
            Assert.Equal("1", r.SortIndex);
        }

        [Fact]
        public void Parallel_analysis_gives_the_same_results_as_one_at_a_time()
        {
            List<string> files = AnalysisSnapshotTests.CreateCorpusFiles(Path.Combine(root, "corpus"));
            SongAnalysisSettings settings = Settings(s => { s.SongTypes = "演唱會,MTV"; s.SpecialStr = "S.H.E|F.I.R."; });
            SingerLookup singers = SingerLookup.LoadFromDatabase();

            Dictionary<string, string> sequential = files.ToDictionary(f => f, f => Text(SongAnalyzer.Analyze(f, settings, singers)));

            for (int round = 0; round < 3; round++)
            {
                ConcurrentDictionary<string, string> parallel = new ConcurrentDictionary<string, string>();
                Parallel.ForEach(files.Concat(files).Concat(files), f => parallel[f] = Text(SongAnalyzer.Analyze(f, settings, singers)));
                foreach (string f in files) Assert.Equal(sequential[f], parallel[f]);
            }
        }

        static string Text(SongAnalysisResult r) => string.Join("|", r.AddStatus, r.Id, r.Lang, r.SingerType, r.Singer, r.SongName, r.Track, r.SongType, r.Volume,
            r.WordCount, r.PlayCount, r.MB, r.Spell, r.SpellNum, r.SongStroke, r.PenStyle, r.PlayState, r.SrcPath, r.SortIndex, r.ReplayGain);
    }
}

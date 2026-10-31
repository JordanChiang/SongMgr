using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Records
{
    /// <summary>
    /// The song grids used to hand songs around as "|"-joined strings read by position.
    /// These tests rebuild those strings exactly as the old code did (copied from master @ 21fb779)
    /// and require <see cref="SongRecord"/> to carry the same values, field by field.
    /// </summary>
    public class SongRecordEquivalenceTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string dir = TestEnv.NewTempDir("records");

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(dir);
        }

        // ---------- fixtures: real ktv_Song schema, as the 歌庫查詢 grid sees it ----------

        DataTable QueryTable()
        {
            TestEnv.UseScratchDatabases(dir);
            using (OleDbConnection conn = CommonFunc.OleDbOpenConn(Global.CrazyktvDatabaseFile, ""))
            {
                string sql = "insert into ktv_Song (Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState, Song_ReplayGain) " +
                             "values (@a, @b, @c, @d, @e, @f, @g, @h, @i, @j, @k, @l, @m, @n, @o, @p, @q, @r, @s, @t)";
                object[][] rows =
                {
                    new object[] { "000123", "國語", 0, "周杰倫", "晴天", 1, "", 87, 2, 15, 12.34f, new DateTime(2024, 1, 2, 3, 4, 5), "周杰倫_晴天_VR.mpg", @"D:\KTV\國語\男\周杰倫\", "ㄑㄊ", "42", 12, "丨一", 0, -3.2 },
                    new object[] { "100045", "台語", 3, "江蕙&陳雷", "愛人醉落去", 5, "演唱會", 100, 5, 0, 0.5f, new DateTime(2019, 12, 31, 23, 59, 59), "x.mkv", @"E:\Import\", "ㄞㄖㄗㄌㄑ", "85621", 13, "丿丿一丿丿", 1, DBNull.Value },
                };
                foreach (object[] r in rows)
                {
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        foreach (object v in r) cmd.Parameters.AddWithValue("@p", v);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            return CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, "select * from ktv_Song order by Song_Id", "");
        }

        /// <summary>The 加歌 grid is bound to the 歌曲分析 table.</summary>
        static DataTable AddTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Song_AddStatus", typeof(string));
            dt.Columns.Add("Song_Id", typeof(string));
            dt.Columns.Add("Song_Lang", typeof(string));
            dt.Columns.Add("Song_SingerType", typeof(int));
            dt.Columns.Add("Song_Singer", typeof(string));
            dt.Columns.Add("Song_SongName", typeof(string));
            dt.Columns.Add("Song_Track", typeof(int));
            dt.Columns.Add("Song_SongType", typeof(string));
            dt.Columns.Add("Song_Volume", typeof(int));
            dt.Columns.Add("Song_WordCount", typeof(int));
            dt.Columns.Add("Song_PlayCount", typeof(int));
            dt.Columns.Add("Song_MB", typeof(float));
            dt.Columns.Add("Song_CreatDate", typeof(DateTime));
            dt.Columns.Add("Song_FileName", typeof(string));
            dt.Columns.Add("Song_Path", typeof(string));
            dt.Columns.Add("Song_Spell", typeof(string));
            dt.Columns.Add("Song_SpellNum", typeof(string));
            dt.Columns.Add("Song_SongStroke", typeof(int));
            dt.Columns.Add("Song_PenStyle", typeof(string));
            dt.Columns.Add("Song_PlayState", typeof(int));
            dt.Columns.Add("Song_SrcPath", typeof(string));
            dt.Columns.Add("Song_SortIndex", typeof(string));
            dt.Columns.Add("Song_ReplayGain", typeof(string));
            dt.Rows.Add("此歌手尚未設定歌手資料!", "", "國語", 10, "新人", "新歌", 2, "MTV", 80, 2, 0, 3.75f, new DateTime(2026, 10, 31, 8, 0, 0), "", "", "ㄒㄍ", "44", 13, "丶一", 0, @"F:\in\新人_新歌.mpg", "2", "1.5");
            dt.Rows.Add("", "", "英語", 6, "Maroon 5", "Sugar", 1, "", 100, 1, 0, 0f, new DateTime(2026, 1, 1), "", "", "S", "7", 1, "", 0, @"F:\in\Maroon 5_Sugar.mp4", "4", "");
            return dt;
        }

        static DataGridView Grid(DataTable dt)
        {
            DataGridView grid = new DataGridView { BindingContext = new BindingContext(), AutoGenerateColumns = true };
            grid.DataSource = dt;
            Assert.Equal(dt.Rows.Count, grid.Rows.Count - (grid.AllowUserToAddRows ? 1 : 0));
            return grid;
        }

        static IEnumerable<DataGridViewRow> Rows(DataGridView grid) => grid.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow);

        // ---------- legacy layouts, copied verbatim ----------

        /// <summary>SongQueryDataGridView.cs SelectionChanged (multi and single select).</summary>
        static string LegacyQuerySelect(DataGridViewRow row)
        {
            string SongId = row.Cells["Song_Id"].Value.ToString();
            string SongLang = row.Cells["Song_Lang"].Value.ToString();
            int SongSingerType = Convert.ToInt32(row.Cells["Song_SingerType"].Value);
            string SongSinger = row.Cells["Song_Singer"].Value.ToString();
            string SongSongName = row.Cells["Song_SongName"].Value.ToString();
            int SongTrack = Convert.ToInt32(row.Cells["Song_Track"].Value);
            string SongSongType = row.Cells["Song_SongType"].Value.ToString();
            string SongVolume = row.Cells["Song_Volume"].Value.ToString();
            string SongWordCount = row.Cells["Song_WordCount"].Value.ToString();
            string SongPlayCount = row.Cells["Song_PlayCount"].Value.ToString();
            string SongMB = row.Cells["Song_MB"].Value.ToString();
            string SongCreatDate = row.Cells["Song_CreatDate"].Value.ToString();
            string SongFileName = row.Cells["Song_FileName"].Value.ToString();
            string SongPath = row.Cells["Song_Path"].Value.ToString();
            string SongSpell = row.Cells["Song_Spell"].Value.ToString();
            string SongSpellNum = row.Cells["Song_SpellNum"].Value.ToString();
            string SongSongStroke = row.Cells["Song_SongStroke"].Value.ToString();
            string SongPenStyle = row.Cells["Song_PenStyle"].Value.ToString();
            string SongPlayState = row.Cells["Song_PlayState"].Value.ToString();
            string SongReplayGain = row.Cells["Song_ReplayGain"].Value.ToString();
            string SongSrcPath = Path.Combine(SongPath, SongFileName);
            return SongId + "|" + SongLang + "|" + SongSingerType + "|" + SongSinger + "|" + SongSongName + "|" + SongTrack + "|" + SongSongType + "|" + SongVolume + "|" + SongWordCount + "|" + SongPlayCount + "|" + SongMB + "|" + SongCreatDate + "|" + SongFileName + "|" + SongPath + "|" + SongSpell + "|" + SongSpellNum + "|" + SongSongStroke + "|" + SongPenStyle + "|" + SongPlayState + "|" + SongSrcPath + "|" + SongReplayGain;
        }

        /// <summary>SongAddDataGridView.cs SelectionChanged, single select (the one that carried ReplayGain).</summary>
        static string LegacyAddSelect(DataGridViewRow row)
        {
            string SongId = row.Cells["Song_Id"].Value.ToString();
            string SongLang = row.Cells["Song_Lang"].Value.ToString();
            int SongSingerType = Convert.ToInt32(row.Cells["Song_SingerType"].Value);
            string SongSinger = row.Cells["Song_Singer"].Value.ToString();
            string SongSongName = row.Cells["Song_SongName"].Value.ToString();
            int SongTrack = Convert.ToInt32(row.Cells["Song_Track"].Value);
            string SongSongType = row.Cells["Song_SongType"].Value.ToString();
            string SongVolume = row.Cells["Song_Volume"].Value.ToString();
            string SongWordCount = row.Cells["Song_WordCount"].Value.ToString();
            string SongPlayCount = row.Cells["Song_PlayCount"].Value.ToString();
            string SongMB = row.Cells["Song_MB"].Value.ToString();
            string SongCreatDate = row.Cells["Song_CreatDate"].Value.ToString();
            string SongFileName = row.Cells["Song_FileName"].Value.ToString();
            string SongPath = row.Cells["Song_Path"].Value.ToString();
            string SongSpell = row.Cells["Song_Spell"].Value.ToString();
            string SongSpellNum = row.Cells["Song_SpellNum"].Value.ToString();
            string SongSongStroke = row.Cells["Song_SongStroke"].Value.ToString();
            string SongPenStyle = row.Cells["Song_PenStyle"].Value.ToString();
            string SongPlayState = row.Cells["Song_PlayState"].Value.ToString();
            string SongReplayGain = row.Cells["Song_ReplayGain"].Value.ToString();
            string SongSrcPath = row.Cells["Song_SrcPath"].Value.ToString();
            return SongId + "|" + SongLang + "|" + SongSingerType + "|" + SongSinger + "|" + SongSongName + "|" + SongTrack + "|" + SongSongType + "|" + SongVolume + "|" + SongWordCount + "|" + SongPlayCount + "|" + SongMB + "|" + SongCreatDate + "|" + SongFileName + "|" + SongPath + "|" + SongSpell + "|" + SongSpellNum + "|" + SongSongStroke + "|" + SongPenStyle + "|" + SongPlayState + "|" + SongSrcPath + "|" + SongReplayGain;
        }

        static string SelectLayout(SongRecord s) =>
            string.Join("|", s.Id, s.Lang, s.SingerType, s.Singer, s.SongName, s.Track, s.SongType, s.Volume, s.WordCount, s.PlayCount, s.MB, s.CreatDate, s.FileName, s.Path, s.Spell, s.SpellNum, s.SongStroke, s.PenStyle, s.PlayState, s.SrcPath, s.ReplayGain);

        [Fact]
        public void Query_grid_row_reads_the_same_values_as_the_old_selection_string()
        {
            DataGridView grid = Grid(QueryTable());
            foreach (DataGridViewRow row in Rows(grid))
            {
                Assert.Equal(LegacyQuerySelect(row), SelectLayout(SongRecord.FromQueryGridRow(row)));
            }
        }

        [Fact]
        public void Add_grid_row_reads_the_same_values_as_the_old_selection_string()
        {
            DataGridView grid = Grid(AddTable());
            foreach (DataGridViewRow row in Rows(grid))
            {
                SongRecord song = SongRecord.FromAddGridRow(row);
                Assert.Equal(LegacyAddSelect(row), SelectLayout(song));
                Assert.Equal(row.Cells["Song_AddStatus"].Value.ToString(), song.AddStatus);
            }
        }

        [Fact]
        public void Player_fields_match_the_old_player_list()
        {
            DataGridView grid = Grid(QueryTable());
            foreach (DataGridViewRow row in Rows(grid))
            {
                // SongQueryDataGridView.cs 播放檔案 menu
                List<string> legacy = new List<string>()
                {
                    row.Cells["Song_Id"].Value.ToString(), row.Cells["Song_Lang"].Value.ToString(), row.Cells["Song_Singer"].Value.ToString(),
                    row.Cells["Song_SongName"].Value.ToString(), row.Cells["Song_Track"].Value.ToString(), row.Cells["Song_Volume"].Value.ToString(),
                    row.Cells["Song_ReplayGain"].Value.ToString(),
                };
                SongRecord s = SongRecord.ForPlayer(row);
                Assert.Equal(legacy, new List<string> { s.Id, s.Lang, s.Singer, s.SongName, s.Track, s.Volume, s.ReplayGain });
            }
        }

        // ---------- log lines: the 操作記錄 still prints the same text ----------

        static SongRecord Sample() => new SongRecord
        {
            Id = "000123", Lang = "國語", SingerType = "0", Singer = "周杰倫", SongName = "晴天", Track = "1", SongType = "演唱會",
            Volume = "87", WordCount = "2", PlayCount = "15", MB = "12.34", CreatDate = "2024/01/02 03:04:05", FileName = "周杰倫_晴天_VR.mpg",
            Path = @"D:\KTV\國語\男\周杰倫\", Spell = "ㄑㄊ", SpellNum = "42", SongStroke = "12", PenStyle = "丨一", PlayState = "0",
            ReplayGain = "-3.2", OldId = "000099", AddSinger = "1", AddAllSinger = "0",
        };

        [Fact]
        public void Log_lines_match_the_old_update_add_and_replace_strings()
        {
            SongRecord r = Sample();
            string SongId = r.Id, SongLang = r.Lang, SongSinger = r.Singer, SongSongName = r.SongName, SongSongType = r.SongType,
                SongMB = r.MB, SongCreatDate = r.CreatDate, SongFileName = r.FileName, SongPath = r.Path, SongSpell = r.Spell, SongSpellNum = r.SpellNum,
                SongPenStyle = r.PenStyle, OldSongId = r.OldId, SongAddSinger = r.AddSinger, SongAddAllSinger = r.AddAllSinger, SongReplayGain = r.ReplayGain;
            int SongSingerType = 0, SongTrack = 1, SongVolume = 87, SongWordCount = 2, SongPlayCount = 15, SongSongStroke = 12, SongPlayState = 0;

            // SongQuery.cs SongQuery_SongUpdate (SongUpdateValue)
            string update = SongId + "|" + SongLang + "|" + SongSingerType + "|" + SongSinger + "|" + SongSongName + "|" + SongTrack + "|" + SongSongType + "|" + SongVolume + "|" + SongWordCount + "|" + SongPlayCount + "|" + SongMB + "|" + SongCreatDate + "|" + SongFileName + "|" + SongPath + "|" + SongSpell + "|" + SongSpellNum + "|" + SongSongStroke + "|" + SongPenStyle + "|" + SongPlayState + "|" + OldSongId;
            // SongAddSong.cs StartAddSong (SongAddValue)
            string add = SongId + "|" + SongLang + "|" + SongSingerType + "|" + SongSinger + "|" + SongSongName + "|" + SongTrack + "|" + SongSongType + "|" + SongVolume + "|" + SongWordCount + "|" + SongPlayCount + "|" + SongMB + "|" + SongCreatDate + "|" + SongFileName + "|" + SongPath + "|" + SongSpell + "|" + SongSpellNum + "|" + SongSongStroke + "|" + SongPenStyle + "|" + SongPlayState + "|" + SongAddSinger + "|" + SongAddAllSinger + "|" + SongReplayGain;
            // SongAddSong.cs StartUpdateSong (SongAddValue)
            string replace = SongId + "|" + SongLang + "|" + SongSingerType + "|" + SongSinger + "|" + SongSongName + "|" + SongTrack + "|" + SongSongType + "|" + SongVolume + "|" + SongWordCount + "|" + SongPlayCount + "|" + SongMB + "|" + SongCreatDate + "|" + SongFileName + "|" + SongPath + "|" + SongSpell + "|" + SongSpellNum + "|" + SongSongStroke + "|" + SongPenStyle + "|" + SongPlayState + "|" + SongReplayGain;

            Assert.Equal(update, r.ToLogString());
            Assert.Equal(add, r.ToAddLogString());
            Assert.Equal(replace, r.ToReplaceLogString());
        }

        [Fact]
        public void Add_pipeline_formats_numbers_and_dates_like_the_old_string()
        {
            // StartAddSong concatenated typed values; the record must hold the same text.
            int SongSingerType = 3, SongTrack = 2, SongVolume = 100, SongWordCount = 4, SongPlayCount = 0, SongSongStroke = 9, SongPlayState = 0;
            float SongMB = 1234.5f;
            DateTime SongCreatDate = new DateTime(2026, 10, 31, 13, 5, 9);
            string legacy = "" + SongSingerType + "|" + SongTrack + "|" + SongVolume + "|" + SongWordCount + "|" + SongPlayCount + "|" + SongMB + "|" + SongCreatDate.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture) + "|" + SongSongStroke + "|" + SongPlayState;

            SongRecord r = new SongRecord
            {
                SingerType = SongSingerType.ToString(), Track = SongTrack.ToString(), Volume = SongVolume.ToString(), WordCount = SongWordCount.ToString(),
                PlayCount = SongPlayCount.ToString(), MB = SongMB.ToString(), CreatDate = SongCreatDate.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture),
                SongStroke = SongSongStroke.ToString(), PlayState = SongPlayState.ToString(),
            };
            Assert.Equal(legacy, string.Join("|", r.SingerType, r.Track, r.Volume, r.WordCount, r.PlayCount, r.MB, r.CreatDate, r.SongStroke, r.PlayState));
        }

        // ---------- writing an edit back into the grid's table ----------

        static List<string> Dump(DataRow row) => row.Table.Columns.Cast<DataColumn>().Select(c => c.ColumnName + "=" + row[c] + ":" + row[c].GetType().Name).ToList();

        [Fact]
        public void Writing_an_edit_back_sets_the_same_cells_as_the_old_code()
        {
            // 歌庫查詢: SongQuery_SongUpdate wrote dtvaluelist[0..18]
            DataTable q = QueryTable();
            DataRow a = q.Rows[0], b = q.Copy().Rows[0];
            SongRecord r = Sample();
            string[] dtvaluelist = { r.Id, r.Lang, r.SingerType, r.Singer, r.SongName, r.Track, r.SongType, r.Volume, r.WordCount, r.PlayCount, r.MB, r.CreatDate, r.FileName, r.Path, r.Spell, r.SpellNum, r.SongStroke, r.PenStyle, r.PlayState };
            string[] cols = { "Song_Id", "Song_Lang", "Song_SingerType", "Song_Singer", "Song_SongName", "Song_Track", "Song_SongType", "Song_Volume", "Song_WordCount", "Song_PlayCount", "Song_MB", "Song_CreatDate", "Song_FileName", "Song_Path", "Song_Spell", "Song_SpellNum", "Song_SongStroke", "Song_PenStyle", "Song_PlayState" };
            for (int i = 0; i < cols.Length; i++) a[cols[i]] = dtvaluelist[i];
            r.WriteTo(b, false);
            Assert.Equal(Dump(a), Dump(b));

            // 加歌: SongAdd_EditApplyChanges wrote valuelist[0..15]
            DataTable t = AddTable();
            DataRow c = t.Rows[0], d = t.Copy().Rows[0];
            SongRecord e = Sample();
            e.AddStatus = "語系類別必須有值才能加歌!";
            e.SrcPath = @"F:\in\x.mpg";
            string[] addvalues = { e.AddStatus, e.Lang, e.SingerType, e.Singer, e.SongName, e.Track, e.SongType, e.Volume, e.WordCount, e.PlayCount, e.CreatDate, e.Spell, e.SpellNum, e.SongStroke, e.PenStyle, e.SrcPath };
            string[] addcols = { "Song_AddStatus", "Song_Lang", "Song_SingerType", "Song_Singer", "Song_SongName", "Song_Track", "Song_SongType", "Song_Volume", "Song_WordCount", "Song_PlayCount", "Song_CreatDate", "Song_Spell", "Song_SpellNum", "Song_SongStroke", "Song_PenStyle", "Song_SrcPath" };
            for (int i = 0; i < addcols.Length; i++) c[addcols[i]] = addvalues[i];
            e.WriteTo(d, true);
            Assert.Equal(Dump(c), Dump(d));
        }

        [Fact]
        public void A_value_containing_the_old_separator_survives()
        {
            // The old strings split on '|', so a 歌手 like "A|B" shifted every later field.
            SongRecord s = new SongRecord { Singer = "A|B", SongName = "歌", Track = "1" };
            SongRecord copy = s.Clone();
            Assert.Equal("A|B", copy.Singer);
            Assert.Equal("歌", copy.SongName);
            Assert.Equal(1, copy.TrackValue);
        }
    }
}

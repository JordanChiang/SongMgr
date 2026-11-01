using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Store
{
    /// <summary>
    /// Each write the features used to issue with their own SQL (copied from master @ 21fb779) runs on one
    /// scratch database; the same write through <see cref="SongLibraryStore"/> runs on a twin. The tables must
    /// end up identical, row for row and column for column.
    /// </summary>
    public class SongLibraryStoreDifferentialTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();
        readonly string root = TestEnv.NewTempDir("store");

        public void Dispose()
        {
            TestEnv.Restore(saved);
            TestEnv.DeleteDir(root);
        }

        // ---------- twin databases ----------

        (string legacySong, string legacyMgr, string moduleSong, string moduleMgr) Twins(string name)
        {
            string a = Path.Combine(root, name, "legacy"), b = Path.Combine(root, name, "module");
            Directory.CreateDirectory(a); Directory.CreateDirectory(b);
            TestEnv.UseScratchDatabases(a); Seed();
            string ls = Global.CrazyktvDatabaseFile, lm = Global.CrazyktvSongMgrDatabaseFile;
            TestEnv.UseScratchDatabases(b); Seed();
            return (ls, lm, Global.CrazyktvDatabaseFile, Global.CrazyktvSongMgrDatabaseFile);
        }

        static void Seed()
        {
            using (OleDbConnection conn = DbHelper.Access.OpenConn(Global.CrazyktvDatabaseFile))
            {
                foreach (SongRecord s in SeedSongs())
                {
                    using (OleDbCommand cmd = new OleDbCommand("insert into ktv_Song (Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState) values (@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s)", conn))
                    {
                        foreach (string v in new[] { s.Id, s.Lang, s.SingerType, s.Singer, s.SongName, s.Track, s.SongType, s.Volume, s.WordCount, s.PlayCount, s.MB, s.CreatDate, s.FileName, s.Path, s.Spell, s.SpellNum, s.SongStroke, s.PenStyle, s.PlayState })
                            cmd.Parameters.AddWithValue("@p", v);
                        cmd.ExecuteNonQuery();
                    }
                }
                using (OleDbCommand cmd = new OleDbCommand("insert into ktv_Favorite (User_Id, Song_Id) values ('0001', '000001')", conn)) cmd.ExecuteNonQuery();
            }
        }

        static IEnumerable<SongRecord> SeedSongs()
        {
            yield return Song("000001", "國語", "0", "周杰倫", "晴天", "1");
            yield return Song("000002", "國語", "1", "蔡依林", "日不落", "2");
            yield return Song("100001", "台語", "3", "江蕙&陳雷", "愛人醉落去", "0");
        }

        static SongRecord Song(string id, string lang, string singerType, string singer, string name, string track) => new SongRecord
        {
            Id = id, Lang = lang, SingerType = singerType, Singer = singer, SongName = name, Track = track, SongType = "",
            Volume = "100", WordCount = "2", PlayCount = "0", MB = "12.5", CreatDate = "2024/01/02 03:04:05",
            FileName = singer + "_" + name + ".mpg", Path = @"D:\KTV\" + lang + @"\", Spell = "ㄑㄊ", SpellNum = "42",
            SongStroke = "12", PenStyle = "丨一", PlayState = "0",
        };

        static string Dump(string db, string table, string order)
        {
            DataTable dt = Db.GetDataTable(db, "select * from " + table + " order by " + order, "");
            return string.Join("\n", dt.AsEnumerable().Select(r => string.Join("|", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName + "=" + (r[c] is DBNull ? "NULL" : r[c].ToString())))));
        }

        void AssertSame((string legacySong, string legacyMgr, string moduleSong, string moduleMgr) t)
        {
            Assert.Equal(Dump(t.legacySong, "ktv_Song", "Song_Id"), Dump(t.moduleSong, "ktv_Song", "Song_Id"));
            Assert.Equal(Dump(t.legacySong, "ktv_Favorite", "User_Id, Song_Id"), Dump(t.moduleSong, "ktv_Favorite", "User_Id, Song_Id"));
            Assert.Equal(Dump(t.legacySong, "ktv_Singer", "Singer_Id"), Dump(t.moduleSong, "ktv_Singer", "Singer_Id"));
            Assert.Equal(Dump(t.legacyMgr, "ktv_AllSinger", "Singer_Id"), Dump(t.moduleMgr, "ktv_AllSinger", "Singer_Id"));
        }

        static void Legacy(string db, string sql, params string[] values)
        {
            using (OleDbConnection conn = DbHelper.Access.OpenConn(db))
            using (OleDbCommand cmd = new OleDbCommand(sql, conn))
            {
                int i = 0;
                foreach (string v in values) cmd.Parameters.AddWithValue("@p" + i++, v);
                cmd.ExecuteNonQuery();
            }
        }

        static void Module(string db, Action<SongLibraryStore> write)
        {
            Global.CrazyktvDatabaseFile = db;
            using (SongLibraryStore store = SongLibraryStore.Open()) write(store);
        }

        const string AllColumnsSet = "Song_Id = @SongId, Song_Lang = @SongLang, Song_SingerType = @SongSingerType, Song_Singer = @SongSinger, Song_SongName = @SongSongName, Song_Track = @SongTrack, Song_SongType = @SongSongType, Song_Volume = @SongVolume, Song_WordCount = @SongWordCount, Song_PlayCount = @SongPlayCount, Song_MB = @SongMB, Song_CreatDate = @SongCreatDate, Song_FileName = @SongFileName, Song_Path = @SongPath, Song_Spell = @SongSpell, Song_SpellNum = @SongSpellNum, Song_SongStroke = @SongSongStroke, Song_PenStyle = @SongPenStyle, Song_PlayState = @SongPlayState";

        static string[] Core(SongRecord s) => new[] { s.Id, s.Lang, s.SingerType, s.Singer, s.SongName, s.Track, s.SongType, s.Volume, s.WordCount, s.PlayCount, s.MB, s.CreatDate, s.FileName, s.Path, s.Spell, s.SpellNum, s.SongStroke, s.PenStyle, s.PlayState };

        // ---------- the writes ----------

        [Fact]
        public void Insert_song_matches_SongAdd_and_SongDBConverter()
        {
            foreach (bool withGain in new[] { false, true })
            {
                var t = Twins("insert" + withGain);
                SongRecord s = Song("000777", "日語", "6", "Maroon 5", "Sugar", "4");
                s.ReplayGain = "-7.25";
                string cols = "Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState" + (withGain ? ", Song_ReplayGain" : "");
                string vals = string.Join(", ", cols.Split(',').Select(c => "@" + c.Trim().Replace("_", "")));
                Legacy(t.legacySong, "insert into ktv_Song ( " + cols + " ) values ( " + vals + " )", withGain ? Core(s).Concat(new[] { s.ReplayGain }).ToArray() : Core(s));
                Module(t.moduleSong, st => st.InsertSong(s, withGain ? SongColumns.All : SongColumns.AllButReplayGain));
                AssertSame(t);
            }
        }

        [Fact]
        public void Full_update_matches_SongQuery_edit_and_SongAdd_replace()
        {
            // 歌庫查詢: renumber + edit, keyed by the old id
            var t = Twins("query");
            SongRecord s = Song("000009", "粵語", "1", "蔡依林", "日不落(粵)", "5");
            s.OldId = "000002";
            Legacy(t.legacySong, "update ktv_Song set " + AllColumnsSet + " where Song_Id = @OldSongId", Core(s).Concat(new[] { s.OldId }).ToArray());
            Module(t.moduleSong, st => st.UpdateSong(s, SongColumns.AllButReplayGain));
            AssertSame(t);

            // 加歌 重複歌曲 replace with 音量偵測: keyed by its own id
            t = Twins("replace");
            SongRecord r = Song("000001", "國語", "0", "周杰倫", "晴天", "2");
            r.Volume = "64"; r.ReplayGain = "3.5";
            Legacy(t.legacySong, "update ktv_Song set " + AllColumnsSet + ", Song_ReplayGain = @SongReplayGain where Song_Id = @OldSongId", Core(r).Concat(new[] { r.ReplayGain, r.Id }).ToArray());
            Module(t.moduleSong, st => st.UpdateSong(r, SongColumns.All));
            AssertSame(t);
        }

        [Fact]
        public void Column_subset_updates_match_each_feature()
        {
            var cases = new (string name, string set, string[] legacyValues, SongRecord record, SongColumns columns)[]
            {
                ("spell", "Song_Spell = @SongSpell, Song_SpellNum = @SongSpellNum, Song_SongStroke = @SongSongStroke, Song_PenStyle = @SongPenStyle",
                    new[] { "ㄖㄅㄌ", "512", "4", "丨一一", "000002" }, new SongRecord { Id = "000002", Spell = "ㄖㄅㄌ", SpellNum = "512", SongStroke = "4", PenStyle = "丨一一" },
                    SongColumns.Spell | SongColumns.SpellNum | SongColumns.SongStroke | SongColumns.PenStyle),
                ("track", "Song_Track = @SongTrack, Song_FileName = @SongFileName",
                    new[] { "2", "周杰倫_晴天_VL.mpg", "000001" }, new SongRecord { Id = "000001", Track = "2", FileName = "周杰倫_晴天_VL.mpg" },
                    SongColumns.Track | SongColumns.FileName),
                ("volume", "Song_Volume = @SongVolume", new[] { "77", "100001" }, new SongRecord { Id = "100001", Volume = "77" }, SongColumns.Volume),
                ("path", "Song_Path = @SongPath", new[] { @"E:\New\", "000001" }, new SongRecord { Id = "000001", Path = @"E:\New\" }, SongColumns.Path),
                ("playcount", "Song_PlayCount = @SongPlayCount", new[] { "42", "000002" }, new SongRecord { Id = "000002", PlayCount = "42" }, SongColumns.PlayCount),
                ("wordcount", "Song_WordCount = @SongWordCount", new[] { "5", "000002" }, new SongRecord { Id = "000002", WordCount = "5" }, SongColumns.WordCount),
                ("rebuild", "Song_Id = @SongId, Song_SingerType = @SongSingerType, Song_Track = @SongTrack, Song_FileName = @SongFileName, Song_Path = @SongPath",
                    new[] { "100001", "3", "1", "x.mpg", @"G:\R\台語\", "100001" }, new SongRecord { Id = "100001", SingerType = "3", Track = "1", FileName = "x.mpg", Path = @"G:\R\台語\" },
                    SongColumns.Id | SongColumns.SingerType | SongColumns.Track | SongColumns.FileName | SongColumns.Path),
                ("gain", "Song_Volume = @SongVolume, Song_ReplayGain = @SongReplayGain", new[] { "55", "-1.5", "000001" }, new SongRecord { Id = "000001", Volume = "55", ReplayGain = "-1.5" },
                    SongColumns.Volume | SongColumns.ReplayGain),
                ("singer-sync", "Song_Singer = @SongSinger, Song_SingerType = @SongSingerType, Song_FileName = @SongFileName, Song_Path = @SongPath",
                    new[] { "周董", "0", "周董_晴天.mpg", @"D:\KTV\國語\周董\", "000001" }, new SongRecord { Id = "000001", Singer = "周董", SingerType = "0", FileName = "周董_晴天.mpg", Path = @"D:\KTV\國語\周董\" },
                    SongColumns.Singer | SongColumns.SingerType | SongColumns.FileName | SongColumns.Path),
                ("location", "Song_FileName = @SongFileName, Song_Path = @SongPath", new[] { "a.mpg", @"F:\b\", "000002" }, new SongRecord { Id = "000002", FileName = "a.mpg", Path = @"F:\b\" },
                    SongColumns.FileName | SongColumns.Path),
                ("cashbox-id", "Song_Id = @SongId", new[] { "000555", "000002" }, new SongRecord { Id = "000555", OldId = "000002" }, SongColumns.Id),
                ("missing-row", "Song_Volume = @SongVolume", new[] { "1", "999999" }, new SongRecord { Id = "999999", Volume = "1" }, SongColumns.Volume),
            };

            foreach (var c in cases)
            {
                var t = Twins(c.name);
                // Cashbox.cs renumbers with "where Song_Id = @OldSongId"; Jet binds repeated names to one value.
                string where = (c.name == "cashbox-id") ? " where Song_Id = @OldSongId" : " where Song_Id = @SongId";
                Legacy(t.legacySong, "update ktv_Song set " + c.set + where, c.legacyValues);
                int rows = 0;
                Module(t.moduleSong, st => rows = st.UpdateSong(c.record, c.columns));
                AssertSame(t);
                Assert.Equal(c.name == "missing-row" ? 0 : 1, rows);
            }
        }

        [Fact]
        public void Guarded_renumber_and_clear_gain_match_SongMaintenance()
        {
            foreach (string name in new[] { "晴天", "別首歌" })
            {
                var t = Twins("renumber" + name);
                Legacy(t.legacySong, "update ktv_Song set Song_Id = @SongId where Song_Id = @OldSongId and Song_SongName = @SongSongName and Song_Lang = @SongLang", "000100", "000001", name, "國語");
                Module(t.moduleSong, st => st.RenumberSong("000001", "000100", name, "國語"));
                AssertSame(t);
            }

            var g = Twins("cleargain");
            Legacy(g.legacySong, "update ktv_Song set Song_ReplayGain = 3");
            Legacy(g.moduleSong, "update ktv_Song set Song_ReplayGain = 3");
            Legacy(g.legacySong, "update ktv_Song set Song_ReplayGain = null");
            Module(g.moduleSong, st => st.ClearAllReplayGain());
            AssertSame(g);
        }

        [Fact]
        public void Delete_and_favorites_match_SongQuery_and_SongMonitor()
        {
            var t = Twins("delete");
            Legacy(t.legacySong, "delete from ktv_Song where Song_Id = @SongId", "000001");
            Legacy(t.legacySong, "delete from ktv_Favorite where Song_Id = @SongId", "000001");
            Module(t.moduleSong, st => { st.DeleteSong("000001"); st.DeleteFavorite("000001"); });
            AssertSame(t);

            t = Twins("favorite");
            Legacy(t.legacySong, "update ktv_Favorite set Song_Id = @SongId where Song_Id = @OldSongId", "000123", "000001");
            Module(t.moduleSong, st => st.RenumberFavorite("000001", "000123"));
            AssertSame(t);
        }

        [Fact]
        public void Singer_insert_matches_SongAdd_SingerMgr_and_rebuild()
        {
            foreach (SingerTable table in new[] { SingerTable.Library, SingerTable.AllSingers })
            {
                var t = Twins("singer" + table.Name);

                Global.CrazyktvDatabaseFile = t.legacySong; Global.CrazyktvSongMgrDatabaseFile = t.legacyMgr;
                List<string> spelllist = CommonFunc.GetSongNameSpell("新歌手甲");
                Legacy(table.DatabaseFile, "insert into " + table.Name + " ( Singer_Id, Singer_Name, Singer_Type, Singer_Spell, Singer_Strokes, Singer_SpellNum, Singer_PenStyle ) values ( @SingerId, @SingerName, @SingerType, @SingerSpell, @SingerStrokes, @SingerSpellNum, @SingerPenStyle )",
                    "99001", "新歌手甲", "1", spelllist[0], spelllist[2], spelllist[1], spelllist[3]);

                Global.CrazyktvDatabaseFile = t.moduleSong; Global.CrazyktvSongMgrDatabaseFile = t.moduleMgr;
                using (SongLibraryStore st = SongLibraryStore.Open(table)) st.InsertSinger(table, "99001", "新歌手甲", "1");

                AssertSame(t);
            }
        }

        [Fact]
        public void Batched_path_changes_commit_or_roll_back_like_SongMaintenance()
        {
            var t = Twins("transaction");
            Global.CrazyktvDatabaseFile = t.moduleSong;
            using (SongLibraryStore st = SongLibraryStore.Open())
            {
                st.BeginTransaction();
                st.UpdateSong(new SongRecord { Id = "000001", Path = @"X:\rolled-back\" }, SongColumns.Path);
                st.Rollback();

                st.BeginTransaction();
                st.UpdateSong(new SongRecord { Id = "000002", Path = @"X:\committed\" }, SongColumns.Path);
                st.Commit();
            }
            string dump = Dump(t.moduleSong, "ktv_Song", "Song_Id");
            Assert.DoesNotContain("rolled-back", dump);
            Assert.Contains(@"X:\committed\", dump);
        }

        [Fact]
        public void A_failing_write_throws_like_ExecuteNonQuery()
        {
            var t = Twins("fail");
            SongRecord duplicate = Song("000001", "國語", "0", "x", "y", "1");
            Global.CrazyktvDatabaseFile = t.moduleSong;
            using (SongLibraryStore st = SongLibraryStore.Open())
            {
                Assert.ThrowsAny<OleDbException>(() => st.InsertSong(duplicate, SongColumns.AllButReplayGain));
                // the store keeps working after a failure
                Assert.Equal(1, st.UpdateSong(new SongRecord { Id = "000001", Volume = "9" }, SongColumns.Volume));
            }
        }
    }
}

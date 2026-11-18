using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;
using Step = CrazyKTV_SongMgr.SchemaUpgrade.Step;

namespace SongMgr.Tests.Store
{
    /// <summary>
    /// 資料庫結構更新 on copies of CrazySongEmpty.mdb, each made old in one way the way earlier versions of CrazyKTV /
    /// SongMgr left it: <see cref="SchemaUpgrade.Check"/> names the step, <see cref="SchemaUpgrade.Run"/> does it and only
    /// it, the songs and 我的最愛 come through unchanged, and a failed step puts the file back.
    /// </summary>
    public class SchemaUpgradeTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("schema");
        string Backups => Path.Combine(root, "Backup");

        public void Dispose()
        {
            System.Data.SQLite.SQLiteConnection.ClearAllPools();
            TestEnv.DeleteDir(root);
        }

        // ---------- making a 歌庫 old ----------

        /// <summary>The SQL that gives a current 歌庫 the old shape <see cref="Step"/> fixes.</summary>
        static readonly Dictionary<Step, string[]> MakeOld = new Dictionary<Step, string[]>
        {
            [Step.DropAllSinger] = new[] { "create table ktv_AllSinger (Singer_Id TEXT(10), Singer_Name TEXT(60))" },
            [Step.DropVersion] = new[] { "create table ktv_Version (Id INTEGER, SongDB TEXT(10))" },
            [Step.WidenSong] = new[] { "alter table ktv_Song alter column Song_SongName TEXT(40) WITH COMPRESSION", "alter table ktv_Song alter column Song_FileName TEXT(120) WITH COMPRESSION" },
            [Step.WidenSinger] = new[] { "alter table ktv_Singer alter column Singer_Name TEXT(30) WITH COMPRESSION" },
            [Step.WidenPhonetics] = new[] { "alter table ktv_Phonetics alter column PenStyle TEXT(20) WITH COMPRESSION" },
            [Step.AddLanguageKeyWord] = new[] { "alter table ktv_Langauage drop column Langauage_KeyWord" },
            [Step.WidenUserId] = new[] { "alter table ktv_User alter column User_Id TEXT(8) WITH COMPRESSION", "alter table ktv_Favorite alter column User_Id TEXT(8) WITH COMPRESSION" },
            [Step.AddReplayGain] = new[] { "alter table ktv_Song drop column Song_ReplayGain" },
            [Step.DropMeanVolume] = new[] { "alter table ktv_Song add column Song_MeanVolume DOUBLE" },
            [Step.DropGodLiu] = new[] { "alter table ktv_Song add column cashboxsongid TEXT(10)", "create index cashboxsongid on ktv_Song (cashboxsongid)", "alter table ktv_Song add column imgpath TEXT(50)" },
            [Step.RenameChorusSwan] = new[] { "update ktv_Swan set Swan_Name = '合唱歌曲' where Swan_Id = 3" },
        };

        public static IEnumerable<object[]> Steps() => Enum.GetNames(typeof(Step)).Select(s => new object[] { s });

        string Library(string name, params Step[] old)
        {
            string dir = Path.Combine(root, name);
            Directory.CreateDirectory(dir);
            string db = Path.Combine(dir, "CrazySong.mdb");
            File.Copy(Path.Combine(TestEnv.FixtureDir, "CrazySongEmpty.mdb"), db);
            File.SetAttributes(db, FileAttributes.Normal);
            // Columns are narrowed while ktv_Song is still empty, as an old 歌庫 was created; then the songs go in.
            foreach (Step step in old)
            {
                foreach (string sql in MakeOld[step]) Execute(db, sql);
            }
            Seed(db);
            return db;
        }

        static void Seed(string db)
        {
            foreach (var (id, lang, singer, name) in new[] { ("100001", "國語", "周杰倫", "晴天"), ("100002", "國語", "王菲", "紅豆"), ("200001", "台語", "江蕙", "家後") })
            {
                Execute(db, "insert into ktv_Song (Song_Id, Song_Lang, Song_SingerType, Song_Singer, Song_SongName, Song_Track, Song_SongType, Song_Volume, Song_WordCount, Song_PlayCount, Song_MB, Song_CreatDate, Song_FileName, Song_Path, Song_Spell, Song_SpellNum, Song_SongStroke, Song_PenStyle, Song_PlayState) values (@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s)",
                    id, lang, "0", singer, name, "1", "", "100", "2", "0", "12.5", "2024/01/02 03:04:05", singer + "_" + name + ".mpg", @"D:\KTV\" + lang + @"\", "ㄑㄊ", "42", "12", "丨一", "0");
            }
            Execute(db, "insert into ktv_Singer (Singer_Id, Singer_Name, Singer_Type) values ('1', '周杰倫', '0')");
            Execute(db, "insert into ktv_User (User_Id, User_Name) values ('0001', '測試')");
            Execute(db, "insert into ktv_Favorite (User_Id, Song_Id) values ('0001', '100002')");
        }

        static void Execute(string db, string sql, params string[] values)
        {
            using (var conn = DbHelper.Access.OpenConn(db))
            using (var cmd = new OleDbCommand(sql, conn))
            {
                foreach (string v in values) cmd.Parameters.AddWithValue("@p", v);
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>The data every step must leave alone.</summary>
        static string Data(string db)
        {
            string Rows(string sql) => string.Join("\n", Db.GetDataTable(db, sql, "").AsEnumerable().Select(r => string.Join("|", r.ItemArray)));
            return Rows("select Song_Id, Song_Lang, Song_Singer, Song_SongName, Song_FileName, Song_Path, Song_Volume from ktv_Song order by Song_Id")
                + "\n" + Rows("select Singer_Id, Singer_Name from ktv_Singer order by Singer_Id")
                + "\n" + Rows("select User_Id, Song_Id from ktv_Favorite order by Song_Id");
        }

        // ---------- the checks ----------

        [Fact]
        public void A_current_歌庫_needs_nothing_and_Run_leaves_it_alone()
        {
            string db = Library("current");
            byte[] before = File.ReadAllBytes(db);

            Assert.Empty(SchemaUpgrade.Check(db));
            Assert.Empty(SchemaUpgrade.Run(db, Backups));

            Assert.Equal(before, File.ReadAllBytes(db));
            Assert.False(Directory.Exists(Backups), "nothing to do, so no backup either");
        }

        [Theory]
        [MemberData(nameof(Steps))]
        public void Each_old_shape_is_found_and_fixed(string name)
        {
            var step = (Step)Enum.Parse(typeof(Step), name);
            string db = Library(step.ToString(), step);
            string data = Data(db);

            Assert.Equal(new[] { step }, SchemaUpgrade.Check(db));
            Assert.Equal(new[] { step }, SchemaUpgrade.Run(db, Backups));

            Assert.Empty(SchemaUpgrade.Check(db));
            Assert.Equal(data, Data(db));
            Assert.Single(Directory.GetFiles(Backups, "*_CrazySong.mdb"));
        }

        [Fact]
        public void A_歌庫_old_in_every_way_is_brought_up_to_date_in_one_run()
        {
            Step[] all = Enum.GetValues(typeof(Step)).Cast<Step>().ToArray();
            string db = Library("all", all);
            string data = Data(db);

            Assert.Equal(all, SchemaUpgrade.Check(db));
            Assert.Equal(all, SchemaUpgrade.Run(db, Backups));

            Assert.Empty(SchemaUpgrade.Check(db));
            Assert.Equal(data, Data(db));
            Assert.Equal("合唱", Db.GetDataTable(db, "select Swan_Name from ktv_Swan where Swan_Id = 3", "").Rows[0][0]);
        }

        [Fact]
        public void A_failed_step_puts_the_file_back()
        {
            string db = Library("fails", Step.AddReplayGain, Step.WidenSong);
            // Something in the way of parking the songs: WidenSong cannot run.
            Execute(db, "create table ktv_Song_temp (x TEXT(10))");
            string data = Data(db);
            byte[] before = File.ReadAllBytes(db);

            var ex = Assert.Throws<SchemaUpgradeException>(() => SchemaUpgrade.Run(db, Backups));

            Assert.Equal(Step.WidenSong, ex.Step);
            Assert.Equal("更新歌曲資料表失敗,已還原為原本的資料庫檔案。", ex.Message);
            Assert.Equal(before, File.ReadAllBytes(db));
            Assert.Equal(new[] { Step.WidenSong, Step.AddReplayGain }, SchemaUpgrade.Check(db));
            Assert.Equal(data, Data(db));
        }

        [Fact]
        public void A_SQLite_歌庫_is_checked_without_text_lengths_and_not_upgraded()
        {
            string mdb = Library("sqlite", Step.WidenSong, Step.DropMeanVolume);
            string db = Path.ChangeExtension(mdb, ".db");
            var results = SqliteExport.Run(new[] { new SqliteExport.Job(mdb, db, SqliteExport.CrazySongDBVer) });
            Assert.True(results.Single().Succeeded, results.Single().Error?.ToString());

            Assert.Equal(new[] { Step.DropMeanVolume }, SchemaUpgrade.Check(db));
            Assert.Throws<InvalidOperationException>(() => SchemaUpgrade.Run(db, Backups));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 資料庫結構更新: brings a 歌庫 database (CrazySong.mdb) up to the tables, columns and data SongMgr needs.
    /// <see cref="Check"/> lists what is missing by looking at the file itself, so no version number is kept anywhere;
    /// <see cref="Run"/> does exactly those steps, after a backup, and puts the backup back if one fails.
    /// </summary>
    static class SchemaUpgrade
    {
        public enum Step
        {
            DropAllSinger,
            DropVersion,
            WidenSong,
            WidenSinger,
            WidenPhonetics,
            AddLanguageKeyWord,
            WidenUserId,
            AddReplayGain,
            DropMeanVolume,
            DropGodLiu,
            RenameChorusSwan,
        }

        /// <summary>Columns of the GodLiu edition of CrazyKTV, dropped from ktv_Song.</summary>
        public static readonly string[] GodLiuColumns = { "Song_SongNameFuzzy", "Song_SingerFuzzy", "Song_FuzzyVer", "DLspace", "Epasswd", "imgpath", "cashboxsongid", "cashboxdat", "holidaysongid" };

        static readonly (string Table, string Column, int Length, Step Step)[] TextLengths =
        {
            ("ktv_Song", "Song_SongName", 80, Step.WidenSong),
            ("ktv_Song", "Song_Singer", 60, Step.WidenSong),
            ("ktv_Song", "Song_Spell", 80, Step.WidenSong),
            ("ktv_Song", "Song_FileName", 255, Step.WidenSong),
            ("ktv_Song", "Song_SpellNum", 80, Step.WidenSong),
            ("ktv_Song", "Song_PenStyle", 80, Step.WidenSong),
            ("ktv_Singer", "Singer_Name", 60, Step.WidenSinger),
            ("ktv_Singer", "Singer_Spell", 60, Step.WidenSinger),
            ("ktv_Singer", "Singer_SpellNum", 60, Step.WidenSinger),
            ("ktv_Singer", "Singer_PenStyle", 60, Step.WidenSinger),
            ("ktv_Phonetics", "PenStyle", 40, Step.WidenPhonetics),
            ("ktv_User", "User_Id", 12, Step.WidenUserId),
            ("ktv_Favorite", "User_Id", 12, Step.WidenUserId),
        };

        /// <summary>What a failed step reports; the backup has been put back by then.</summary>
        public static string FailureMessage(Step step)
        {
            switch (step)
            {
                case Step.DropAllSinger: return "移除 ktv_AllSinger 資料表失敗,已還原為原本的資料庫檔案。";
                case Step.DropVersion: return "移除 ktv_Version 資料表失敗,已還原為原本的資料庫檔案。";
                case Step.WidenSong: return "更新歌曲資料表失敗,已還原為原本的資料庫檔案。";
                case Step.WidenSinger: return "更新歌手資料表失敗,已還原為原本的資料庫檔案。";
                case Step.WidenPhonetics: return "更新拼音資料表失敗,已還原為原本的資料庫檔案。";
                case Step.AddLanguageKeyWord: return "更新語系資料表失敗,已還原為原本的資料庫檔案。";
                case Step.WidenUserId: return "更新我的最愛資料表失敗,已還原為原本的資料庫檔案。";
                case Step.AddReplayGain: return "加入 Song_ReplayGain 欄位失敗,已還原為原本的資料庫檔案。";
                case Step.DropMeanVolume: return "移除 Song_MeanVolume 欄位失敗,已還原為原本的資料庫檔案。";
                case Step.DropGodLiu: return "刪除 GodLiu 相關欄位失敗,已還原為原本的資料庫檔案。";
                default: return "變更歌手類別資料失敗,已還原為原本的資料庫檔案。";
            }
        }

        /// <summary>
        /// The steps <paramref name="databaseFile"/> still needs, in the order <see cref="Run"/> does them; empty when it
        /// is up to date. Text lengths are an Access matter: a SQLite 歌庫 is not checked for them.
        /// The file must already have SongMgr's tables (ktv_Song, ktv_Singer, ktv_Swan...).
        /// </summary>
        public static List<Step> Check(string databaseFile)
        {
            var steps = new List<Step>();
            List<string> tables = Db.GetTableList(databaseFile, "");
            List<string> songColumns = Db.GetColumnList(databaseFile, "", "ktv_Song");

            if (tables.Contains("ktv_AllSinger")) steps.Add(Step.DropAllSinger);
            if (tables.Contains("ktv_Version")) steps.Add(Step.DropVersion);

            if (!Db.IsSQLite(databaseFile))
            {
                using (DbConnection conn = Db.OpenConn(databaseFile, ""))
                {
                    foreach (var t in TextLengths)
                    {
                        if (steps.Contains(t.Step)) continue;
                        int? length = TextLength(conn, t.Table, t.Column);
                        if (length.HasValue && length.Value != t.Length) steps.Add(t.Step);
                    }
                }
            }

            if (!Db.GetColumnList(databaseFile, "", "ktv_Langauage").Contains("Langauage_KeyWord")) steps.Add(Step.AddLanguageKeyWord);
            if (!songColumns.Contains("Song_ReplayGain")) steps.Add(Step.AddReplayGain);
            if (songColumns.Contains("Song_MeanVolume")) steps.Add(Step.DropMeanVolume);
            if (GodLiuColumns.Any(songColumns.Contains)) steps.Add(Step.DropGodLiu);
            if (HasOldChorusSwan(databaseFile)) steps.Add(Step.RenameChorusSwan);

            // Keep Run's order whatever order the checks above found them in.
            steps.Sort();
            return steps;
        }

        static int? TextLength(DbConnection conn, string table, string column)
        {
            using (DataTable dt = conn.GetSchema("Columns", new[] { null, null, table, column }))
            {
                if (dt.Rows.Count == 0) return null;
                object length = dt.Rows[0]["CHARACTER_MAXIMUM_LENGTH"];
                return (length == DBNull.Value) ? (int?)null : Convert.ToInt32(length);
            }
        }

        /// <summary>歌手類別 3 used to be called 合唱歌曲; it is 合唱 now.</summary>
        static bool HasOldChorusSwan(string databaseFile)
        {
            using (DataTable dt = Db.GetDataTable(databaseFile, "select Swan_Id, Swan_Name from ktv_Swan", ""))
            {
                return dt.AsEnumerable().Any(r => Convert.ToString(r["Swan_Id"]) == "3" && Convert.ToString(r["Swan_Name"]) == "合唱歌曲");
            }
        }

        /// <summary>
        /// Does every step <see cref="Check"/> finds, on an Access 歌庫. Copies the file into <paramref name="backupFolder"/>
        /// first (as <c>&lt;date&gt;_CrazySong.mdb</c>); when a step fails the copy is put back and a
        /// <see cref="SchemaUpgradeException"/> names the step. Returns the steps done, empty when there was nothing to do.
        /// </summary>
        public static List<Step> Run(string databaseFile, string backupFolder)
        {
            if (Db.IsSQLite(databaseFile)) throw new InvalidOperationException("SQLite 歌庫無法更新資料庫結構");

            List<Step> steps = Check(databaseFile);
            if (steps.Count == 0) return steps;

            Directory.CreateDirectory(backupFolder);
            string backup = Path.Combine(backupFolder, DateTime.Now.ToLongDateString() + "_CrazySong.mdb");
            File.Copy(databaseFile, backup, true);

            Step current = steps[0];
            Exception failure = null;
            List<string> songColumns = Db.GetColumnList(databaseFile, "", "ktv_Song");
            using (DbConnection conn = Db.OpenConn(databaseFile, ""))
            {
                try
                {
                    foreach (Step step in steps)
                    {
                        current = step;
                        Do(conn, step, songColumns);
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            }

            if (failure != null)
            {
                // The connection is closed by now, so the file can be replaced.
                File.Copy(backup, databaseFile, true);
                throw new SchemaUpgradeException(current, failure);
            }
            return steps;
        }

        static void Do(DbConnection conn, Step step, List<string> songColumns)
        {
            switch (step)
            {
                case Step.DropAllSinger:
                    Execute(conn, "drop table ktv_AllSinger");
                    break;
                case Step.DropVersion:
                    Execute(conn, "drop table ktv_Version");
                    break;
                case Step.WidenSong:
                    // Access will not alter a column of a table with rows: park the songs, alter, put them back.
                    Execute(conn, "select * into ktv_Song_temp from ktv_Song");
                    Execute(conn, "delete * from ktv_Song");
                    Execute(conn, "alter table ktv_Song alter column Song_SongName TEXT(80) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Song alter column Song_Singer TEXT(60) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Song alter column Song_Spell TEXT(80) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Song alter column Song_FileName TEXT(255) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Song alter column Song_SpellNum TEXT(80) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Song alter column Song_PenStyle TEXT(80) WITH COMPRESSION");
                    Execute(conn, "insert into ktv_Song select * from ktv_Song_temp");
                    Execute(conn, "drop table ktv_Song_temp");
                    break;
                case Step.WidenSinger:
                    Execute(conn, "alter table ktv_Singer alter column Singer_Name TEXT(60) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Singer alter column Singer_Spell TEXT(60) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Singer alter column Singer_SpellNum TEXT(60) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Singer alter column Singer_PenStyle TEXT(60) WITH COMPRESSION");
                    break;
                case Step.WidenPhonetics:
                    Execute(conn, "alter table ktv_Phonetics alter column PenStyle TEXT(40) WITH COMPRESSION");
                    break;
                case Step.AddLanguageKeyWord:
                    Execute(conn, "alter table ktv_Langauage add column Langauage_KeyWord TEXT(255) WITH COMPRESSION");
                    break;
                case Step.WidenUserId:
                    Execute(conn, "alter table ktv_User alter column User_Id TEXT(12) WITH COMPRESSION");
                    Execute(conn, "alter table ktv_Favorite alter column User_Id TEXT(12) WITH COMPRESSION");
                    break;
                case Step.AddReplayGain:
                    Execute(conn, "alter table ktv_Song add column Song_ReplayGain DOUBLE");
                    break;
                case Step.DropMeanVolume:
                    Execute(conn, "alter table ktv_Song drop column Song_MeanVolume");
                    break;
                case Step.DropGodLiu:
                    var indexed = new HashSet<string>();
                    using (DataTable dt = conn.GetSchema("Indexes", new[] { null, null, null, null, "ktv_Song" }))
                    {
                        foreach (DataRow row in dt.Rows) indexed.Add(row["COLUMN_NAME"].ToString());
                    }
                    foreach (string column in GodLiuColumns.Where(songColumns.Contains))
                    {
                        if ((column == "cashboxsongid" || column == "holidaysongid") && indexed.Contains(column)) Execute(conn, "drop index " + column + " on ktv_Song");
                        Execute(conn, "alter table ktv_Song drop column " + column);
                    }
                    break;
                case Step.RenameChorusSwan:
                    using (DbCommand cmd = Db.Command("update ktv_Swan set Swan_Name = @SwanName where Swan_Id = @SwanId", conn))
                    {
                        cmd.Parameters.AddWithValue("@SwanName", "合唱");
                        cmd.Parameters.AddWithValue("@SwanId", "3");
                        cmd.ExecuteNonQuery();
                    }
                    break;
            }
        }

        static void Execute(DbConnection conn, string sql)
        {
            using (DbCommand cmd = Db.Command(sql, conn)) cmd.ExecuteNonQuery();
        }
    }

    /// <summary>A 資料庫結構更新 step failed; the 歌庫 database has been put back as it was.</summary>
    sealed class SchemaUpgradeException : Exception
    {
        public SchemaUpgrade.Step Step { get; }

        public SchemaUpgradeException(SchemaUpgrade.Step step, Exception inner) : base(SchemaUpgrade.FailureMessage(step), inner)
        {
            Step = step;
        }
    }
}

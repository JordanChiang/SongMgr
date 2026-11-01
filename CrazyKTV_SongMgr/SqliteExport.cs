using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 匯出 SQLite 資料庫: CrazySong.mdb and CrazySongMgr.mdb become CrazySong.db and CrazySongMgr.db, each
    /// beside its .mdb. UltimateKtv reads CrazySong.db instead of CrazySong.mdb once it exists next to its exe,
    /// and SongMgr opens SongMgr\CrazySongMgr.db instead of CrazySongMgr.mdb at its next start.
    /// Each database is exported to a temporary file first; a .db already there is replaced only when it
    /// differs and the user agrees (<see cref="Replace"/> / <see cref="Keep"/>).
    /// </summary>
    static class SqliteExport
    {
        /// <summary>Export format version of CrazySong.db. Bump it when that file's exported structure changes.</summary>
        public const int CrazySongDBVer = 1;

        /// <summary>Export format version of CrazySongMgr.db. Bump it when that file's exported structure changes.</summary>
        public const int CrazySongMgrDBVer = 1;

        /// <summary>At most this many differences are listed in the overwrite question.</summary>
        private const int MaxListedDifferences = 15;

        public sealed class Job
        {
            public string Source { get; }
            public string Dest { get; }
            public int DBVer { get; }
            public bool IsSongMgr { get; }
            public string TempFile => Dest + ".tmp";

            public Job(string source, string dest, int dbVer, bool isSongMgr = false) { Source = source; Dest = dest; DBVer = dbVer; IsSongMgr = isSongMgr; }
        }

        public enum Outcome
        {
            /// <summary>No .db was there; the export is now in place.</summary>
            Created,
            /// <summary>The .db was there with the same content; it is left as it was.</summary>
            Unchanged,
            /// <summary>The .db differs from the export, which waits in <see cref="Job.TempFile"/> for <see cref="Replace"/> or <see cref="Keep"/>.</summary>
            Different,
            /// <summary>The user chose to overwrite the differing .db.</summary>
            Replaced,
            /// <summary>The user chose to keep the differing .db.</summary>
            Kept,
            Failed,
        }

        public sealed class Result
        {
            public Job Job { get; }
            public List<(string Table, int Rows)> Tables { get; }
            /// <summary>How the existing .db differs from the export, when <see cref="Outcome"/> is Different / Replaced / Kept.</summary>
            public List<string> Differences { get; }
            public Outcome Outcome { get; internal set; }
            public Exception Error { get; internal set; }
            public bool Succeeded => Outcome != Outcome.Failed;

            public Result(Job job, Outcome outcome, List<(string Table, int Rows)> tables = null, List<string> differences = null, Exception error = null)
            {
                Job = job; Outcome = outcome; Tables = tables; Differences = differences ?? new List<string>(); Error = error;
            }
        }

        /// <summary>
        /// Both databases, each written as a .db beside its own file. A database that is already SQLite is
        /// left out: it has nothing to export, and its .db would be the file itself.
        /// </summary>
        public static List<Job> Jobs(string songDatabase, string songMgrDatabase)
        {
            var jobs = new List<Job>()
            {
                new Job(songDatabase, Path.ChangeExtension(songDatabase, ".db"), CrazySongDBVer, false),
                new Job(songMgrDatabase, Path.ChangeExtension(songMgrDatabase, ".db"), CrazySongMgrDBVer, true),
            };
            jobs.RemoveAll(j => Db.IsSQLite(j.Source));
            return jobs;
        }

        /// <summary>
        /// Exports each job to its temporary file, then puts it in place when no .db exists, drops it when the
        /// existing .db has the same content, and otherwise leaves it as <see cref="Outcome.Different"/>.
        /// A failed job does not stop the others.
        /// </summary>
        public static List<Result> Build(IEnumerable<Job> jobs)
        {
            var results = new List<Result>();
            foreach (Job job in jobs)
            {
                try
                {
                    List<(string Table, int Rows)> tables;
                    using (AccessExporter exporter = new AccessExporter(job.Source))
                    {
                        tables = exporter.export(job.TempFile, job.DBVer);
                    }

                    if (!File.Exists(job.Dest))
                    {
                        File.Move(job.TempFile, job.Dest);
                        results.Add(new Result(job, Outcome.Created, tables));
                        continue;
                    }

                    List<string> differences = Compare(job.Dest, job.TempFile);
                    if (differences.Count == 0)
                    {
                        File.Delete(job.TempFile);
                        results.Add(new Result(job, Outcome.Unchanged, tables));
                    }
                    else
                    {
                        results.Add(new Result(job, Outcome.Different, tables, differences));
                    }
                }
                catch (Exception ex)
                {
                    DeleteTemp(job);
                    results.Add(new Result(job, Outcome.Failed, error: ex));
                }
            }
            return results;
        }

        /// <summary><see cref="Build"/>, then overwrite every differing .db without asking.</summary>
        public static List<Result> Run(IEnumerable<Job> jobs)
        {
            List<Result> results = Build(jobs);
            foreach (Result result in results.Where(r => r.Outcome == Outcome.Different)) Replace(result);
            return results;
        }

        /// <summary>Overwrites the differing .db with the export.</summary>
        public static void Replace(Result result)
        {
            try
            {
                File.Move(result.Job.TempFile, result.Job.Dest, true);
                result.Outcome = Outcome.Replaced;
            }
            catch (Exception ex)
            {
                DeleteTemp(result.Job);
                result.Error = ex;
                result.Outcome = Outcome.Failed;
            }
        }

        /// <summary>Keeps the differing .db and drops the export.</summary>
        public static void Keep(Result result)
        {
            DeleteTemp(result.Job);
            result.Outcome = Outcome.Kept;
        }

        /// <summary>
        /// How <paramref name="existing"/> differs from <paramref name="exported"/>: export version, tables added
        /// or missing, table definitions, and table rows (compared as a whole, in any order). Empty when the same.
        /// </summary>
        public static List<string> Compare(string existing, string exported)
        {
            var differences = new List<string>();
            try
            {
                using (SQLiteConnection oldConn = OpenReadOnly(existing))
                using (SQLiteConnection newConn = OpenReadOnly(exported))
                {
                    long oldVer = Convert.ToInt64(Scalar(oldConn, "pragma user_version"));
                    long newVer = Convert.ToInt64(Scalar(newConn, "pragma user_version"));
                    if (oldVer != newVer) differences.Add("資料庫版本: " + oldVer + " → " + newVer);

                    Dictionary<string, string> oldTables = TableDefinitions(oldConn);
                    Dictionary<string, string> newTables = TableDefinitions(newConn);

                    foreach (string table in newTables.Keys.Concat(oldTables.Keys.Where(t => !newTables.ContainsKey(t))))
                    {
                        if (!oldTables.ContainsKey(table))
                        {
                            differences.Add(table + ": 現有檔案沒有此資料表");
                        }
                        else if (!newTables.ContainsKey(table))
                        {
                            differences.Add(table + ": 歌庫沒有此資料表 (覆蓋後會移除)");
                        }
                        else if (oldTables[table] != newTables[table])
                        {
                            differences.Add(table + ": 資料表結構不同");
                        }
                        else
                        {
                            List<string> oldRows = Rows(oldConn, table);
                            List<string> newRows = Rows(newConn, table);
                            if (!oldRows.SequenceEqual(newRows, StringComparer.Ordinal))
                            {
                                differences.Add(table + ": 資料不同 (現有 " + oldRows.Count + " 筆 → 匯出 " + newRows.Count + " 筆)");
                            }
                        }
                    }
                }
            }
            catch (SQLiteException ex)
            {
                differences.Add("無法讀取現有檔案: " + ex.Message);
            }
            return differences;
        }

        /// <summary>The question asked before overwriting a differing .db.</summary>
        public static string OverwritePrompt(Result result)
        {
            var sb = new StringBuilder();
            sb.AppendLine(result.Job.Dest);
            sb.AppendLine();
            sb.AppendLine("此檔案已存在,且內容與歌庫匯出的結果不同:");
            foreach (string difference in result.Differences.Take(MaxListedDifferences)) sb.AppendLine("    " + difference);
            if (result.Differences.Count > MaxListedDifferences) sb.AppendLine("    ...等共 " + result.Differences.Count + " 項差異");

            if (!result.Job.IsSongMgr)
            {
                sb.AppendLine();
                sb.AppendLine("注意: 播放器寫進 " + Path.GetFileName(result.Job.Dest) + " 的點播次數、YouTube 紀錄、我的最愛,會被 " + Path.GetFileName(result.Job.Source) + " 的內容取代。");
            }
            sb.AppendLine();
            sb.Append("是否要覆蓋?");
            return sb.ToString();
        }

        /// <summary>The message shown when the export ends: what happened to each file, with its table row counts, or why it failed.</summary>
        public static string Describe(List<Result> results)
        {
            var sb = new StringBuilder();
            foreach (Result result in results)
            {
                if (sb.Length > 0) sb.AppendLine();
                switch (result.Outcome)
                {
                    case Outcome.Created:
                    case Outcome.Replaced:
                        sb.AppendLine(((result.Outcome == Outcome.Created) ? "已建立: " : "已覆蓋: ") + result.Job.Dest + " (資料庫版本 " + result.Job.DBVer + ")");
                        foreach (var table in result.Tables) sb.AppendLine("    " + table.Table + ": " + table.Rows + " 筆");
                        if (result.Job.IsSongMgr) sb.AppendLine("    SongMgr 下次啟動時會改用此檔案。");
                        break;
                    case Outcome.Unchanged:
                        sb.AppendLine("內容相同,未變更: " + result.Job.Dest);
                        break;
                    case Outcome.Kept:
                    case Outcome.Different:
                        sb.AppendLine("未覆蓋,保留現有檔案: " + result.Job.Dest);
                        break;
                    case Outcome.Failed:
                        sb.AppendLine("匯出失敗: " + result.Job.Dest);
                        sb.AppendLine("    " + result.Error.Message);
                        break;
                }
            }
            return sb.ToString().TrimEnd();
        }

        private static void DeleteTemp(Job job)
        {
            try { if (File.Exists(job.TempFile)) File.Delete(job.TempFile); } catch { }
        }

        private static SQLiteConnection OpenReadOnly(string file)
        {
            SQLiteConnection conn = new SQLiteConnection("Data Source=" + file + ";Version=3;Read Only=True;FailIfMissing=True;");
            conn.Open();
            return conn;
        }

        private static object Scalar(SQLiteConnection conn, string sql)
        {
            using (SQLiteCommand cmd = new SQLiteCommand(sql, conn)) return cmd.ExecuteScalar();
        }

        /// <summary>Table name → CREATE TABLE statement, in creation order.</summary>
        private static Dictionary<string, string> TableDefinitions(SQLiteConnection conn)
        {
            var tables = new Dictionary<string, string>();
            using (SQLiteCommand cmd = new SQLiteCommand("select name, sql from sqlite_master where type = 'table' and name not like 'sqlite_%' order by rowid", conn))
            using (SQLiteDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read()) tables[reader.GetString(0)] = reader.IsDBNull(1) ? "" : reader.GetString(1);
            }
            return tables;
        }

        /// <summary>
        /// Every row as one string of quote()d values, sorted. quote() keeps text, numbers, blobs and NULL apart
        /// without converting any value, so dates SQLite cannot parse compare as stored.
        /// </summary>
        private static List<string> Rows(SQLiteConnection conn, string table)
        {
            var columns = new List<string>();
            using (SQLiteCommand cmd = new SQLiteCommand("pragma table_info(" + Quote(table) + ")", conn))
            using (SQLiteDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read()) columns.Add(reader.GetString(1));
            }

            var rows = new List<string>();
            string select = "select " + string.Join(" || char(31) || ", columns.Select(c => "quote(" + Quote(c) + ")")) + " from " + Quote(table);
            using (SQLiteCommand cmd = new SQLiteCommand(select, conn))
            using (SQLiteDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read()) rows.Add(reader.GetString(0));
            }
            rows.Sort(StringComparer.Ordinal);
            return rows;
        }

        private static string Quote(string identifier) => "\"" + identifier.Replace("\"", "\"\"") + "\"";
    }
}

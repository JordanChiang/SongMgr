using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 匯出 SQLite 資料庫: CrazySong.mdb and CrazySongMgr.mdb become CrazySong.db and CrazySongMgr.db
    /// beside CrazySong.mdb. UltimateKtv reads CrazySong.db instead of CrazySong.mdb once it exists
    /// next to its exe, so an export replaces whatever the player has written to the .db since.
    /// </summary>
    static class SqliteExport
    {
        /// <summary>Export format version of CrazySong.db. Bump it when that file's exported structure changes.</summary>
        public const int CrazySongDBVer = 1;

        /// <summary>Export format version of CrazySongMgr.db. Bump it when that file's exported structure changes.</summary>
        public const int CrazySongMgrDBVer = 1;

        public sealed class Job
        {
            public string Source { get; }
            public string Dest { get; }
            public int DBVer { get; }

            public Job(string source, string dest, int dbVer) { Source = source; Dest = dest; DBVer = dbVer; }
        }

        public sealed class Result
        {
            public Job Job { get; }
            public List<(string Table, int Rows)> Tables { get; }
            public Exception Error { get; }
            public bool Succeeded => Error == null;

            public Result(Job job, List<(string Table, int Rows)> tables, Exception error) { Job = job; Tables = tables; Error = error; }
        }

        /// <summary>
        /// Both databases, written as .db files into the folder of <paramref name="songDatabase"/>. A database
        /// that is already SQLite is left out: it has nothing to export, and its .db would be the file itself.
        /// </summary>
        public static List<Job> Jobs(string songDatabase, string songMgrDatabase)
        {
            string dir = Path.GetDirectoryName(songDatabase);
            var jobs = new List<Job>()
            {
                new Job(songDatabase, Path.Combine(dir, Path.GetFileNameWithoutExtension(songDatabase) + ".db"), CrazySongDBVer),
                new Job(songMgrDatabase, Path.Combine(dir, Path.GetFileNameWithoutExtension(songMgrDatabase) + ".db"), CrazySongMgrDBVer),
            };
            jobs.RemoveAll(j => Db.IsSQLite(j.Source));
            return jobs;
        }

        /// <summary>Exports each job in turn; a failed job does not stop the others.</summary>
        public static List<Result> Run(IEnumerable<Job> jobs)
        {
            var results = new List<Result>();
            foreach (Job job in jobs)
            {
                try
                {
                    using (AccessExporter exporter = new AccessExporter(job.Source))
                    {
                        results.Add(new Result(job, exporter.export(job.Dest, job.DBVer), null));
                    }
                }
                catch (Exception ex)
                {
                    results.Add(new Result(job, null, ex));
                }
            }
            return results;
        }

        /// <summary>The overwrite question for the jobs whose .db already exists, or "" when none does.</summary>
        public static string OverwritePrompt(List<Job> jobs, string songDatabase)
        {
            var existing = jobs.Where(j => File.Exists(j.Dest)).ToList();
            if (existing.Count == 0) return "";

            var sb = new StringBuilder("以下檔案已存在,是否要覆蓋?" + Environment.NewLine + Environment.NewLine);
            foreach (Job job in existing) sb.AppendLine(job.Dest);

            Job songJob = existing.FirstOrDefault(j => j.Source == songDatabase);
            if (songJob != null)
            {
                sb.AppendLine();
                sb.Append("注意: 播放器寫進 " + Path.GetFileName(songJob.Dest) + " 的點播次數、YouTube 紀錄、我的最愛,會被 " + Path.GetFileName(songDatabase) + " 的內容取代。");
            }
            return sb.ToString().TrimEnd();
        }

        /// <summary>The message shown when the export ends: each file with its table row counts, or why it failed.</summary>
        public static string Describe(List<Result> results)
        {
            var sb = new StringBuilder();
            foreach (Result result in results)
            {
                if (sb.Length > 0) sb.AppendLine();
                if (result.Succeeded)
                {
                    sb.AppendLine("匯出完成: " + result.Job.Dest + " (資料庫版本 " + result.Job.DBVer + ")");
                    foreach (var table in result.Tables) sb.AppendLine("    " + table.Table + ": " + table.Rows + " 筆");
                }
                else
                {
                    sb.AppendLine("匯出失敗: " + result.Job.Dest);
                    sb.AppendLine("    " + result.Error.Message);
                }
            }
            return sb.ToString().TrimEnd();
        }
    }
}

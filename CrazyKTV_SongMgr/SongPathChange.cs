using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 變更歌曲路徑 (歌庫維護): 把歌曲路徑中的來源資料夾換成目的資料夾, 分批寫入歌庫.
    /// 計數: [0] 成功 (已 Commit), [1] 失敗 (寫入出錯, 或所在批次被 Rollback).
    /// </summary>
    static class SongPathChange
    {
        /// <summary>每批寫入的筆數, 避免 Access 單一交易過大.</summary>
        public const int BatchSize = 500;

        public readonly struct Change
        {
            public Change(string songId, string path)
            {
                SongId = songId;
                Path = path;
            }

            public string SongId { get; }
            public string Path { get; }
        }

        /// <summary>統一成以反斜線結尾的資料夾路徑; 空白傳回 "".</summary>
        public static string NormalizeDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "";
            string normalized = path.Trim().Replace('/', '\\');
            normalized = Regex.Replace(normalized, @"(?<!^)\\{2,}", @"\");
            if (!normalized.EndsWith(@"\")) normalized += @"\";
            return normalized;
        }

        /// <summary>
        /// 找出路徑含 <paramref name="srcPath"/> (不分大小寫) 的歌曲, 算出換成 <paramref name="destPath"/> 後的路徑.
        /// <paramref name="songs"/> 需有 Song_Id, Song_Path 欄位.
        /// </summary>
        public static List<Change> Plan(DataTable songs, string srcPath, string destPath, JobContext ctx)
        {
            string src = NormalizeDirectory(srcPath);
            string dest = NormalizeDirectory(destPath);
            List<Change> changes = new List<Change>();

            foreach (DataRow row in songs.Rows)
            {
                string path = row.Field<string>("Song_Path") ?? "";
                if (path.IndexOf(src, StringComparison.OrdinalIgnoreCase) < 0) continue;

                string newPath = Regex.Replace(path, Regex.Escape(src), dest.Replace("$", "$$"), RegexOptions.IgnoreCase);
                if (!newPath.EndsWith(@"\")) newPath += @"\";
                changes.Add(new Change(row.Field<string>("Song_Id"), newPath));
                ctx.Report("正在解析第 " + changes.Count + " 首歌曲的歌曲路徑資料,請稍待...");
            }
            return changes;
        }

        /// <summary>
        /// 分批寫入 <paramref name="changes"/>. 成功數只在該批 Commit 後才計入; 被 Rollback 的批次整批算失敗.
        /// 只要有任何一筆寫入失敗, 最後一批就 Rollback (前面已 Commit 的批次保留).
        /// 交易本身出錯 (例如 Commit 失敗) 時 Rollback 目前的批次並丟出例外.
        /// </summary>
        public static void Apply(SongLibraryStore store, IReadOnlyList<Change> changes, JobContext ctx, int batchSize = BatchSize)
        {
            int pending = 0;        // 目前這批寫入成功, 還沒 Commit 的筆數
            bool hasError = false;
            bool inTransaction = false;

            try
            {
                store.BeginTransaction();
                inTransaction = true;

                for (int i = 0; i < changes.Count; i++)
                {
                    Change change = changes[i];
                    try
                    {
                        store.UpdateSong(new SongRecord { Id = change.SongId, Path = change.Path }, SongColumns.Path);
                        pending++;
                    }
                    catch (Exception ex)
                    {
                        hasError = true;
                        ctx.Counts.Increment(1);
                        Global.SongLog.Add("【變更歌曲路徑】更新資料庫時發生錯誤: " + change.Path + "|" + change.SongId + " - " + ex.Message);
                    }

                    bool lastOfBatch = (i + 1) % batchSize == 0;
                    bool last = i == changes.Count - 1;
                    if (last && hasError)
                    {
                        store.Rollback();
                        inTransaction = false;
                        if (pending > 0) Global.SongLog.Add("【變更歌曲路徑】因發生錯誤, 最後一批 " + pending + " 首歌曲的變更已取消。");
                        ctx.Counts.Increment(1, pending);
                        pending = 0;
                    }
                    else if (lastOfBatch || last)
                    {
                        store.Commit();
                        inTransaction = false;
                        ctx.Counts.Increment(0, pending);
                        pending = 0;
                        if (!last)
                        {
                            store.BeginTransaction();
                            inTransaction = true;
                        }
                    }

                    ctx.Report("正在轉換第 " + (i + 1) + " 首歌曲的歌曲路徑資料,請稍待...");
                }

                if (inTransaction)
                {
                    // 沒有任何要變更的歌曲
                    store.Commit();
                    inTransaction = false;
                }
            }
            catch
            {
                ctx.Counts.Increment(1, pending);
                if (inTransaction)
                {
                    try { store.Rollback(); } catch { }
                }
                throw;
            }
        }
    }
}

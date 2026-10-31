using System;
using System.Collections.Generic;
using System.IO;
using CrazyKTV_SongMgr;

namespace SongMgr.Tests.Placement
{
    /// <summary>
    /// Frozen copies of the file-moving blocks as they were before the 歌庫歸檔 module existed
    /// (master @ 21fb779). Only two edits: log writes go to <paramref name="log"/>, and the
    /// RemoveSong folder is a parameter instead of Application.StartupPath.
    /// The differential tests run these side by side with the module.
    /// </summary>
    internal static class LegacyPlacement
    {
        // SingerMgr.cs:446-503 (歌手管理, sync singer rename). Returns MoveError.
        public static bool SingerMgrRelocate(string SongSrcPath, string SongDestPath, List<string> log)
        {
            bool MoveError = false;
            string SongPath = Path.GetDirectoryName(SongDestPath) + @"\";
            string SongFileName = Path.GetFileName(SongDestPath);

            if (File.Exists(SongSrcPath))
            {
                if (!Directory.Exists(SongPath)) Directory.CreateDirectory(SongPath);

                if (File.Exists(SongDestPath))
                {
                    if (Path.GetFullPath(SongSrcPath).ToLower() == Path.GetFullPath(SongDestPath).ToLower())
                    {
                        try
                        {
                            FileAttributes attributes = File.GetAttributes(SongSrcPath);
                            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                                File.SetAttributes(SongSrcPath, attributes);
                            }

                            File.Move(SongSrcPath, SongPath + "Temp_" + SongFileName);
                            File.Move(SongPath + "Temp_" + SongFileName, SongDestPath);
                        }
                        catch
                        {
                            MoveError = true;
                            log.Add("【歌手管理】異動檔案時發生錯誤: " + SongSrcPath + " (檔案唯讀或正在使用)");
                        }
                    }
                    else
                    {
                        MoveError = true;
                        log.Add("【歌手管理】異動檔案時發生錯誤: " + SongSrcPath + " (歌庫裡已存在該首歌曲的檔案)");
                    }
                }
                else
                {
                    try
                    {
                        File.Move(SongSrcPath, SongDestPath);
                    }
                    catch
                    {
                        MoveError = true;
                        log.Add("【歌手管理】異動檔案時發生錯誤: " + SongSrcPath + " (檔案唯讀或正在使用)");
                    }
                }
            }
            else
            {
                MoveError = true;
                log.Add("【歌手管理】異動檔案時發生錯誤: " + SongSrcPath + " (檔案不存在)");
            }
            return MoveError;
        }

        // SongQuery.cs:954-1039 (歌庫查詢, edit song). Returns MoveError.
        public static bool SongQueryRelocate(string SongSrcPath, string SongDestPath, List<string> log)
        {
            bool MoveError = false;
            string SongPath = Path.GetDirectoryName(SongDestPath) + @"\";
            string SongFileName = Path.GetFileName(SongDestPath);

            if (File.Exists(SongSrcPath))
            {
                if (!Directory.Exists(SongPath)) Directory.CreateDirectory(SongPath);

                if (File.Exists(SongDestPath))
                {
                    if (Path.GetFullPath(SongSrcPath).ToLower() == Path.GetFullPath(SongDestPath).ToLower())
                    {
                        try
                        {
                            FileAttributes attributes = File.GetAttributes(SongSrcPath);
                            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                                File.SetAttributes(SongSrcPath, attributes);
                            }

                            File.Move(SongSrcPath, SongPath + "Temp_" + SongFileName);
                            File.Move(SongPath + "Temp_" + SongFileName, SongDestPath);
                        }
                        catch
                        {
                            MoveError = true;
                            log.Add("【歌庫查詢】異動檔案時發生錯誤: " + SongSrcPath + " (檔案唯讀或正在使用)");
                        }
                    }
                    else
                    {
                        MoveError = true;
                        log.Add("【歌庫查詢】異動檔案時發生錯誤: " + SongSrcPath + " (歌庫裡已存在該首歌曲的檔案)");
                    }
                }
                else
                {
                    try
                    {
                        File.Move(SongSrcPath, SongDestPath);
                    }
                    catch
                    {
                        MoveError = true;
                        log.Add("【歌庫查詢】異動檔案時發生錯誤: " + SongSrcPath + " (檔案唯讀或正在使用)");
                    }
                }
            }
            else
            {
                MoveError = true;
                log.Add("【歌庫查詢】異動檔案時發生錯誤: " + SongSrcPath + " (來源檔案不存在)");
            }
            return MoveError;
        }

        // SongAddSong.cs:501-567 (加歌, StartAddSong). Returns FileIOError; failures go to the failure log.
        public static bool AddSongImport(string SongSrcPath, string SongDestPath, string mode, bool backup, string removeSongDir, List<string> failures)
        {
            bool FileIOError = false;
            string SongPath = Path.GetDirectoryName(SongDestPath) + @"\";
            string SongFileName = Path.GetFileName(SongDestPath);

            if (!Directory.Exists(SongPath))
            {
                Directory.CreateDirectory(SongPath);
            }
            try
            {
                switch (mode)
                {
                    case "1":
                        if (File.Exists(SongDestPath))
                        {
                            FileAttributes attributes = File.GetAttributes(SongDestPath);
                            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                                File.SetAttributes(SongDestPath, attributes);
                            }

                            if (backup)
                            {
                                if (!Directory.Exists(removeSongDir)) Directory.CreateDirectory(removeSongDir);
                                if (File.Exists(removeSongDir + @"\" + SongFileName)) File.Delete(removeSongDir + @"\" + SongFileName);
                                File.Move(SongDestPath, removeSongDir + @"\" + SongFileName);
                                CommonFunc.SetFileTime(removeSongDir + @"\" + SongFileName, DateTime.Now);
                            }
                            else
                            {
                                File.Delete(SongDestPath);
                            }
                        }
                        File.Move(SongSrcPath, SongDestPath);
                        break;
                    case "2":
                        if (File.Exists(SongDestPath))
                        {
                            FileAttributes attributes = File.GetAttributes(SongDestPath);
                            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                                File.SetAttributes(SongDestPath, attributes);
                            }

                            if (backup)
                            {
                                if (!Directory.Exists(removeSongDir)) Directory.CreateDirectory(removeSongDir);
                                if (File.Exists(removeSongDir + @"\" + SongFileName)) File.Delete(removeSongDir + @"\" + SongFileName);
                                File.Move(SongDestPath, removeSongDir + @"\" + SongFileName);
                                CommonFunc.SetFileTime(removeSongDir + @"\" + SongFileName, DateTime.Now);
                            }
                        }
                        File.Copy(SongSrcPath, SongDestPath, true);
                        break;
                }
            }
            catch
            {
                FileIOError = true;
                failures.Add("檔案處理發生錯誤: " + SongDestPath + "(唯讀或使用中)");
            }
            return FileIOError;
        }

        // SongAddSong.cs:633-675 (加歌, StartUpdateSong: remove the existing duplicate's file). Returns DeleteError.
        public static bool UpdateSongRemoveOld(string oldfile, string SongSrcPath, bool backup, string removeSongDir, List<string> log)
        {
            bool DeleteError = false;
            string SongFileName = Path.GetFileName(oldfile);

            if (File.Exists(oldfile) && oldfile != SongSrcPath)
            {
                FileAttributes attributes = File.GetAttributes(oldfile);
                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                    File.SetAttributes(oldfile, attributes);
                }

                try
                {
                    if (backup)
                    {
                        if (!Directory.Exists(removeSongDir)) Directory.CreateDirectory(removeSongDir);
                        if (File.Exists(removeSongDir + @"\" + SongFileName)) File.Delete(removeSongDir + @"\" + SongFileName);
                        File.Move(oldfile, removeSongDir + @"\" + SongFileName);
                        CommonFunc.SetFileTime(removeSongDir + @"\" + SongFileName, DateTime.Now);
                    }
                    else
                    {
                        File.Delete(oldfile);
                    }
                }
                catch
                {
                    DeleteError = true;
                    log.Add("【加歌頁面】重複歌曲原有檔案是唯讀檔或正在使用中: " + oldfile);
                }
            }
            else
            {
                log.Add("【加歌頁面】重複歌曲原有檔案不存在或為同檔案,已自動忽略移除原有檔案: " + oldfile);
            }
            return DeleteError;
        }

        // SongAddSong.cs:710-777 (加歌, StartUpdateSong: place the replacement). Returns FileIOError.
        public static bool UpdateSongImport(string SongSrcPath, string SongDestPath, string mode, bool backup, string removeSongDir, List<string> failures)
        {
            bool FileIOError = AddSongImport(SongSrcPath, SongDestPath, mode, backup, removeSongDir, new List<string>());
            if (FileIOError) failures.Add("檔案處理發生錯誤: " + SongDestPath + " (唯讀或使用中)");
            return FileIOError;
        }

        // SongMaintenance.cs:2132-2160 (歌庫結構重建). Returns FileIOError.
        public static bool Rebuild(string SongSrcPath, string SongDestPath, string RebuildMode, List<string> log)
        {
            bool FileIOError = false;
            string SongPath = Path.GetDirectoryName(SongDestPath) + @"\";
            if (File.Exists(SongSrcPath))
            {
                if (!Directory.Exists(SongPath)) Directory.CreateDirectory(SongPath);

                try
                {
                    switch (RebuildMode)
                    {
                        case "1":
                            FileAttributes attributes = File.GetAttributes(SongSrcPath);
                            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                                File.SetAttributes(SongSrcPath, attributes);
                            }

                            if (File.Exists(SongDestPath)) File.Delete(SongDestPath);
                            File.Move(SongSrcPath, SongDestPath);
                            break;
                        case "2":
                            File.Copy(SongSrcPath, SongDestPath, true);
                            break;
                    }
                }
                catch
                {
                    FileIOError = true;
                    log.Add("【歌庫結構重建】檔案處理發生錯誤: " + SongSrcPath + " (唯讀或使用中)");
                }
            }
            else
            {
                FileIOError = true;
                log.Add("【歌庫結構重建】此首歌曲檔案不存在,已自動忽略重建檔案: " + SongSrcPath);
            }
            return FileIOError;
        }

        // SongMgrCfg.cs:991-1021 (歌庫結構更新). Returns true when the file was moved.
        public static bool UpdateStructure(string src, string dest, List<string> log)
        {
            string SongPath = Path.GetDirectoryName(dest) + @"\";
            if (src == dest) return true; // the caller skips SongSrcPath == SongDestPath (case-sensitive)

            if (File.Exists(src))
            {
                if (!Directory.Exists(SongPath)) Directory.CreateDirectory(SongPath);
                try
                {
                    FileAttributes attributes = File.GetAttributes(src);
                    if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                        File.SetAttributes(src, attributes);
                    }

                    if (File.Exists(dest)) File.Delete(dest);
                    File.Move(src, dest);
                    return true;
                }
                catch
                {
                    log.Add("【歌庫結構更新】檔案移動失敗: " + src + " (已忽略更新)");
                    return false;
                }
            }
            else
            {
                log.Add("【歌庫結構更新】原始檔案不存在: " + src + " (已忽略更新)");
                return false;
            }
        }

        // SongQuery.cs:1238-1275 (歌庫查詢, remove song file). Returns RemoveError.
        public static bool RemoveSongFile(string str, bool backup, string removeSongDir, List<string> log)
        {
            bool RemoveError = false;
            if (File.Exists(str))
            {
                try
                {
                    FileAttributes attributes = File.GetAttributes(str);
                    if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        attributes = CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly);
                        File.SetAttributes(str, attributes);
                    }

                    if (backup)
                    {
                        string SongFileName = Path.GetFileName(str);
                        if (!Directory.Exists(removeSongDir)) Directory.CreateDirectory(removeSongDir);
                        if (File.Exists(removeSongDir + @"\" + SongFileName)) File.Delete(removeSongDir + @"\" + SongFileName);
                        File.Move(str, removeSongDir + @"\" + SongFileName);
                        CommonFunc.SetFileTime(removeSongDir + @"\" + SongFileName, DateTime.Now);
                    }
                    else
                    {
                        File.Delete(str);
                    }
                }
                catch
                {
                    RemoveError = true;
                    log.Add("【歌庫查詢】異動檔案時發生錯誤: " + str + " (唯讀或使用中)");
                }
            }
            return RemoveError;
        }
    }
}

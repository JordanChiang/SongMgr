using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>Why a song is being placed; decides which folder the 歌庫結構 is rooted at.</summary>
    enum PlacementPurpose
    {
        /// <summary>加歌: a new file goes under the 歌庫 folder.</summary>
        Add,
        /// <summary>歌庫查詢 edit, 歌手管理 rename, 重複歌曲 replace: stay under the 歌庫 folder the file is already in.</summary>
        Relocate,
        /// <summary>歌庫結構重建: everything goes under a new, empty folder.</summary>
        Rebuild,
        /// <summary>歌庫結構更新: keep the part of the path above the 語系 folder.</summary>
        UpdateStructure,
    }

    enum PlacementOutcome
    {
        Placed,
        SourceMissing,
        /// <summary>Another file already sits at the destination; nothing was moved.</summary>
        DestinationOccupied,
        /// <summary>The file system refused (read-only, in use, …); nothing was moved.</summary>
        FileLocked,
    }

    /// <summary>
    /// 歌庫歸檔: where a song's file belongs under the 歌庫結構, and moving, copying,
    /// replacing or removing that file. Callers decide what to log; this module only reports the outcome.
    /// </summary>
    static class LibraryPlacement
    {
        /// <summary>備份移除歌曲 folder.</summary>
        public static string RemoveSongFolder = Application.StartupPath + @"\SongMgr\RemoveSong";

        /// <summary>True for 加歌模式 1 (move) and 2 (copy); modes 3 and 4 leave files where they are.</summary>
        public static bool MovesFiles => Global.AddMode != SongAddMode.InPlace && Global.AddMode != SongAddMode.Monitor;

        /// <summary>Both paths name the same file (Windows paths are case-insensitive).</summary>
        public static bool IsSameFile(string a, string b)
        {
            return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
        }

        #region --- Destination ---

        /// <summary>
        /// The full path <paramref name="song"/> should have under the current 歌庫結構 settings.
        /// Uses Id, Lang, SingerType, Singer, SongName, Track, SongType, FileName (for the extension) and Path.
        /// In 加歌模式 3 and 4 the current location is returned unchanged.
        /// </summary>
        public static string Destination(SongRecord song, PlacementPurpose purpose, string rebuildRoot = "")
        {
            string SongPath = song.Path;
            string SongFileName = song.FileName;
            if (!MovesFiles) return Path.Combine(SongPath, SongFileName);

            bool RebuildSongStructure = purpose == PlacementPurpose.Rebuild;
            bool DetectMultiSongPath = purpose == PlacementPurpose.Relocate || purpose == PlacementPurpose.UpdateStructure;
            bool UpdateStructure = purpose == PlacementPurpose.UpdateStructure;
            string RebuildSongPath = RebuildSongStructure ? rebuildRoot : "";

            string SongId = song.Id;
            string SongLang = song.Lang;
            int SongSingerType = song.SingerTypeValue;
            string SongSinger = song.Singer;
            string SongSongName = song.SongName;
            string SongSongType = song.SongType;

            string SongSingerStr = SongSinger;
            string SingerTypeStr = CommonFunc.GetSingerTypeStr(SongSingerType, 2, "null");
            string CrtchorusSeparate = (Global.SongMgrChorusSeparate == "1") ? "&" : "+";
            string SongInfoSeparate = (Global.SongMgrSongInfoSeparate == "1") ? "_" : "-";
            string SongTrackStr = CommonFunc.GetSongTrackStr(song.TrackValue, 1, "null");
            string SongExtension = Path.GetExtension(Path.Combine(SongPath, SongFileName));

            if (SongSingerType == 3) SongSingerStr = Regex.Replace(SongSinger, "[&+]", CrtchorusSeparate, RegexOptions.IgnoreCase);

            // 資料夾結構: 1 語系\歌手類別\歌手, 2 語系\歌手類別, 3 語系
            string folderTail;
            int levels;
            switch (Global.SongMgrFolderStructure)
            {
                case "1":
                    if (Global.SongMgrChorusMerge == "True" && SongSingerType == 3)
                    {
                        folderTail = SongLang + @"\" + SingerTypeStr + @"\";
                        levels = 2;
                    }
                    else
                    {
                        string fSongSingerStr = Regex.Replace(SongSingerStr, @"\.$", "", RegexOptions.IgnoreCase);
                        folderTail = SongLang + @"\" + SingerTypeStr + @"\" + fSongSingerStr + @"\";
                        levels = 3;
                    }
                    break;
                case "2":
                    folderTail = SongLang + @"\" + SingerTypeStr + @"\";
                    levels = 2;
                    break;
                case "3":
                    folderTail = SongLang + @"\";
                    levels = 1;
                    break;
                default:
                    folderTail = null;
                    levels = 0;
                    break;
            }

            if (folderTail != null)
            {
                if (!RebuildSongStructure && DetectMultiSongPath && CommonFunc.GetParentDirectoryPath(SongPath, levels) != Global.SongMgrDestFolder && !UpdateStructure)
                {
                    SongPath = CommonFunc.GetParentDirectoryPath(SongPath, levels) + @"\" + folderTail;
                }
                else if (UpdateStructure)
                {
                    SongPath = SongPath.Substring(0, SongPath.IndexOf(SongLang) - 1) + @"\" + folderTail;
                }
                else
                {
                    SongPath = ((!RebuildSongStructure) ? Global.SongMgrDestFolder : RebuildSongPath) + @"\" + folderTail;
                }
            }

            // 檔案結構: 1 歌手_歌名, 2 歌名_歌手, 3 編號_歌手_歌名 (資料夾結構 3 adds 歌手類別)
            string SongTypePart = (SongSongType == "") ? "" : SongSongType + SongInfoSeparate;
            string Tail = SongTypePart + SongTrackStr + SongExtension;
            bool ByLangOnly = Global.SongMgrFolderStructure == "3";
            bool ByFolder = Global.SongMgrFolderStructure == "1" || Global.SongMgrFolderStructure == "2";

            switch (Global.SongMgrFileStructure)
            {
                case "1":
                    if (ByFolder) SongFileName = SongSingerStr + SongInfoSeparate + SongSongName + SongInfoSeparate + Tail;
                    else if (ByLangOnly) SongFileName = SingerTypeStr + SongInfoSeparate + SongSingerStr + SongInfoSeparate + SongSongName + SongInfoSeparate + Tail;
                    break;
                case "2":
                    if (ByFolder) SongFileName = SongSongName + SongInfoSeparate + SongSingerStr + SongInfoSeparate + Tail;
                    else if (ByLangOnly) SongFileName = SongId + SongInfoSeparate + SingerTypeStr + SongInfoSeparate + SongSingerStr + SongInfoSeparate + SongSongName + SongInfoSeparate + Tail;
                    break;
                case "3":
                    if (ByFolder) SongFileName = SongId + SongInfoSeparate + SongSingerStr + SongInfoSeparate + SongSongName + SongInfoSeparate + Tail;
                    break;
            }

            return Path.Combine(SongPath, SongFileName);
        }

        #endregion

        #region --- File operations ---

        /// <summary>
        /// 歌庫查詢 edit / 歌手管理 rename: move a library file to <paramref name="dest"/> without
        /// overwriting a different file. A change of letter case only is done through a temporary name.
        /// </summary>
        public static PlacementOutcome Relocate(string src, string dest)
        {
            if (!File.Exists(src)) return PlacementOutcome.SourceMissing;

            string destDir = Path.GetDirectoryName(dest) + @"\";
            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

            if (File.Exists(dest))
            {
                if (!IsSameFile(src, dest)) return PlacementOutcome.DestinationOccupied;
                return RenameInPlace(src, dest) ? PlacementOutcome.Placed : PlacementOutcome.FileLocked;
            }

            try
            {
                File.Move(src, dest);
                return PlacementOutcome.Placed;
            }
            catch
            {
                return PlacementOutcome.FileLocked;
            }
        }

        /// <summary>
        /// 加歌 (加歌模式 1 move, 2 copy): put <paramref name="src"/> at <paramref name="dest"/>,
        /// first moving whatever is there to RemoveSong (<paramref name="backupReplaced"/>) or, when moving, deleting it.
        /// Returns false when the file system refused. Throws if the destination folder cannot be created.
        /// </summary>
        public static bool Import(string src, string dest, SongAddMode mode, bool backupReplaced)
        {
            string destDir = Path.GetDirectoryName(dest) + @"\";
            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

            try
            {
                if (File.Exists(dest) && IsSameFile(src, dest))
                {
                    // The file is already where it belongs: never back it up or delete it.
                    return RenameInPlace(src, dest);
                }

                switch (mode)
                {
                    case SongAddMode.Move:
                        if (File.Exists(dest))
                        {
                            ClearReadOnly(dest);
                            if (backupReplaced) MoveToRemoveSong(dest); else File.Delete(dest);
                        }
                        File.Move(src, dest);
                        break;
                    case SongAddMode.Copy:
                        if (File.Exists(dest))
                        {
                            ClearReadOnly(dest);
                            if (backupReplaced) MoveToRemoveSong(dest);
                        }
                        File.Copy(src, dest, true);
                        break;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 歌庫結構重建 / 歌庫結構更新: move (or copy) a library file to <paramref name="dest"/>,
        /// overwriting a different file that is already there.
        /// </summary>
        public static PlacementOutcome Overwrite(string src, string dest, bool copy)
        {
            if (!File.Exists(src)) return PlacementOutcome.SourceMissing;

            string destDir = Path.GetDirectoryName(dest) + @"\";
            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

            try
            {
                if (File.Exists(dest) && IsSameFile(src, dest))
                {
                    return RenameInPlace(src, dest) ? PlacementOutcome.Placed : PlacementOutcome.FileLocked;
                }

                if (copy)
                {
                    File.Copy(src, dest, true);
                }
                else
                {
                    ClearReadOnly(src);
                    if (File.Exists(dest)) File.Delete(dest);
                    File.Move(src, dest);
                }
                return PlacementOutcome.Placed;
            }
            catch
            {
                return PlacementOutcome.FileLocked;
            }
        }

        /// <summary>
        /// Takes a song file out of the 歌庫: moved to RemoveSong (<paramref name="backup"/>) or deleted.
        /// A missing file counts as removed. Returns false when the file system refused.
        /// </summary>
        public static bool Remove(string file, bool backup)
        {
            if (!File.Exists(file)) return true;
            try
            {
                ClearReadOnly(file);
                if (backup) MoveToRemoveSong(file); else File.Delete(file);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool RenameInPlace(string src, string dest)
        {
            try
            {
                ClearReadOnly(src);
                string temp = Path.GetDirectoryName(dest) + @"\" + "Temp_" + Path.GetFileName(dest);
                File.Move(src, temp);
                File.Move(temp, dest);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ClearReadOnly(string file)
        {
            FileAttributes attributes = File.GetAttributes(file);
            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                File.SetAttributes(file, CommonFunc.RemoveAttribute(attributes, FileAttributes.ReadOnly));
            }
        }

        private static void MoveToRemoveSong(string file)
        {
            string backup = RemoveSongFolder + @"\" + Path.GetFileName(file);
            if (!Directory.Exists(RemoveSongFolder)) Directory.CreateDirectory(RemoveSongFolder);
            if (File.Exists(backup)) File.Delete(backup);
            File.Move(file, backup);
            CommonFunc.SetFileTime(backup, DateTime.Now);
        }

        #endregion
    }
}

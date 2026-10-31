using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Placement
{
    /// <summary>
    /// Runs every 歌庫歸檔 file operation and its frozen pre-refactor copy on identical scratch folders
    /// and requires the same result, log lines and resulting files. The only scenarios allowed to differ
    /// are those where source and destination are the same file, where the old code destroyed it.
    /// </summary>
    public class FileOperationDifferentialTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("placement");
        readonly string savedRemoveSongFolder = LibraryPlacement.RemoveSongFolder;

        public void Dispose()
        {
            LibraryPlacement.RemoveSongFolder = savedRemoveSongFolder;
            TestEnv.DeleteDir(root);
        }

        // ---------- scenario model ----------

        public enum Dest { Absent, OtherFile, OtherFileReadOnly, SameFileExact, SameFileOtherCase }

        public sealed class Scenario
        {
            public bool SrcExists = true;
            public bool SrcReadOnly;
            public bool SrcLocked;
            public bool DestDirExists = true;
            public Dest Dest;
            public bool BackupExists;
            public override string ToString() =>
                $"src={(SrcExists ? "y" : "n")}{(SrcReadOnly ? "ro" : "")}{(SrcLocked ? "lock" : "")} destdir={(DestDirExists ? "y" : "n")} dest={Dest} backup={(BackupExists ? "y" : "n")}";
        }

        public static IEnumerable<Scenario> Scenarios()
        {
            foreach (bool srcExists in new[] { true, false })
            foreach (bool ro in new[] { false, true })
            foreach (bool locked in new[] { false, true })
            foreach (bool destDir in new[] { true, false })
            foreach (Dest dest in Enum.GetValues(typeof(Dest)))
            foreach (bool backup in new[] { false, true })
            {
                if (!srcExists && (ro || locked)) continue;
                if (!destDir && dest != Dest.Absent) continue;
                if (!srcExists && (dest == Dest.SameFileExact || dest == Dest.SameFileOtherCase)) continue;
                yield return new Scenario { SrcExists = srcExists, SrcReadOnly = ro, SrcLocked = locked, DestDirExists = destDir, Dest = dest, BackupExists = backup };
            }
        }

        static bool SameFile(Scenario s) => s.Dest == Dest.SameFileExact || s.Dest == Dest.SameFileOtherCase;

        sealed class Setup : IDisposable
        {
            public string Root, Src, Dst, RemoveSong;
            FileStream lockStream;

            public Setup(string root, Scenario s)
            {
                Root = root;
                Directory.CreateDirectory(root);
                string srcDir = Path.Combine(root, "lib", "國語", "男", "周杰倫");
                string destDir = Path.Combine(root, "lib", "國語", "男", "周董");
                RemoveSong = Path.Combine(root, "RemoveSong");
                Directory.CreateDirectory(srcDir);
                if (s.DestDirExists) Directory.CreateDirectory(destDir);

                Src = Path.Combine(srcDir, "周杰倫_晴天_VR.mpg");
                switch (s.Dest)
                {
                    case Dest.SameFileExact: Dst = Src; break;
                    case Dest.SameFileOtherCase: Dst = Path.Combine(srcDir, "周杰倫_晴天_vr.MPG"); break;
                    default: Dst = Path.Combine(destDir, "周董_晴天_VR.mpg"); break;
                }

                if (s.SrcExists) File.WriteAllText(Src, "SRC");
                if (s.Dest == Dest.OtherFile || s.Dest == Dest.OtherFileReadOnly) File.WriteAllText(Dst, "DEST");
                if (s.Dest == Dest.OtherFileReadOnly) File.SetAttributes(Dst, FileAttributes.ReadOnly);
                if (s.BackupExists)
                {
                    Directory.CreateDirectory(RemoveSong);
                    File.WriteAllText(Path.Combine(RemoveSong, Path.GetFileName(Dst)), "OLDBACKUP");
                    File.WriteAllText(Path.Combine(RemoveSong, Path.GetFileName(Src)), "OLDBACKUP");
                }
                if (s.SrcReadOnly) File.SetAttributes(Src, FileAttributes.ReadOnly);
                if (s.SrcLocked) lockStream = new FileStream(Src, FileMode.Open, FileAccess.Read, FileShare.None);
            }

            public void Dispose() => lockStream?.Dispose();

            /// <summary>Every file under the root: relative path (exact case), content, read-only flag.</summary>
            public List<string> Tree()
            {
                lockStream?.Dispose();
                lockStream = null;
                return Directory.GetFiles(Root, "*", SearchOption.AllDirectories)
                    .Select(f => f.Substring(Root.Length) + " = " + File.ReadAllText(f) + ((File.GetAttributes(f) & FileAttributes.ReadOnly) != 0 ? " [ro]" : ""))
                    .OrderBy(x => x, StringComparer.Ordinal).ToList();
            }

            public string Norm(string text) => text.Replace(Root, "<ROOT>");
        }

        /// <summary>Runs <paramref name="legacy"/> and <paramref name="module"/> on twin folders and compares everything.</summary>
        void Compare(string op, Scenario s, Func<Setup, string> legacy, Func<Setup, string> module, Action<Scenario, Setup> sameFileExpectation)
        {
            string a = Path.Combine(root, op, "legacy");
            string b = Path.Combine(root, op, "module");
            TestEnv.DeleteDir(a);
            TestEnv.DeleteDir(b);

            string legacyResult, moduleResult;
            List<string> legacyTree, moduleTree;
            using (Setup sa = new Setup(a, s))
            {
                legacyResult = sa.Norm(Run(() => legacy(sa)));
                legacyTree = sa.Tree();
            }
            using (Setup sb = new Setup(b, s))
            {
                LibraryPlacement.RemoveSongFolder = sb.RemoveSong;
                moduleResult = sb.Norm(Run(() => module(sb)));
                moduleTree = sb.Tree();

                if (SameFile(s))
                {
                    sameFileExpectation(s, sb);
                    return;
                }
            }

            Assert.True(legacyResult == moduleResult, $"{op} {s}: result\n legacy: {legacyResult}\n module: {moduleResult}");
            Assert.True(legacyTree.SequenceEqual(moduleTree), $"{op} {s}: files\n legacy: {string.Join(" ; ", legacyTree)}\n module: {string.Join(" ; ", moduleTree)}");

            TestEnv.DeleteDir(a);
            TestEnv.DeleteDir(b);
        }

        static string Run(Func<string> f)
        {
            try { return f(); }
            catch (Exception ex) { return "THROW " + ex.GetType().Name; }
        }

        /// <summary>
        /// The fix: when source and destination are one file it is never deleted or backed up; it ends up
        /// under the destination name, or keeps its own name when it is locked.
        /// </summary>
        static void KeptAtDestination(Scenario s, Setup set)
        {
            List<string> tree = set.Tree().Select(t => t.Replace(" [ro]", "")).ToList();
            string expected = (s.SrcLocked ? set.Src : set.Dst).Substring(set.Root.Length) + " = SRC";
            Assert.True(tree.Contains(expected), $"{s}: expected {expected} in {string.Join(" ; ", tree)}");
        }

        // ---------- 歌手管理 / 歌庫查詢: Relocate ----------

        static string RelocateMessage(string feature, string missingText, string src, PlacementOutcome o)
        {
            switch (o)
            {
                case PlacementOutcome.SourceMissing: return "【" + feature + "】異動檔案時發生錯誤: " + src + " " + missingText;
                case PlacementOutcome.DestinationOccupied: return "【" + feature + "】異動檔案時發生錯誤: " + src + " (歌庫裡已存在該首歌曲的檔案)";
                case PlacementOutcome.FileLocked: return "【" + feature + "】異動檔案時發生錯誤: " + src + " (檔案唯讀或正在使用)";
                default: return null;
            }
        }

        [Fact]
        public void Relocate_matches_SingerMgr_and_SongQuery()
        {
            foreach (Scenario s in Scenarios())
            {
                Compare("singer", s,
                    st => { var log = new List<string>(); bool err = LegacyPlacement.SingerMgrRelocate(st.Src, st.Dst, log); return err + " " + string.Join(";", log); },
                    st => { var o = LibraryPlacement.Relocate(st.Src, st.Dst); string m = RelocateMessage("歌手管理", "(檔案不存在)", st.Src, o); return (o != PlacementOutcome.Placed) + " " + m; },
                    (sc, st) => KeptAtDestination(sc, st));

                Compare("query", s,
                    st => { var log = new List<string>(); bool err = LegacyPlacement.SongQueryRelocate(st.Src, st.Dst, log); return err + " " + string.Join(";", log); },
                    st => { var o = LibraryPlacement.Relocate(st.Src, st.Dst); string m = RelocateMessage("歌庫查詢", "(來源檔案不存在)", st.Src, o); return (o != PlacementOutcome.Placed) + " " + m; },
                    (sc, st) => KeptAtDestination(sc, st));
            }
        }

        // ---------- 加歌: Import ----------

        [Fact]
        public void Import_matches_StartAddSong_and_StartUpdateSong()
        {
            foreach (Scenario s in Scenarios())
            foreach (string mode in new[] { "1", "2" })
            foreach (bool backup in new[] { false, true })
            {
                Compare("add" + mode + backup, s,
                    st => { var f = new List<string>(); bool err = LegacyPlacement.AddSongImport(st.Src, st.Dst, mode, backup, st.RemoveSong, f); return err + " " + string.Join(";", f); },
                    st => { bool ok = LibraryPlacement.Import(st.Src, st.Dst, mode == "1" ? SongAddMode.Move : SongAddMode.Copy, backup); return (!ok) + " " + (ok ? "" : "檔案處理發生錯誤: " + st.Dst + "(唯讀或使用中)"); },
                    (sc, st) => KeptAtDestination(sc, st));
            }
        }

        // ---------- 加歌 replace duplicate: remove old file ----------

        [Fact]
        public void Remove_matches_StartUpdateSong_old_file_removal_and_SongQuery_remove()
        {
            foreach (Scenario s in Scenarios())
            foreach (bool backup in new[] { false, true })
            {
                // The duplicate's existing file is Dst; the incoming file is Src.
                Compare("updremove" + backup, s,
                    st => { var log = new List<string>(); bool err = LegacyPlacement.UpdateSongRemoveOld(st.Dst, st.Src, backup, st.RemoveSong, log); return err + " " + string.Join(";", log); },
                    st =>
                    {
                        if (File.Exists(st.Dst) && !LibraryPlacement.IsSameFile(st.Dst, st.Src))
                        {
                            bool ok = LibraryPlacement.Remove(st.Dst, backup);
                            return (!ok) + " " + (ok ? "" : "【加歌頁面】重複歌曲原有檔案是唯讀檔或正在使用中: " + st.Dst);
                        }
                        return "False 【加歌頁面】重複歌曲原有檔案不存在或為同檔案,已自動忽略移除原有檔案: " + st.Dst;
                    },
                    (sc, st) => Assert.Contains(st.Src.Substring(st.Root.Length) + " = SRC", st.Tree().Select(t => t.Replace(" [ro]", ""))));

                Compare("queryremove" + backup, s,
                    st => { var log = new List<string>(); bool err = LegacyPlacement.RemoveSongFile(st.Src, backup, st.RemoveSong, log); return err + " " + string.Join(";", log); },
                    st => { bool ok = LibraryPlacement.Remove(st.Src, backup); return (!ok) + " " + (ok ? "" : "【歌庫查詢】異動檔案時發生錯誤: " + st.Src + " (唯讀或使用中)"); },
                    (sc, st) => { /* removing a file is unaffected by where the destination is */ });
            }
        }

        // ---------- 歌庫結構重建 / 歌庫結構更新: Overwrite ----------

        [Fact]
        public void Overwrite_matches_Rebuild_and_UpdateStructure()
        {
            foreach (Scenario s in Scenarios())
            {
                foreach (string mode in new[] { "1", "2" })
                {
                    Compare("rebuild" + mode, s,
                        st => { var log = new List<string>(); bool err = LegacyPlacement.Rebuild(st.Src, st.Dst, mode, log); return err + " " + string.Join(";", log); },
                        st =>
                        {
                            var o = LibraryPlacement.Overwrite(st.Src, st.Dst, mode == "2");
                            switch (o)
                            {
                                case PlacementOutcome.SourceMissing: return "True 【歌庫結構重建】此首歌曲檔案不存在,已自動忽略重建檔案: " + st.Src;
                                case PlacementOutcome.FileLocked: return "True 【歌庫結構重建】檔案處理發生錯誤: " + st.Src + " (唯讀或使用中)";
                                default: return "False ";
                            }
                        },
                        (sc, st) => KeptAtDestination(sc, st));
                }

                Compare("update", s,
                    st => { var log = new List<string>(); bool moved = LegacyPlacement.UpdateStructure(st.Src, st.Dst, log); return moved + " " + string.Join(";", log); },
                    st =>
                    {
                        if (st.Src == st.Dst) return "True ";
                        var o = LibraryPlacement.Overwrite(st.Src, st.Dst, false);
                        switch (o)
                        {
                            case PlacementOutcome.SourceMissing: return "False 【歌庫結構更新】原始檔案不存在: " + st.Src + " (已忽略更新)";
                            case PlacementOutcome.FileLocked: return "False 【歌庫結構更新】檔案移動失敗: " + st.Src + " (已忽略更新)";
                            default: return "True ";
                        }
                    },
                    (sc, st) => KeptAtDestination(sc, st));
            }
        }

        // ---------- the bugs the module fixes, stated directly ----------

        [Fact]
        public void Legacy_structure_update_deleted_a_file_whose_new_name_differs_only_in_case()
        {
            Scenario s = new Scenario { Dest = Dest.SameFileOtherCase };
            using (Setup legacy = new Setup(Path.Combine(root, "bug-legacy"), s))
            {
                LegacyPlacement.UpdateStructure(legacy.Src, legacy.Dst, new List<string>());
                Assert.Empty(legacy.Tree());
            }
            using (Setup module = new Setup(Path.Combine(root, "bug-module"), s))
            {
                Assert.Equal(PlacementOutcome.Placed, LibraryPlacement.Overwrite(module.Src, module.Dst, false));
                Assert.Equal(new[] { module.Dst.Substring(module.Root.Length) + " = SRC" }, module.Tree());
            }
        }

        [Fact]
        public void Legacy_add_in_move_mode_deleted_a_file_that_was_already_at_its_destination()
        {
            Scenario s = new Scenario { Dest = Dest.SameFileExact };
            using (Setup legacy = new Setup(Path.Combine(root, "add-legacy"), s))
            {
                bool err = LegacyPlacement.AddSongImport(legacy.Src, legacy.Dst, "1", false, legacy.RemoveSong, new List<string>());
                Assert.True(err);
                Assert.Empty(legacy.Tree());
            }
            using (Setup module = new Setup(Path.Combine(root, "add-module"), s))
            {
                LibraryPlacement.RemoveSongFolder = module.RemoveSong;
                Assert.True(LibraryPlacement.Import(module.Src, module.Dst, SongAddMode.Move, false));
                Assert.Equal(new[] { module.Dst.Substring(module.Root.Length) + " = SRC" }, module.Tree());
            }
        }
    }
}

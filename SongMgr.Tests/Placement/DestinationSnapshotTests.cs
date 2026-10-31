using System;
using System.Collections.Generic;
using System.Text;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Placement
{
    /// <summary>
    /// Golden master for the 歌庫結構 destination rule: every combination of
    /// 加歌模式 × 資料夾結構 × 檔案結構 × separators × 合唱歸類, for each placement purpose.
    /// </summary>
    public class DestinationSnapshotTests : IDisposable
    {
        readonly Dictionary<string, object> saved = TestEnv.Capture();

        public void Dispose() => TestEnv.Restore(saved);

        internal static readonly string DestFolder = @"D:\KTV\Songs";

        // Id, Lang, SingerType, Singer, SongName, Track, SongType, FileName, Path
        internal static readonly string[][] Songs =
        {
            new[] { "000123", "國語", "0", "周杰倫", "晴天", "1", "", "周杰倫_晴天_VR.mpg", @"D:\KTV\Songs\國語\男歌星\周杰倫\" },
            new[] { "000124", "國語", "1", "S.H.E.", "Super Star", "2", "演唱會", "she.MKV", @"D:\KTV\Songs\國語\女歌星\S.H.E\" },
            new[] { "100001", "台語", "3", "江蕙+陳雷", "愛人醉落去", "0", "", "a.avi", @"E:\Import\台語\合唱\" },
            new[] { "200002", "粵語", "3", "張學友&陳慧嫻", "相思風雨中", "5", "MTV", "b.dat", @"D:\KTV\Songs\粵語\合唱\" },
            new[] { "300003", "日語", "10", "未知歌手", "テスト", "3", "", "c.mp4", @"D:\KTV\Songs\日語\新進\未知歌手\" },
            new[] { "400004", "英語", "7", "Maroon 5", "Sugar", "4", "", "d.mkv", @"F:\Other\Mix\" },
            new[] { "500005", "客語", "-1", "某人", "山歌", "9", "", "e.mpg", @"D:\KTV\Songs\客語\男\某人\" },
            new[] { "600006", "兒歌", "2", "小小", "兩隻老虎", "1", "", "f.vob", @"D:\KTV\Songs\" },
        };

        internal static readonly string[] Purposes = { "Add", "Relocate", "Rebuild", "UpdateStructure" };

        internal static string Describe(Func<string[], string, string> destinationFor)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string mode in new[] { "1", "3" })
            foreach (string folder in new[] { "1", "2", "3" })
            foreach (string file in new[] { "1", "2", "3" })
            foreach (string infoSep in new[] { "1", "2" })
            foreach (string chorusSep in new[] { "1", "2" })
            foreach (string merge in new[] { "True", "False" })
            {
                Global.SongMgrDestFolder = DestFolder;
                Global.SongMgrSongAddMode = mode;
                Global.SongMgrFolderStructure = folder;
                Global.SongMgrFileStructure = file;
                Global.SongMgrSongInfoSeparate = infoSep;
                Global.SongMgrChorusSeparate = chorusSep;
                Global.SongMgrChorusMerge = merge;

                foreach (string[] s in Songs)
                foreach (string purpose in Purposes)
                {
                    string result;
                    try
                    {
                        result = destinationFor(s, purpose);
                    }
                    catch (Exception ex)
                    {
                        result = "EXCEPTION " + ex.GetType().Name;
                    }
                    sb.Append(mode).Append(folder).Append(file).Append(infoSep).Append(chorusSep).Append(merge[0])
                      .Append(' ').Append(s[0]).Append(' ').Append(purpose).Append(" => ").Append(result).Append('\n');
                }
            }
            return sb.ToString();
        }

        [Fact]
        public void Destination_matches_pre_refactor_snapshot()
        {
            Snapshot.Match("placement-destination", Describe(ModuleDestination));
        }

        static string ModuleDestination(string[] s, string purpose)
        {
            SongRecord song = new SongRecord
            {
                Id = s[0], Lang = s[1], SingerType = s[2], Singer = s[3], SongName = s[4],
                Track = s[5], SongType = s[6], FileName = s[7], Path = s[8],
            };
            PlacementPurpose p = (PlacementPurpose)Enum.Parse(typeof(PlacementPurpose), purpose);
            return LibraryPlacement.Destination(song, p, @"G:\Rebuild");
        }
    }
}

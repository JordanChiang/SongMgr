using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace CrazyKTV_SongMgr
{
    class SongAnalysis
    {
        private static object LockThis = new object();

        #region --- SongAnalysis 建立資料表 ---

        public static DataTable SongAnalysisDT;
        public static SongAnalysisSettings Settings;
        public static SingerLookup Singers;
        public static List<string> SongDataLowCaseList;
        public static List<string> SongDataFilePathList;
        public static List<string> SongAnalysisSongList;
        public static bool SongAnalysisCompleted;
        public static bool SongAnalysisError;

        public static void CreateSongDataTable()
        {
            SongAnalysisCompleted = false;
            SongAnalysisError = false;

            SongAnalysisDT = new DataTable();
            SongAnalysisDT.Columns.Add("Song_AddStatus", typeof(string));
            SongAnalysisDT.Columns.Add("Song_Id", typeof(string));
            SongAnalysisDT.Columns.Add("Song_Lang", typeof(string));
            SongAnalysisDT.Columns.Add("Song_SingerType", typeof(int));
            SongAnalysisDT.Columns.Add("Song_Singer", typeof(string));
            SongAnalysisDT.Columns.Add("Song_SongName", typeof(string));
            SongAnalysisDT.Columns.Add("Song_Track", typeof(int));
            SongAnalysisDT.Columns.Add("Song_SongType", typeof(string));
            SongAnalysisDT.Columns.Add("Song_Volume", typeof(int));
            SongAnalysisDT.Columns.Add("Song_WordCount", typeof(int));
            SongAnalysisDT.Columns.Add("Song_PlayCount", typeof(int));
            SongAnalysisDT.Columns.Add("Song_MB", typeof(float));
            SongAnalysisDT.Columns.Add("Song_CreatDate", typeof(DateTime));
            SongAnalysisDT.Columns.Add("Song_FileName", typeof(string));
            SongAnalysisDT.Columns.Add("Song_Path", typeof(string));
            SongAnalysisDT.Columns.Add("Song_Spell", typeof(string));
            SongAnalysisDT.Columns.Add("Song_SpellNum", typeof(string));
            SongAnalysisDT.Columns.Add("Song_SongStroke", typeof(int));
            SongAnalysisDT.Columns.Add("Song_PenStyle", typeof(string));
            SongAnalysisDT.Columns.Add("Song_PlayState", typeof(int));
            SongAnalysisDT.Columns.Add("Song_SrcPath", typeof(string));
            SongAnalysisDT.Columns.Add("Song_SortIndex", typeof(string));
            SongAnalysisDT.Columns.Add("Song_ReplayGain", typeof(string));

            SongAnalysisSongList = new List<string>();
            Settings = SongAnalysisSettings.FromGlobal();

            var tasks = new List<Task>();

            tasks.Add(Task.Factory.StartNew(() =>
            {
                Singers = SingerLookup.LoadFromDatabase();
            }));

            SongDataLowCaseList = new List<string>();
            SongDataFilePathList = new List<string>();

            tasks.Add(Task.Factory.StartNew(() =>
            {
                string SongDataSqlStr = "select Song_Id, Song_Lang, Song_Singer, Song_SongName, Song_SongType, Song_FileName, Song_Path from ktv_Song";
                using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, SongDataSqlStr, ""))
                {
                    foreach (DataRow row in dt.AsEnumerable())
                    {
                        SongDataLowCaseList.Add(row["Song_Lang"].ToString() + "|" + row["Song_Singer"].ToString().ToLower() + "|" + row["Song_SongName"].ToString().ToLower() + "|" + row["Song_SongType"].ToString().ToLower());
                        SongDataFilePathList.Add(Path.GetFullPath(Path.Combine(row["Song_Path"].ToString(), row["Song_FileName"].ToString())).ToLower());
                    }
                }
            }));

            bool TaskFinished = false;
            Task.Factory.ContinueWhenAll(tasks.ToArray(), EndTask =>
            {
                TaskFinished = true;
            });
            SpinWait.SpinUntil(() => TaskFinished == true);
        }

        public static void DisposeSongDataTable()
        {
            Settings = null;
            Singers = null;
            SongDataLowCaseList.Clear();
            SongDataLowCaseList = null;
            SongDataFilePathList.Clear();
            SongDataFilePathList = null;
            SongAnalysisSongList.Clear();
            SongAnalysisSongList = null;
            SongAnalysisDT.Dispose();
            SongAnalysisDT = null;
            SongAnalysisCompleted = true;
            GC.Collect();
        }

        #endregion

        #region --- SongAnalysis 分析歌曲 ---

        public static void SongInfoAnalysis(string file)
        {
            SongAnalysisResult r = SongAnalyzer.Analyze(file, Settings, Singers);

            // 歌庫監視
            if (Global.AddMode == SongAddMode.Monitor && r.Lang == "未知")
            {
                Global.SongLog.Add("【歌庫監視】檔案結構中須有語系類別資訊才能加歌: " + r.SrcPath);
                return;
            }
            CreateDataRow(r);
        }

        /// <summary>Adds the result to the 加歌 table unless it duplicates a song already there or a file already in ktv_Song.</summary>
        public static void CreateDataRow(SongAnalysisResult r)
        {
            string SongLang = r.Lang, SongSinger = r.Singer, SongSongName = r.SongName, SongSongType = r.SongType, SongSrcPath = r.SrcPath;
            lock (LockThis)
            {
                DataRow row = SongAnalysisDT.NewRow();
                row["Song_AddStatus"] = r.AddStatus;
                row["Song_Id"] = r.Id;
                row["Song_Lang"] = r.Lang;
                row["Song_SingerType"] = r.SingerType;
                row["Song_Singer"] = r.Singer;
                row["Song_SongName"] = r.SongName;
                row["Song_Track"] = r.Track;
                row["Song_SongType"] = r.SongType;
                row["Song_Volume"] = r.Volume;
                row["Song_WordCount"] = r.WordCount;
                row["Song_PlayCount"] = r.PlayCount;
                row["Song_MB"] = r.MB;
                row["Song_CreatDate"] = r.CreatDate;
                row["Song_Spell"] = r.Spell;
                row["Song_SpellNum"] = r.SpellNum;
                row["Song_SongStroke"] = r.SongStroke;
                row["Song_PenStyle"] = r.PenStyle;
                row["Song_PlayState"] = r.PlayState;
                row["Song_SrcPath"] = r.SrcPath;
                row["Song_SortIndex"] = r.SortIndex;
                row["Song_ReplayGain"] = r.ReplayGain;

                string SongSrcPathLower = SongSrcPath.ToLower();
                bool IsDupSongPath = SongDataFilePathList.Contains(SongSrcPathLower);

                CommonFunc.DupSongStatus DupSongStatus = CommonFunc.IsDupSong(SongLang, SongSinger, Convert.ToInt32(r.SingerType), SongSongName, SongSongType, SongAnalysisSongList);
                if (!DupSongStatus.IsDupSong && !IsDupSongPath)
                {
                    SongAnalysisDT.Rows.Add(row);
                    SongAnalysisSongList.Add(SongLang + "|" + SongSinger.ToLower() + "|" + SongSongName.ToLower() + "|" + SongSongType.ToLower());
                }
                else
                {
                    Global.TotalList[0]++;
                    if (!IsDupSongPath)
                    {
                        Global.SongLog.Add("【歌曲分析】判斷為重複歌曲,已忽略: " + SongSrcPath);
                    }
                }
            }
        }

        #endregion

    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 歌曲分析 settings, captured once per run so every file in the run is analysed the same way.
    /// </summary>
    class SongAnalysisSettings
    {
        /// <summary>歌曲辨識模式 1 智能辨識 … 6 歌手\歌名.</summary>
        public string IdentificationMode { get; set; } = "1";
        public string ChorusSeparatorPattern { get; set; } = "[&+、]";
        public bool ConvertToTraditional { get; set; }
        /// <summary>特殊歌手及歌曲名稱, "|"-separated.</summary>
        public string SpecialStr { get; set; } = "";
        /// <summary>歌曲類別 list, ","-separated.</summary>
        public string SongTypes { get; set; } = "";
        public string DefaultSongLang { get; set; } = "11";
        public string DefaultSingerType { get; set; } = "8";
        /// <summary>"6" means detect the track with ffprobe.</summary>
        public string DefaultSongTrack { get; set; } = "1";
        public string DefaultSongType { get; set; } = "1";
        public string DefaultSongVolume { get; set; } = "100";
        public bool DetectVolume { get; set; }
        public string ReplayGainBaseVolume { get; set; } = "50";
        public bool TitleCaseEnglish { get; set; }
        /// <summary>語系 / 歌手類別 / 聲道 keywords, flattened from the "a,b,c" entries.</summary>
        public List<string> LangKeyWords { get; set; } = new List<string>();
        public List<string> SingerTypeKeyWords { get; set; } = new List<string>();
        public List<string> TrackKeyWords { get; set; } = new List<string>();

        public static SongAnalysisSettings FromGlobal()
        {
            return new SongAnalysisSettings
            {
                IdentificationMode = Global.SongAddSongIdentificationMode,
                ChorusSeparatorPattern = Global.RegexChorusSeparate,
                ConvertToTraditional = Global.SongAddEnableConvToTC == "True",
                SpecialStr = Global.SongAddSpecialStr,
                SongTypes = Global.SongMgrSongType,
                DefaultSongLang = Global.SongAddDefaultSongLang,
                DefaultSingerType = Global.SongAddDefaultSingerType,
                DefaultSongTrack = Global.SongAddDefaultSongTrack,
                DefaultSongType = Global.SongAddDefaultSongType,
                DefaultSongVolume = Global.SongAddDefaultSongVolume,
                DetectVolume = Global.SongAddEnableVolumeDetect == "True",
                ReplayGainBaseVolume = Global.SongMaintenanceReplayGainVolume,
                TitleCaseEnglish = Global.SongAddEngSongNameFormat == "True",
                LangKeyWords = Flatten(Global.CrazyktvSongLangKeyWordList),
                SingerTypeKeyWords = Flatten(Global.CrazyktvSingerTypeKeyWordList),
                TrackKeyWords = Flatten(Global.CrazyktvSongTrackKeyWordList),
            };
        }

        private static List<string> Flatten(List<string> entries)
        {
            List<string> list = new List<string>();
            foreach (string str in entries) list.AddRange(str.Split(','));
            return list;
        }
    }

    /// <summary>
    /// The singers 歌曲分析 recognises in file and folder names: ktv_AllSinger, then any ktv_Singer
    /// entry not already there. Read-only once built.
    /// </summary>
    class SingerLookup
    {
        public List<string> Names { get; }
        public List<string> LowCaseNames { get; }
        public List<string> Types { get; }

        public SingerLookup(List<string> names, List<string> types)
        {
            Names = names;
            Types = types;
            LowCaseNames = names.ConvertAll(str => str.ToLower());
        }

        public static SingerLookup LoadFromDatabase()
        {
            List<string> names = new List<string>();
            List<string> types = new List<string>();

            using (DataTable dt = Db.GetDataTable(Global.CrazyktvSongMgrDatabaseFile, "select Singer_Name, Singer_Type from ktv_AllSinger", ""))
            {
                foreach (DataRow row in dt.AsEnumerable())
                {
                    names.Add(row["Singer_Name"].ToString());
                    types.Add(row["Singer_Type"].ToString());
                }
            }
            List<string> lowCase = names.ConvertAll(str => str.ToLower());

            using (DataTable dt = Db.GetDataTable(Global.CrazyktvDatabaseFile, "select Singer_Name, Singer_Type from ktv_Singer", ""))
            {
                foreach (DataRow row in dt.AsEnumerable())
                {
                    if (lowCase.IndexOf(row["Singer_Name"].ToString().ToLower()) < 0)
                    {
                        names.Add(row["Singer_Name"].ToString());
                        types.Add(row["Singer_Type"].ToString());
                    }
                }
            }
            return new SingerLookup(names, types);
        }
    }

    /// <summary>What 歌曲分析 worked out for one file.</summary>
    class SongAnalysisResult
    {
        public string AddStatus, Id, Lang, SingerType, Singer, SongName, Track, SongType, Volume, WordCount, PlayCount;
        public float MB;
        public DateTime CreatDate;
        public string Spell, SpellNum, SongStroke, PenStyle, PlayState, SrcPath, SortIndex, ReplayGain;
    }

    /// <summary>
    /// 歌曲分析: reads 語系, 歌手, 歌手類別, 歌名, 歌曲類別 and 聲道 out of a song file's name and folders.
    /// Holds no state between calls, so files can be analysed in parallel. The reference tables
    /// (語系 / 歌手類別 / 聲道 keywords, 錢櫃 catalog, 注音) are still read from Global.
    /// </summary>
    static class SongAnalyzer
    {
        public static SongAnalysisResult Analyze(string file, SongAnalysisSettings settings, SingerLookup singerLookup)
        {
            List<string> SongLangKeyWordList = settings.LangKeyWords;
            List<string> SingerTypeKeyWordList = settings.SingerTypeKeyWords;
            List<string> SongTrackKeyWordList = settings.TrackKeyWords;

            string SongAddStatus = "";
            string SongID = "";
            string SongLang = "";
            string SongSingerType = "";
            string SongSinger = "";
            string SongSongName = "";
            string SongTrack = "";
            string SongSongType = "";
            string SongVolume = "";
            string SongWordCount = "";
            string SongPlayCount = "";
            float SongMB = 0;
            DateTime SongCreatDate = new DateTime();
            string SongSpell = "";
            string SongSpellNum = "";
            string SongSongStroke = "";
            string SongPenStyle = "";
            string SongPlayState = "";
            string SongSrcPath = Path.GetFullPath(file);
            string SongSortIndex = "";
            string SongReplayGain = "";

            string SongLangStr = "";
            string SongSingerTypeStr = "";
            string SongTrackStr = "";
            string SongSongTypeStr = "";

            List<string> list = new List<string>();

            bool PreFind = true;
            if (settings.IdentificationMode != "1") PreFind = false;

            Regex ChorusRegex = new Regex(settings.ChorusSeparatorPattern);

            // 從檔案名稱查找歌曲資訊
            string FileStr = Path.GetFileNameWithoutExtension(file);
            if (settings.ConvertToTraditional) FileStr = CommonFunc.ConvToTraditionalChinese(FileStr);

            // 特殊歌手及歌曲名稱處理
            List<string> SongAddSpecialStrList = new List<string>(Regex.Split(settings.SpecialStr.ToLower(), @"\|", RegexOptions.IgnoreCase));
            foreach (string str in SongAddSpecialStrList)
            {
                FileStr = Regex.Replace(FileStr, Regex.Escape(str), "%%" + str + "%%", RegexOptions.IgnoreCase);
            }

            // 去除檔案開頭的編號
            FileStr = Regex.Replace(FileStr, @"^(\d{2}|\d{1})\.", "", RegexOptions.IgnoreCase);

            // 歌手-歌名(語系-類別-聲道) 名稱處理
            MatchCollection BracketMatches = Regex.Matches(FileStr, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", RegexOptions.IgnoreCase);
            if (BracketMatches.Count > 0)
            {
                foreach (Match m in BracketMatches)
                {
                    if (m.Value.ContainsAny("-", "_"))
                    {
                        FileStr = FileStr.Replace(m.Value, "%%" + m.Value + "%%");
                    }
                }
            }

            List<string> SongSongTypeList = new List<string>(settings.SongTypes.Split(','));
            List<string> SongSongTypeLowCaseList = SongSongTypeList.ConvertAll(str => str.ToLower());
            List<string> FileNameRemoveList = new List<string>();

            string SeparateStr = "[-_](?=(?:[^%]*%%[^%]*%%)*(?![^%]*%%))";
            list = new List<string>(Regex.Split(FileStr, SeparateStr, RegexOptions.None));

            foreach (string str in list)
            {
                if (SongSingerType != "" && SongSinger != "" && SongLang != "" && SongTrack != "" && SongSongType != "") break;
                string splitstr = Regex.Replace(str, @"%%|^\s*|\s*$", "");

                // 查看檔案名稱中有無歌手
                if (SongSinger == "" && PreFind)
                {
                    string strippedStr = Regex.Replace(splitstr, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                    if (singerLookup.LowCaseNames.Contains(splitstr.ToLower()))
                    {
                        int index = singerLookup.LowCaseNames.IndexOf(splitstr.ToLower());
                        SongSinger = singerLookup.Names[index];
                        if (SongSingerType == "") SongSingerType = singerLookup.Types[index];
                    }
                    else if (strippedStr != "" && singerLookup.LowCaseNames.Contains(strippedStr))
                    {
                        int index = singerLookup.LowCaseNames.IndexOf(strippedStr);
                        SongSinger = singerLookup.Names[index];
                        if (SongSingerType == "") SongSingerType = singerLookup.Types[index];
                    }
                    else
                    {
                        if (ChorusRegex.IsMatch(splitstr.ToLower()))
                        {
                            string[] singers = Regex.Split(splitstr.ToLower(), settings.ChorusSeparatorPattern, RegexOptions.None);
                            foreach (string singer in singers)
                            {
                                string SingerName = Regex.Replace(singer, @"^\s*|\s*$", ""); //去除頭尾空白
                                string strippedSinger = Regex.Replace(SingerName, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                if (singerLookup.LowCaseNames.Contains(SingerName.ToLower()) || (strippedSinger != "" && singerLookup.LowCaseNames.Contains(strippedSinger)))
                                {
                                    SongSinger = splitstr;
                                    if (SongSingerType == "") SongSingerType = "3";
                                    break;
                                }
                            }
                        }
                    }
                }

                // 查看檔案名稱中有無歌手類別
                if (SingerTypeKeyWordList.IndexOf(splitstr) >= 0)
                {
                    SongSingerType = GetSongInfo("SongSingerType", splitstr);
                    SongSingerTypeStr = splitstr;
                }

                // 括號字串處理
                MatchCollection matches = Regex.Matches(splitstr, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", RegexOptions.IgnoreCase);
                if (matches.Count > 0)
                {
                    foreach (Match match in matches) // 比對括號內字串
                    {
                        string MatchStr = Regex.Replace(match.Value, @"^[\{\(\[｛（［【]|[】］）｝\]\)\}]$", "", RegexOptions.IgnoreCase);
                        if (MatchStr.ContainsAny("_", "-")) // 括號字串有分隔符號
                        {
                            List<string> BracketStrlist = new List<string>(Regex.Split(MatchStr, SeparateStr, RegexOptions.None));
                            string FileNameRemoveStr = "";
                            string BracketSeparateStr = (MatchStr.Contains("_")) ? "_" : "-";
                            
                            foreach (string BracketStr in BracketStrlist)
                            {
                                if (SongLangKeyWordList.IndexOf(BracketStr) >= 0)
                                {
                                    SongLang = GetSongInfo("SongLang", BracketStr);
                                    FileNameRemoveStr += ((BracketStrlist.IndexOf(BracketStr) == 0) ? BracketStr : BracketSeparateStr + BracketStr);
                                }
                                else if (SongTrackKeyWordList.IndexOf(BracketStr.ToLower()) >= 0 || BracketStr.ToLower() == "l" || BracketStr.ToLower() == "r")
                                {
                                    SongTrack = GetSongInfo("SongTrack", BracketStr.ToLower());
                                    FileNameRemoveStr += ((BracketStrlist.IndexOf(BracketStr) == 0) ? BracketStr : BracketSeparateStr + BracketStr);
                                }
                                else if (SongSongTypeLowCaseList.IndexOf(BracketStr.ToLower()) >= 0)
                                {
                                    SongSongType = SongSongTypeList[SongSongTypeLowCaseList.IndexOf(BracketStr.ToLower())];
                                    FileNameRemoveStr += ((BracketStrlist.IndexOf(BracketStr) == 0) ? BracketStr : BracketSeparateStr + BracketStr);
                                }
                                else if (BracketStr.ContainsAny("合唱", "對唱"))
                                {
                                    SongSingerType = "3";
                                }
                            }
                            if (FileNameRemoveStr != "") FileNameRemoveList.Add(FileNameRemoveStr);
                        }
                        else // 括號字串無分隔符號
                        {
                            if (SongLangKeyWordList.IndexOf(MatchStr) >= 0)
                            {
                                SongLang = GetSongInfo("SongLang", MatchStr);
                                FileNameRemoveList.Add(MatchStr);
                            }
                            else if (SongTrackKeyWordList.IndexOf(MatchStr.ToLower()) >= 0 || MatchStr.ToLower() == "l" || MatchStr.ToLower() == "r")
                            {
                                SongTrack = GetSongInfo("SongTrack", MatchStr.ToLower());
                                FileNameRemoveList.Add(MatchStr);
                            }
                            else if (SongSongTypeLowCaseList.IndexOf(MatchStr.ToLower()) >= 0)
                            {
                                SongSongType = SongSongTypeList[SongSongTypeLowCaseList.IndexOf(MatchStr.ToLower())];
                                FileNameRemoveList.Add(MatchStr);
                            }
                            else if (MatchStr.ContainsAny("合唱", "對唱"))
                            {
                                SongSingerType = "3";
                            }
                        }
                    }
                }

                // 查看檔案名稱中有無語系類別
                if (SongLang == "")
                {
                    if (SongLangKeyWordList.IndexOf(splitstr) >= 0)
                    {
                        SongLang = GetSongInfo("SongLang", splitstr);
                        SongLangStr = splitstr;
                    }
                }

                // 查看檔案名稱中有無歌曲聲道
                if (SongTrack == "")
                {
                    if (SongTrackKeyWordList.IndexOf(splitstr.ToLower()) >= 0)
                    {
                        SongTrack = GetSongInfo("SongTrack", splitstr.ToLower());
                        SongTrackStr = splitstr;
                    }
                }

                // 查看檔案名稱中有無歌曲類別
                if (SongSongType == "")
                {
                    if (SongSongTypeLowCaseList.IndexOf(splitstr.ToLower()) >= 0)
                    {
                        SongSongType = SongSongTypeList[SongSongTypeLowCaseList.IndexOf(splitstr.ToLower())];
                        SongSongTypeStr = splitstr;
                    }
                }
            }
            SongSongTypeList.Clear();
            SongSongTypeList = null;
            SongSongTypeLowCaseList.Clear();
            SongSongTypeLowCaseList = null;
            list.Clear();
            list = null;

            string DirStr = Path.GetDirectoryName(file);

            if (SongSingerType == "" || SongSinger == "" || SongLang == "")
            {
                list = new List<string>(Regex.Split(DirStr, @"\\", RegexOptions.None));

                foreach (string str in list)
                {
                    if (SongSingerType != "" && SongSinger != "" && SongLang != "") break;
                    string splitstr = Regex.Replace(str, @"%%|^\s*|\s*$", "");

                    // 查看資料夾名稱中有無歌手
                    if (SongSinger == "" && PreFind)
                    {
                        if (singerLookup.LowCaseNames.IndexOf(splitstr.ToLower()) >= 0 && (settings.IdentificationMode == "1" || settings.IdentificationMode == "6"))
                        {
                            SongSinger = singerLookup.Names[singerLookup.LowCaseNames.IndexOf(splitstr.ToLower())];
                            if (SongSingerType == "") SongSingerType = singerLookup.Types[singerLookup.LowCaseNames.IndexOf(splitstr.ToLower())];
                        }
                        else
                        {
                            if (singerLookup.LowCaseNames.IndexOf(splitstr.ToLower()) >= 0)
                            {
                                SongSinger = singerLookup.Names[singerLookup.LowCaseNames.IndexOf(splitstr.ToLower())];
                                if (SongSingerType == "") SongSingerType = singerLookup.Types[singerLookup.LowCaseNames.IndexOf(splitstr.ToLower())];
                            }
                            else
                            {
                                if (ChorusRegex.IsMatch(splitstr.ToLower()))
                                {
                                    string[] singers = Regex.Split(splitstr.ToLower(), settings.ChorusSeparatorPattern, RegexOptions.None);
                                    foreach (string singer in singers)
                                    {
                                        string SingerName = Regex.Replace(singer, @"^\s*|\s*$", ""); //去除頭尾空白
                                        if (singerLookup.LowCaseNames.IndexOf(SingerName.ToLower()) >= 0)
                                        {
                                            SongSinger = splitstr;
                                            if (SongSingerType == "") SongSingerType = "3";
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 查看資料夾名稱中有無歌手類別
                    if (SingerTypeKeyWordList.IndexOf(splitstr) >= 0)
                    {
                        if (SongSingerType == "")
                        {
                            SongSingerType = GetSongInfo("SongSingerType", splitstr);
                        }
                    }


                    // 查看資料夾名稱中有無語系類別
                    if (SongLangKeyWordList.IndexOf(splitstr) >= 0)
                    {
                        if (SongLang == "")
                        {
                            SongLang = GetSongInfo("SongLang", splitstr);
                        }
                    }
                }
                list.Clear();
                list = null;
            }

            // 套用預設歌曲聲道
            if (SongTrack == "")
            {
                if (settings.DefaultSongTrack == "6")
                {
                    var task = Task.Factory.StartNew(() => MediaProbe.DetectTrack(SongSrcPath));
                    SongTrack = task.Result;
                }
                else
                {
                    SongTrack = settings.DefaultSongTrack;
                }
            }

            // 套用預設歌曲類別
            if (SongSongType == "")
            {
                string str = "";
                if (settings.SongTypes != "") { str = "無類別," + settings.SongTypes; } else { str = "無類別"; }
                list = new List<string>(str.Split(','));
                if (list[int.Parse(settings.DefaultSongType) - 1] == "無類別")
                {
                    SongSongType = "";
                }
                else
                {
                    SongSongType = list[int.Parse(settings.DefaultSongType) - 1];
                }
                list.Clear();
            }

            // 套用預設歌曲音量
            if (SongVolume == "")
            {
                if (settings.DetectVolume)
                {
                    int basevolume = Convert.ToInt32(settings.ReplayGainBaseVolume);

                    MediaProbe.SongVolumeValue result = MediaProbe.MeasureGain(file);
                    SongVolume = MediaProbe.SongVolume(basevolume, result.GainDB, VolumeRange.Library);
                    SongReplayGain = result.GainDB.ToString();
                    NativeMethods.SystemSleepManagement.ResetSleepTimer(false);
                }
                else
                {
                    SongVolume = settings.DefaultSongVolume;
                }
            }

            int FileStrCount = 0;
            // 判斷歌手及歌曲名稱
            if (SongSongName == "")
            {
                string str = FileStr;
                if (PreFind)
                {
                    list = new List<string>() { SongLangStr, SongSingerTypeStr, SongTrackStr, SongSongTypeStr };
                    foreach (string s in list)
                    {
                        str = Regex.Replace(str, "^" + s + "[_-]|[_-]" + s + "[_-]|[_-]" + s + "$", "_", RegexOptions.IgnoreCase);
                        str = Regex.Replace(str, "^[_-]|[_-]$", "", RegexOptions.IgnoreCase);
                    }
                    list.Clear();
                }

                // 去除括號內要移除的字串
                if (FileNameRemoveList.Count > 0)
                {
                    foreach (string RemoveStr in FileNameRemoveList)
                    {
                        str = Regex.Replace(str, @"([\{\(\[｛（［【](.+)?)" + Regex.Escape(RemoveStr) + @"([】］）｝\]\)\}])", "$1$3", RegexOptions.IgnoreCase);
                        str = Regex.Replace(str, @"[\{\(\[｛（［【]" + @"[】］）｝\]\)\}]", "", RegexOptions.IgnoreCase);
                    }
                }
                FileNameRemoveList.Clear();

                list = new List<string>(Regex.Split(str, SeparateStr, RegexOptions.None));
                FileStrCount = list.Count;

                switch (settings.IdentificationMode)
                {
                    case "1": // 智能辨識模式
                        switch (list.Count)
                        {
                            case 1:
                                {
                                    SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                    string strippedSegment0 = Regex.Replace(SongSongName, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                    if (SongSinger == "" || SongSinger.ToLower() == SongSongName.ToLower() || (strippedSegment0 != "" && SongSinger.ToLower() == strippedSegment0))
                                    {
                                        SongSinger = Regex.Replace(Path.GetFileName(DirStr), @"%%|^\s*|\s*$", "");
                                        if (ChorusRegex.IsMatch(SongSinger))
                                        {
                                            if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                            {
                                                SongSingerType = "3";
                                            }
                                        }
                                    }
                                }
                                break;
                            case 2:
                                {
                                    if (SongSinger == "") // 未找到歌手
                                    {
                                        SongSinger = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                        if (ChorusRegex.IsMatch(SongSinger))
                                        {
                                            if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                            {
                                                SongSingerType = "3";
                                            }
                                        }
                                        SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                    }
                                    else // 已找到歌手
                                    {
                                        string segment0 = Regex.Replace(list[0], @"%%|^\s*|\s*$", "").ToLower();
                                        string strippedSegment0 = Regex.Replace(segment0, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim();
                                        if (SongSinger.ToLower() == segment0 || (strippedSegment0 != "" && SongSinger.ToLower() == strippedSegment0))
                                        {
                                            SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                        }
                                        else
                                        {
                                            SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                        }
                                    }
                                }
                                break;
                            default:
                                {
                                    if (SongSinger == "") // 未找到歌手
                                    {
                                        if (CommonFunc.IsSongId(list[0]))
                                        {
                                            SongID = list[0];
                                            SongSinger = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            if (ChorusRegex.IsMatch(SongSinger))
                                            {
                                                if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                                {
                                                    SongSingerType = "3";
                                                }
                                            }
                                            SongSongName = Regex.Replace(list[2], @"%%|^\s*|\s*$", "");
                                        }
                                        else
                                        {
                                            SongSinger = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                            if (ChorusRegex.IsMatch(SongSinger))
                                            {
                                                if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                                {
                                                    SongSingerType = "3";
                                                }
                                            }
                                            SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                        }
                                    }
                                    else // 已找到歌手
                                    {
                                        if (CommonFunc.IsSongId(list[0]))
                                        {
                                            SongID = list[0];
                                            string segment1 = Regex.Replace(list[1], @"%%|^\s*|\s*$", "").ToLower();
                                            string strippedSegment1 = Regex.Replace(segment1, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim();
                                            if (SongSinger.ToLower() == segment1 || (strippedSegment1 != "" && SongSinger.ToLower() == strippedSegment1))
                                            {
                                                SongSongName = Regex.Replace(list[2], @"%%|^\s*|\s*$", "");
                                            }
                                            else
                                            {
                                                SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            }
                                        }
                                        else // List[0] 不是歌曲編號
                                        {
                                            string segment0 = Regex.Replace(list[0], @"%%|^\s*|\s*$", "").ToLower();
                                            string strippedSegment0 = Regex.Replace(segment0, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim();
                                            if (SongSinger.ToLower() == segment0 || (strippedSegment0 != "" && SongSinger.ToLower() == strippedSegment0))
                                            {
                                                SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            }
                                            else
                                            {
                                                SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                            }
                                        }
                                    }
                                }
                                break;
                        }
                        break;
                    case "2": // 歌手_歌名
                    case "3": // 歌名_歌手
                        switch (list.Count)
                        {
                            case 1:
                                {
                                    SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                    string strippedSegment0_C2 = Regex.Replace(SongSongName, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                    if (SongSinger == "" || SongSinger.ToLower() == SongSongName.ToLower() || (strippedSegment0_C2 != "" && SongSinger.ToLower() == strippedSegment0_C2))
                                    {
                                        SongSinger = Regex.Replace(Path.GetFileName(DirStr), @"%%|^\s*|\s*$", "");
                                        string strippedSinger_C2 = Regex.Replace(SongSinger, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                        if (singerLookup.LowCaseNames.Contains(SongSinger.ToLower()))
                                        {
                                            int index = singerLookup.LowCaseNames.IndexOf(SongSinger.ToLower());
                                            SongSinger = singerLookup.Names[index];
                                            SongSingerType = singerLookup.Types[index];
                                        }
                                        else if (strippedSinger_C2 != "" && singerLookup.LowCaseNames.Contains(strippedSinger_C2))
                                        {
                                            int index = singerLookup.LowCaseNames.IndexOf(strippedSinger_C2);
                                            SongSinger = singerLookup.Names[index];
                                            SongSingerType = singerLookup.Types[index];
                                        }
                                        else
                                        {
                                            if (ChorusRegex.IsMatch(SongSinger))
                                            {
                                                if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                                {
                                                    SongSingerType = "3";
                                                }
                                            }
                                        }
                                    }
                                }
                                break;
                            default:
                                {
                                    switch (settings.IdentificationMode)
                                    {
                                        case "2":
                                            SongSinger = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                            SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            break;
                                        case "3":
                                            SongSinger = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                            break;
                                    }

                                    string strippedSinger_C2_Def = Regex.Replace(SongSinger, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                    if (singerLookup.LowCaseNames.Contains(SongSinger.ToLower()))
                                    {
                                        int index = singerLookup.LowCaseNames.IndexOf(SongSinger.ToLower());
                                        SongSinger = singerLookup.Names[index];
                                        SongSingerType = singerLookup.Types[index];
                                    }
                                    else if (strippedSinger_C2_Def != "" && singerLookup.LowCaseNames.Contains(strippedSinger_C2_Def))
                                    {
                                        int index = singerLookup.LowCaseNames.IndexOf(strippedSinger_C2_Def);
                                        SongSinger = singerLookup.Names[index];
                                        SongSingerType = singerLookup.Types[index];
                                    }
                                    else
                                    {
                                        if (ChorusRegex.IsMatch(SongSinger))
                                        {
                                            if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                            {
                                                SongSingerType = "3";
                                            }
                                        }
                                    }
                                }
                                break;
                        }
                        break;
                    case "4": // 歌曲編號_歌手_歌名
                    case "5": // 歌曲編號_歌名_歌手
                        switch (list.Count)
                        {
                            case 1:
                                {
                                    SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                    string strippedSegment0_C4 = Regex.Replace(SongSongName, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                    if (SongSinger == "" || SongSinger.ToLower() == SongSongName.ToLower() || (strippedSegment0_C4 != "" && SongSinger.ToLower() == strippedSegment0_C4))
                                    {
                                        SongSinger = Regex.Replace(Path.GetFileName(DirStr), @"%%|^\s*|\s*$", "");
                                        string strippedSinger_C4 = Regex.Replace(SongSinger, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                        if (singerLookup.LowCaseNames.Contains(SongSinger.ToLower()))
                                        {
                                            int index = singerLookup.LowCaseNames.IndexOf(SongSinger.ToLower());
                                            SongSinger = singerLookup.Names[index];
                                            SongSingerType = singerLookup.Types[index];
                                        }
                                        else if (strippedSinger_C4 != "" && singerLookup.LowCaseNames.Contains(strippedSinger_C4))
                                        {
                                            int index = singerLookup.LowCaseNames.IndexOf(strippedSinger_C4);
                                            SongSinger = singerLookup.Names[index];
                                            SongSingerType = singerLookup.Types[index];
                                        }
                                        else
                                        {
                                            if (ChorusRegex.IsMatch(SongSinger))
                                            {
                                                if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                                {
                                                    SongSingerType = "3";
                                                }
                                            }
                                        }
                                    }
                                }
                                break;
                            case 2:
                                {
                                    switch (settings.IdentificationMode)
                                    {
                                        case "4":
                                            SongSinger = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                            SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            break;
                                        case "5":
                                            SongSinger = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                                            break;
                                    }

                                    string strippedSinger_C4_2 = Regex.Replace(SongSinger, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                    if (singerLookup.LowCaseNames.Contains(SongSinger.ToLower()))
                                    {
                                        int index = singerLookup.LowCaseNames.IndexOf(SongSinger.ToLower());
                                        SongSinger = singerLookup.Names[index];
                                        SongSingerType = singerLookup.Types[index];
                                    }
                                    else if (strippedSinger_C4_2 != "" && singerLookup.LowCaseNames.Contains(strippedSinger_C4_2))
                                    {
                                        int index = singerLookup.LowCaseNames.IndexOf(strippedSinger_C4_2);
                                        SongSinger = singerLookup.Names[index];
                                        SongSingerType = singerLookup.Types[index];
                                    }
                                    else
                                    {
                                        if (ChorusRegex.IsMatch(SongSinger))
                                        {
                                            if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                            {
                                                SongSingerType = "3";
                                            }
                                        }
                                    }
                                }
                                break;
                            default:
                                {
                                    if (CommonFunc.IsSongId(list[0])) SongID = list[0];
                                    switch (settings.IdentificationMode)
                                    {
                                        case "4":
                                            SongSinger = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            SongSongName = Regex.Replace(list[2], @"%%|^\s*|\s*$", "");
                                            break;
                                        case "5":
                                            SongSinger = Regex.Replace(list[2], @"%%|^\s*|\s*$", "");
                                            SongSongName = Regex.Replace(list[1], @"%%|^\s*|\s*$", "");
                                            break;
                                    }

                                    string strippedSinger_C4_Def = Regex.Replace(SongSinger, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                    if (singerLookup.LowCaseNames.Contains(SongSinger.ToLower()))
                                    {
                                        int index = singerLookup.LowCaseNames.IndexOf(SongSinger.ToLower());
                                        SongSinger = singerLookup.Names[index];
                                        SongSingerType = singerLookup.Types[index];
                                    }
                                    else if (strippedSinger_C4_Def != "" && singerLookup.LowCaseNames.Contains(strippedSinger_C4_Def))
                                    {
                                        int index = singerLookup.LowCaseNames.IndexOf(strippedSinger_C4_Def);
                                        SongSinger = singerLookup.Names[index];
                                        SongSingerType = singerLookup.Types[index];
                                    }
                                    else
                                    {
                                        if (ChorusRegex.IsMatch(SongSinger))
                                        {
                                            if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                            {
                                                SongSingerType = "3";
                                            }
                                        }
                                    }
                                }
                                break;
                        }
                        break;
                    case "6": // 歌手\歌名
                        {
                            SongSongName = Regex.Replace(list[0], @"%%|^\s*|\s*$", "");
                            string strippedSegment0_C6 = Regex.Replace(SongSongName, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                            if (SongSinger == "" || SongSinger.ToLower() == SongSongName.ToLower() || (strippedSegment0_C6 != "" && SongSinger.ToLower() == strippedSegment0_C6))
                            {
                                SongSinger = Regex.Replace(Path.GetFileName(DirStr), @"%%|^\s*|\s*$", "");
                                string strippedSinger_C6 = Regex.Replace(SongSinger, @"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", "").Trim().ToLower();
                                if (singerLookup.LowCaseNames.Contains(SongSinger.ToLower()))
                                {
                                    int index = singerLookup.LowCaseNames.IndexOf(SongSinger.ToLower());
                                    SongSinger = singerLookup.Names[index];
                                    SongSingerType = singerLookup.Types[index];
                                }
                                else if (strippedSinger_C6 != "" && singerLookup.LowCaseNames.Contains(strippedSinger_C6))
                                {
                                    int index = singerLookup.LowCaseNames.IndexOf(strippedSinger_C6);
                                    SongSinger = singerLookup.Names[index];
                                    SongSingerType = singerLookup.Types[index];
                                }
                                else
                                {
                                    if (ChorusRegex.IsMatch(SongSinger))
                                    {
                                        if (SongAddSpecialStrList.IndexOf(SongSinger.ToLower()) < 0)
                                        {
                                            SongSingerType = "3";
                                        }
                                    }
                                }
                            }
                        }
                        break;
                }
                list.Clear();
            }

            // 問題歌曲處理
            if (Global.SongAnalysisExceptionSingerList.IndexOf(SongSinger) >= 0)
            {
                int i = Global.SongAnalysisExceptionSingerList.IndexOf(SongSinger);

                switch (SongSinger)
                {
                    case "台灣女孩":
                        if (SongSongName == "羅百吉&寶貝")
                        {
                            if (FileStr.ContainsAll(Global.SongAnalysisExceptionSongList[i].Split('|')))
                            {
                                SongLang = "國語";
                                SongSinger = "羅百吉&寶貝";
                                SongSongName = "台灣女孩";
                                SongSingerType = "3";
                            }
                        }
                        break;
                    case "梁靜茹":
                    case "張懸":
                        if (SongLang == "兒歌")
                        {
                            if (FileStr.ContainsAll(Global.SongAnalysisExceptionSongList[i].Split('|')))
                            {
                                SongLang = "國語";
                                SongSongName = "兒歌";
                            }
                            else if (FileStrCount == 1 && SongSongName == "兒歌")
                            {
                                SongLang = "國語";
                            }
                        }
                        break;
                    case "莫文蔚":
                        if (SongLang == "台語")
                        {
                            if (FileStr.ContainsAll(Global.SongAnalysisExceptionSongList[i].Split('|')))
                            {
                                SongLang = "國語";
                                SongSongName = "台";
                            }
                            else if (FileStrCount == 1 && SongSongName == "台")
                            {
                                SongLang = "國語";
                            }
                        }
                        break;
                }
            }

            // 套用預設歌手類別
            if (SongSingerType == "")
            {
                string str = CommonFunc.GetSingerTypeStr(int.Parse(settings.DefaultSingerType) - 1, 3, "null");
                SongSingerType = CommonFunc.GetSingerTypeStr(0, 1, str);
            }

            // 統一資料庫中合唱歌手分隔符號
            if (SongSingerType == "3")
            {
                if (ChorusRegex.IsMatch(SongSinger)) SongSinger = Regex.Replace(SongSinger, settings.ChorusSeparatorPattern, "&", RegexOptions.IgnoreCase);
            }

            // 套用預設語系類別
            if (SongLang == "")
            {
                SongLang = CommonFunc.GetSongLangStr(int.Parse(settings.DefaultSongLang) - 1, 0, "null");

                int SongDataIndex = CommonFunc.MatchCashboxSong("CashboxLang", null, SongLang, SongSinger, SongSongName, Global.CashboxSongDataLowCaseList, Global.CashboxSongDataFuzzyList);

                if (SongDataIndex >= 0)
                {
                    string CashboxSongDataLang = Global.CashboxSongDataLangList[SongDataIndex];
                    if (Global.CrazyktvSongLangList.IndexOf(CashboxSongDataLang) >= 0)
                    {
                        SongLang = Global.CrazyktvSongLangList[Global.CrazyktvSongLangList.IndexOf(CashboxSongDataLang)];
                    }
                }
            }

            // 計算歌曲字數
            List<string> SongWordCountList = new List<string>();
            SongWordCountList = CommonFunc.GetSongWordCount(SongSongName);
            SongWordCount = SongWordCountList[0];

            if (settings.TitleCaseEnglish)
            {
                if (SongWordCountList[1] == "True")
                {
                    TextInfo ti = new CultureInfo("en-US",false).TextInfo;
                    SongSongName = ti.ToTitleCase(SongSongName.ToLower());
                }
            }

            // 點播次數初始值
            SongPlayCount = "0";

            // 計算歌曲大小
            FileInfo f = new FileInfo(file);
            SongMB = float.Parse(((f.Length / 1024f) / 1024f).ToString("F2"));

            // 取得加歌日期
            SongCreatDate = DateTime.Now;

            // 取得歌曲拼音
            List<string> SongSpellList = new List<string>();
            SongSpellList = CommonFunc.GetSongNameSpell(SongSongName);
            SongSpell = SongSpellList[0];
            SongSpellNum = SongSpellList[1];
            if (SongSpellList[2] == "") SongSpellList[2] = "0";
            SongSongStroke = SongSpellList[2];
            SongPenStyle = SongSpellList[3];

            // 播放狀態初始值
            SongPlayState = "0";

            // 排序索引
            if (SongLang == "未知") SongSortIndex = "1";
            if (SongSortIndex == "") { if (SongSingerType == "10") SongSortIndex = "2"; }
            if (SongSortIndex == "") { if (SongID != "") SongSortIndex = "3"; }
            if (SongSortIndex == "") SongSortIndex = "4";

            // 加歌狀態
            if (SongLang == "未知") SongAddStatus = "語系類別必須有值才能加歌!";
            if (SongAddStatus == "") { if (SongSingerType == "10") SongAddStatus = "此歌手尚未設定歌手資料!";}

            return new SongAnalysisResult
            {
                AddStatus = SongAddStatus, Id = SongID, Lang = SongLang, SingerType = SongSingerType, Singer = SongSinger,
                SongName = SongSongName, Track = SongTrack, SongType = SongSongType, Volume = SongVolume, WordCount = SongWordCount,
                PlayCount = SongPlayCount, MB = SongMB, CreatDate = SongCreatDate, Spell = SongSpell, SpellNum = SongSpellNum,
                SongStroke = SongSongStroke, PenStyle = SongPenStyle, PlayState = SongPlayState, SrcPath = SongSrcPath,
                SortIndex = SongSortIndex, ReplayGain = SongReplayGain,
            };
        }

        private static string GetSongInfo(string SongInfoType, string SongInfoValue)
        {
            string infovalue = "";
            List<string> list = new List<string>();

            switch (SongInfoType)
            {
                case "SongLang":
                    foreach (string str in Global.CrazyktvSongLangKeyWordList)
                    {
                        list = new List<string>(str.Split(','));
                        foreach (string liststr in list)
                        {
                            if (SongInfoValue == liststr)
                            {
                                infovalue = Global.CrazyktvSongLangList[Global.CrazyktvSongLangKeyWordList.IndexOf(str)];
                                break;
                            }
                        }
                        if (infovalue != "") break;
                    }
                    break;
                case "SongSingerType":
                    foreach (string str in Global.CrazyktvSingerTypeKeyWordList)
                    {
                        list = new List<string>(str.Split(','));
                        foreach (string liststr in list)
                        {
                            if (SongInfoValue == liststr)
                            {
                                infovalue = Global.CrazyktvSingerTypeKeyWordList.IndexOf(str).ToString();
                                break;
                            }
                        }
                        if (infovalue != "") break;
                    }
                    break;
                case "SongTrack":
                    foreach (string str in Global.CrazyktvSongTrackKeyWordList)
                    {
                        list = new List<string>(str.Split(','));
                        foreach (string liststr in list)
                        {
                            if (SongInfoValue == liststr)
                            {
                                infovalue = Global.CrazyktvSongTrackKeyWordList.IndexOf(str).ToString();
                                break;
                            }
                        }
                        if (infovalue != "") break;
                    }
                    break;
            }
            list.Clear();
            return infovalue;
        }
    }
}

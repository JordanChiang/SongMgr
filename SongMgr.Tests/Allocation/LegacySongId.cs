using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SongMgr.Tests.Allocation
{
    /// <summary>
    /// Frozen copies of the 歌曲編號 code as it was before the 歌曲編號配置 module existed (master @ 32d61f5).
    /// Only edits: Global's MaxIDList / UnusedSongIdList / RemainingSongIdCountList, SongMgrLangCode,
    /// SongMgrMaxDigitCode and CrazyktvSongLangList are fields of this object, and the "select Song_Id from ktv_Song"
    /// results come from <see cref="LibraryIds"/>. SongDataIdList is SongAddSong's list of the 歌庫's Song_Id plus
    /// the songs added so far; the tests append to it after each successful add, as SongAddSong.cs:521 did.
    /// </summary>
    internal class LegacySongId
    {
        public List<int> MaxIDList = new List<int>();
        public List<List<string>> UnusedSongIdList = new List<List<string>>();
        public List<int> RemainingSongIdCountList = new List<int>();
        public string SongMgrLangCode;
        public string SongMgrMaxDigitCode;
        public List<string> CrazyktvSongLangList;
        public List<string> LibraryIds;
        public List<string> SongDataIdList;
        readonly object LockThis = new object();

        public LegacySongId(List<string> langs, string langCode, int digits, List<string> libraryIds)
        {
            CrazyktvSongLangList = langs;
            SongMgrLangCode = langCode;
            SongMgrMaxDigitCode = digits == 5 ? "1" : "2";
            LibraryIds = libraryIds;
            SongDataIdList = new List<string>(libraryIds);
        }

        int DigitCode => SongMgrMaxDigitCode == "1" ? 5 : 6;

        // The preload every caller ran first: CommonFunc.GetMaxSongId + GetUnusedSongId (CommonFunc.cs:2190-2290).
        public void Load()
        {
            GetMaxSongId(DigitCode);
            GetUnusedSongId(DigitCode);
        }

        // CommonFunc.cs:2144-2181
        public void GetRemainingSongIdCount(int DigitCode)
        {
            List<string> StartIdlist = new List<string>(Regex.Split(SongMgrLangCode, ",", RegexOptions.None));
            RemainingSongIdCountList = new List<int>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            StartIdlist.Add((DigitCode == 5) ? "100000" : "1000000");

            List<string> dt = LibraryIds.OrderBy(s => s, StringComparer.Ordinal).ToList();
            int RemainingSongId;

            foreach (string StartIdStr in StartIdlist)
            {
                if (StartIdStr == ((DigitCode == 5) ? "100000" : "1000000")) break;

                int sid = Convert.ToInt32(StartIdStr);
                int eid = Convert.ToInt32(StartIdlist[StartIdlist.IndexOf(StartIdStr) + 1]) - 1;
                int tid = eid - sid + 1;

                var query = from row in dt
                            where Convert.ToInt32(row) >= sid &&
                                  Convert.ToInt32(row) <= eid &&
                                  row.Length == DigitCode
                            select row;

                if (query.Count() > 0)
                {
                    RemainingSongId = tid - query.Count();
                }
                else
                {
                    RemainingSongId = tid;
                }
                RemainingSongIdCountList[StartIdlist.IndexOf(StartIdStr)] = RemainingSongId;
            }
        }

        // CommonFunc.cs:2190-2234
        public void GetMaxSongId(int DigitCode)
        {
            MaxIDList = new List<int>();
            List<string> StartIdlist = new List<string>(Regex.Split(SongMgrLangCode, ",", RegexOptions.None));

            for (int i = 0; i < 10; i++)
            {
                MaxIDList.Add(Convert.ToInt32(StartIdlist[i]) - 1);
            }

            List<string> dt = LibraryIds.OrderByDescending(s => s, StringComparer.Ordinal).ToList();
            if (dt.Count > 0)
            {
                Parallel.For(0, CrazyktvSongLangList.Count, i =>
                {
                    int sid = Convert.ToInt32(StartIdlist[i]);
                    int eid;
                    if (i < 9)
                    {
                        eid = Convert.ToInt32(StartIdlist[i + 1]) - 1;
                    }
                    else
                    {
                        eid = (SongMgrMaxDigitCode == "1") ? 99999 : 999999;
                    }

                    foreach (string songIdStr in dt)
                    {
                        if (songIdStr.Length == DigitCode)
                        {
                            int id = Convert.ToInt32(songIdStr);
                            if (id >= sid && id <= eid)
                            {
                                MaxIDList[i] = id;
                                break;
                            }
                        }
                    }
                });
            }
        }

        // CommonFunc.cs:2236-2290
        public void GetUnusedSongId(int DigitCode)
        {
            string MaxDigitCode = (SongMgrMaxDigitCode == "1") ? "D5" : "D6";
            List<string> StartIdlist = new List<string>(Regex.Split(SongMgrLangCode, ",", RegexOptions.None));

            UnusedSongIdList = new List<List<string>>();
            for (int i = 0; i < CrazyktvSongLangList.Count; i++)
            {
                UnusedSongIdList.Add(new List<string>());
            }

            if (LibraryIds.Count > 0)
            {
                HashSet<string> allUsedIds = new HashSet<string>(LibraryIds);

                Parallel.For(0, CrazyktvSongLangList.Count, i =>
                {
                    int sid = Convert.ToInt32(StartIdlist[i]);
                    int eid;
                    if (i < 9)
                    {
                        eid = Convert.ToInt32(StartIdlist[i + 1]) - 1;
                    }
                    else
                    {
                        eid = (SongMgrMaxDigitCode == "1") ? 99999 : 999999;
                    }

                    int currentMax = sid;
                    foreach (string usedId in allUsedIds)
                    {
                        if (usedId.Length == DigitCode)
                        {
                            int idVal = Convert.ToInt32(usedId);
                            if (idVal >= sid && idVal <= eid && idVal > currentMax)
                            {
                                currentMax = idVal;
                            }
                        }
                    }

                    for (int id = sid; id <= currentMax; id++)
                    {
                        string idStr = id.ToString(MaxDigitCode);
                        if (!allUsedIds.Contains(idStr))
                        {
                            UnusedSongIdList[i].Add(idStr);
                        }
                    }
                });
            }
        }

        // SongQuery.cs:3275-3297 (SongMaintenance.cs:4042 and Cashbox.cs:1901 are this with UpdateIdList = true)
        public string GetNextSongId(string SongLang, bool UpdateIdList)
        {
            string NewSongID = "";

            if (UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Count > 0)
            {
                NewSongID = UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)][0];
                if (UpdateIdList) UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Remove(NewSongID);
            }

            if (NewSongID == "")
            {
                string MaxDigitCode = (SongMgrMaxDigitCode == "1") ? "D5" : "D6";
                MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)]++;
                NewSongID = MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)].ToString(MaxDigitCode);
                if (!UpdateIdList) MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)]--;
            }
            return NewSongID;
        }

        // SongAddSong.cs:311-453: the 自訂編號 check, then claiming it or 自動分配編號. Returns the Song_Id used and
        // whether the 自訂編號 was accepted. (SongAddUseCustomSongID is on and SongId is not empty.)
        public (string SongId, bool UseCustomSongID) AddSongId(string SongId, string SongLang)
        {
            bool UseCustomSongID = false;
            {
                List<string> StartIdlist = new List<string>(Regex.Split(SongMgrLangCode, ",", RegexOptions.None));

                switch (SongMgrMaxDigitCode)
                {
                    case "1":
                        if (SongId.Length == 5 && SongDataIdList.IndexOf(SongId) < 0)
                        {
                            if (CrazyktvSongLangList.IndexOf(SongLang) < 9)
                            {
                                if (Convert.ToInt32(SongId) >= Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang)]) &&
                                Convert.ToInt32(SongId) < Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang) + 1]))
                                {
                                    UseCustomSongID = true;
                                }
                            }
                            else
                            {
                                if (Convert.ToInt32(SongId) >= Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang)]) &&
                                Convert.ToInt32(SongId) < 100000)
                                {
                                    UseCustomSongID = true;
                                }
                            }
                        }
                        break;
                    case "2":
                        if (SongId.Length == 6 && SongDataIdList.IndexOf(SongId) < 0)
                        {
                            if (Convert.ToInt32(SongId) >= 0 && Convert.ToInt32(SongId) <= 99999)
                            {
                                UseCustomSongID = true;
                            }
                            else
                            {
                                if (CrazyktvSongLangList.IndexOf(SongLang) < 9)
                                {
                                    if (Convert.ToInt32(SongId) >= Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang)]) &&
                                    Convert.ToInt32(SongId) < Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang) + 1]))
                                    {
                                        UseCustomSongID = true;
                                    }
                                }
                                else
                                {
                                    if (Convert.ToInt32(SongId) >= Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang)]) &&
                                    Convert.ToInt32(SongId) < 1000000)
                                    {
                                        UseCustomSongID = true;
                                    }
                                }
                            }
                        }
                        break;
                }
            }

            if (UseCustomSongID)
            {
                if (SongId.Length == 6 && Convert.ToInt32(SongId) >= 0 && Convert.ToInt32(SongId) <= 99999)
                {
                    List<string> StartIdlist = new List<string>(Regex.Split(SongMgrLangCode, ",", RegexOptions.None));
                    bool RefreshID = false;
                    foreach (string startid in StartIdlist)
                    {
                        if (Convert.ToInt32(startid) <= 99999)
                        {
                            RefreshID = true;
                        }
                    }
                    if (RefreshID)
                    {
                        Parallel.ForEach(CrazyktvSongLangList, (SongLangStr, loopState) =>
                        {
                            if (UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLangStr)].Count > 0)
                            {
                                if (UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLangStr)].IndexOf(SongId) >= 0)
                                {
                                    lock (LockThis)
                                    {
                                        UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLangStr)].Remove(SongId);
                                        if (Convert.ToInt32(SongId) > MaxIDList[CrazyktvSongLangList.IndexOf(SongLangStr)])
                                        {
                                            MaxIDList[CrazyktvSongLangList.IndexOf(SongLangStr)] = Convert.ToInt32(SongId);
                                            GetUnUsedSongId(SongId, SongLangStr);
                                        }
                                        loopState.Break();
                                    }
                                }
                            }
                        });
                    }
                }
                else
                {
                    if (UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Count > 0)
                    {
                        if (UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].IndexOf(SongId) >= 0)
                        {
                            lock (LockThis)
                            {
                                UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Remove(SongId);
                            }
                        }
                    }

                    if (Convert.ToInt32(SongId) > MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)])
                    {
                        lock (LockThis)
                        {
                            MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)] = Convert.ToInt32(SongId);
                            GetUnUsedSongId(SongId, SongLang);
                        }
                    }
                }
            }
            else
            {
                SongId = "";

                if (UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Count > 0)
                {
                    SongId = UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)][0];
                    UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Remove(SongId);
                }

                if (SongId == "")
                {
                    string MaxDigitCode = (SongMgrMaxDigitCode == "1") ? "D5" : "D6";
                    MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)]++;
                    SongId = MaxIDList[CrazyktvSongLangList.IndexOf(SongLang)].ToString(MaxDigitCode);
                }
            }

            return (SongId, UseCustomSongID);
        }

        // SongAddSong.cs:657-672
        void GetUnUsedSongId(string MaxSongID, string SongLang)
        {
            string MaxDigitCode = (SongMgrMaxDigitCode == "1") ? "D5" : "D6";

            List<string> StartIdlist = new List<string>(Regex.Split(SongMgrLangCode, ",", RegexOptions.None));
            int iMin = Convert.ToInt32(StartIdlist[CrazyktvSongLangList.IndexOf(SongLang)]);
            int iMax = Convert.ToInt32(MaxSongID);

            UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Clear();

            for (int i = iMin; i < iMax; i++)
            {
                if (SongDataIdList.IndexOf(i.ToString(MaxDigitCode)) < 0)
                {
                    UnusedSongIdList[CrazyktvSongLangList.IndexOf(SongLang)].Add(i.ToString(MaxDigitCode));
                }
            }
        }

        // SongMaintenance.cs:657-701 (轉換為 6 位數): every song renumbered from its 語系's start, in Song_Id order.
        public static List<string> ConvertTo6(List<string> langs, string langCode, List<(string Id, string Lang)> songs)
        {
            string MaxDigitCode = "D6";
            List<string> StrIdlist = new List<string>(Regex.Split(langCode, ",", RegexOptions.None));
            List<int> Idlist = StrIdlist.Select(s => Convert.ToInt32(s)).ToList();
            var result = new List<string>();
            foreach (var row in songs)
            {
                int LangIndex = langs.IndexOf(row.Lang);
                string NewSongId = Idlist[LangIndex].ToString(MaxDigitCode);
                Idlist[LangIndex]++;
                result.Add(NewSongId);
            }
            return result;
        }
    }
}

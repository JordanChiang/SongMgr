using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 歌曲記錄: one song's fields by name, as they travel between the song grids, the edit panel,
    /// the 加歌 pipeline and ktv_Song. Values are kept as the text the grid or the pipeline produced,
    /// so nothing is reformatted on the way to the database.
    /// </summary>
    class SongRecord
    {
        public string Id { get; set; } = "";
        public string Lang { get; set; } = "";
        public string SingerType { get; set; } = "";
        public string Singer { get; set; } = "";
        public string SongName { get; set; } = "";
        public string Track { get; set; } = "";
        public string SongType { get; set; } = "";
        public string Volume { get; set; } = "";
        public string WordCount { get; set; } = "";
        public string PlayCount { get; set; } = "";
        public string MB { get; set; } = "";
        public string CreatDate { get; set; } = "";
        public string FileName { get; set; } = "";
        public string Path { get; set; } = "";
        public string Spell { get; set; } = "";
        public string SpellNum { get; set; } = "";
        public string SongStroke { get; set; } = "";
        public string PenStyle { get; set; } = "";
        public string PlayState { get; set; } = "";
        public string ReplayGain { get; set; } = "";

        /// <summary>Where the file is now. In the 加歌 grid this is the file being added.</summary>
        public string SrcPath { get; set; } = "";

        /// <summary>The Song_Id this record replaces when it is written back (歌庫查詢 edits).</summary>
        public string OldId { get; set; } = "";

        /// <summary>加歌 grid only: the status text shown in the 加歌狀態 column.</summary>
        public string AddStatus { get; set; } = "";

        /// <summary>加歌 pipeline only: "1" when the singer must be added to ktv_Singer / ktv_AllSinger.</summary>
        public string AddSinger { get; set; } = "";
        public string AddAllSinger { get; set; } = "";

        public int SingerTypeValue => Convert.ToInt32(SingerType);
        public int TrackValue => Convert.ToInt32(Track);

        public SongRecord Clone() => (SongRecord)MemberwiseClone();

        /// <summary>A row of the 歌庫查詢 grid; the file is at Path\FileName.</summary>
        public static SongRecord FromQueryGridRow(DataGridViewRow row)
        {
            SongRecord song = ReadCommonCells(row);
            song.SrcPath = System.IO.Path.Combine(song.Path, song.FileName);
            return song;
        }

        /// <summary>A row of the 加歌 grid; the file being added is in the Song_SrcPath column.</summary>
        public static SongRecord FromAddGridRow(DataGridViewRow row)
        {
            SongRecord song = ReadCommonCells(row);
            song.SrcPath = row.Cells["Song_SrcPath"].Value.ToString();
            song.AddStatus = row.Cells["Song_AddStatus"].Value.ToString();
            return song;
        }

        /// <summary>The fields the player shows, read straight from a grid row.</summary>
        public static SongRecord ForPlayer(DataGridViewRow row)
        {
            return new SongRecord
            {
                Id = row.Cells["Song_Id"].Value.ToString(),
                Lang = row.Cells["Song_Lang"].Value.ToString(),
                Singer = row.Cells["Song_Singer"].Value.ToString(),
                SongName = row.Cells["Song_SongName"].Value.ToString(),
                Track = row.Cells["Song_Track"].Value.ToString(),
                Volume = row.Cells["Song_Volume"].Value.ToString(),
                ReplayGain = row.Cells["Song_ReplayGain"].Value.ToString(),
            };
        }

        private static SongRecord ReadCommonCells(DataGridViewRow row)
        {
            SongRecord song = new SongRecord
            {
                Id = row.Cells["Song_Id"].Value.ToString(),
                Lang = row.Cells["Song_Lang"].Value.ToString(),
                SingerType = Convert.ToInt32(row.Cells["Song_SingerType"].Value).ToString(),
                Singer = row.Cells["Song_Singer"].Value.ToString(),
                SongName = row.Cells["Song_SongName"].Value.ToString(),
                Track = Convert.ToInt32(row.Cells["Song_Track"].Value).ToString(),
                SongType = row.Cells["Song_SongType"].Value.ToString(),
                Volume = row.Cells["Song_Volume"].Value.ToString(),
                WordCount = row.Cells["Song_WordCount"].Value.ToString(),
                PlayCount = row.Cells["Song_PlayCount"].Value.ToString(),
                MB = row.Cells["Song_MB"].Value.ToString(),
                CreatDate = row.Cells["Song_CreatDate"].Value.ToString(),
                FileName = row.Cells["Song_FileName"].Value.ToString(),
                Path = row.Cells["Song_Path"].Value.ToString(),
                Spell = row.Cells["Song_Spell"].Value.ToString(),
                SpellNum = row.Cells["Song_SpellNum"].Value.ToString(),
                SongStroke = row.Cells["Song_SongStroke"].Value.ToString(),
                PenStyle = row.Cells["Song_PenStyle"].Value.ToString(),
                PlayState = row.Cells["Song_PlayState"].Value.ToString(),
                ReplayGain = row.Cells["Song_ReplayGain"].Value.ToString(),
            };
            return song;
        }

        /// <summary>The edited fields written back over a grid row (歌庫查詢 / 加歌 edit panel).</summary>
        public void WriteTo(DataRow row, bool addGrid)
        {
            if (addGrid)
            {
                row["Song_AddStatus"] = AddStatus;
                row["Song_Lang"] = Lang;
                row["Song_SingerType"] = SingerType;
                row["Song_Singer"] = Singer;
                row["Song_SongName"] = SongName;
                row["Song_Track"] = Track;
                row["Song_SongType"] = SongType;
                row["Song_Volume"] = Volume;
                row["Song_WordCount"] = WordCount;
                row["Song_PlayCount"] = PlayCount;
                row["Song_CreatDate"] = CreatDate;
                row["Song_Spell"] = Spell;
                row["Song_SpellNum"] = SpellNum;
                row["Song_SongStroke"] = SongStroke;
                row["Song_PenStyle"] = PenStyle;
                row["Song_SrcPath"] = SrcPath;
                return;
            }

            row["Song_Id"] = Id;
            row["Song_Lang"] = Lang;
            row["Song_SingerType"] = SingerType;
            row["Song_Singer"] = Singer;
            row["Song_SongName"] = SongName;
            row["Song_Track"] = Track;
            row["Song_SongType"] = SongType;
            row["Song_Volume"] = Volume;
            row["Song_WordCount"] = WordCount;
            row["Song_PlayCount"] = PlayCount;
            row["Song_MB"] = MB;
            row["Song_CreatDate"] = CreatDate;
            row["Song_FileName"] = FileName;
            row["Song_Path"] = Path;
            row["Song_Spell"] = Spell;
            row["Song_SpellNum"] = SpellNum;
            row["Song_SongStroke"] = SongStroke;
            row["Song_PenStyle"] = PenStyle;
            row["Song_PlayState"] = PlayState;
        }

        // The record as one line for the 操作記錄 / 加歌結果 lists, in the field order those lists
        // have always shown: Id … PlayState, then what the operation adds.

        /// <summary>歌庫查詢 edits: … PlayState|OldId.</summary>
        public string ToLogString() => CoreFields() + "|" + OldId;

        /// <summary>加歌: … PlayState|AddSinger|AddAllSinger|ReplayGain.</summary>
        public string ToAddLogString() => CoreFields() + "|" + AddSinger + "|" + AddAllSinger + "|" + ReplayGain;

        /// <summary>重複歌曲 replace: … PlayState|ReplayGain.</summary>
        public string ToReplaceLogString() => CoreFields() + "|" + ReplayGain;

        /// <summary>歌庫轉換: … PlayState.</summary>
        public string ToConvertLogString() => CoreFields();

        private string CoreFields()
        {
            return Id + "|" + Lang + "|" + SingerType + "|" + Singer + "|" + SongName + "|" + Track + "|" + SongType + "|" + Volume + "|" + WordCount + "|" + PlayCount + "|" + MB + "|" + CreatDate + "|" + FileName + "|" + Path + "|" + Spell + "|" + SpellNum + "|" + SongStroke + "|" + PenStyle + "|" + PlayState;
        }
    }
}

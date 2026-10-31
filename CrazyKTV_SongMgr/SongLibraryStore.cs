using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace CrazyKTV_SongMgr
{
    /// <summary>ktv_Song columns, for choosing which fields a write touches.</summary>
    [Flags]
    enum SongColumns
    {
        None = 0,
        Id = 1 << 0,
        Lang = 1 << 1,
        SingerType = 1 << 2,
        Singer = 1 << 3,
        SongName = 1 << 4,
        Track = 1 << 5,
        SongType = 1 << 6,
        Volume = 1 << 7,
        WordCount = 1 << 8,
        PlayCount = 1 << 9,
        MB = 1 << 10,
        CreatDate = 1 << 11,
        FileName = 1 << 12,
        Path = 1 << 13,
        Spell = 1 << 14,
        SpellNum = 1 << 15,
        SongStroke = 1 << 16,
        PenStyle = 1 << 17,
        PlayState = 1 << 18,
        ReplayGain = 1 << 19,

        /// <summary>Every column except Song_ReplayGain, which only exists when 音量偵測 wrote it.</summary>
        AllButReplayGain = (1 << 19) - 1,
        All = AllButReplayGain | ReplayGain,
    }

    /// <summary>
    /// 歌手資料表: ktv_Singer lives in CrazySong.mdb (the 歌庫's own singers), ktv_AllSinger in
    /// CrazySongMgr.mdb (the 預設 singer list that 歌曲分析 recognises).
    /// </summary>
    sealed class SingerTable
    {
        public string Name { get; }
        public bool IsLibrary => Name == "ktv_Singer";
        public string DatabaseFile => (Name == "ktv_Singer") ? Global.CrazyktvDatabaseFile : Global.CrazyktvSongMgrDatabaseFile;

        private SingerTable(string name) { Name = name; }

        public static readonly SingerTable Library = new SingerTable("ktv_Singer");
        public static readonly SingerTable AllSingers = new SingerTable("ktv_AllSinger");

        /// <summary>The table 歌手管理 works on (歌手管理 → 預設歌手資料表).</summary>
        public static SingerTable Default => (Global.SingerMgrDefaultSingerDataTable == "ktv_Singer") ? Library : AllSingers;
    }

    /// <summary>
    /// 歌庫存取: every write to ktv_Song, ktv_Favorite and the 歌手資料表 goes through here, so the SQL,
    /// the column lists and the connection strings live in one place. One instance holds one connection;
    /// dispose it when the batch is done. Writes throw like OleDbCommand.ExecuteNonQuery does.
    /// </summary>
    sealed class SongLibraryStore : IDisposable
    {
        private static readonly (SongColumns Column, string Name, Func<SongRecord, string> Value)[] SongFields =
        {
            (SongColumns.Id, "Song_Id", s => s.Id),
            (SongColumns.Lang, "Song_Lang", s => s.Lang),
            (SongColumns.SingerType, "Song_SingerType", s => s.SingerType),
            (SongColumns.Singer, "Song_Singer", s => s.Singer),
            (SongColumns.SongName, "Song_SongName", s => s.SongName),
            (SongColumns.Track, "Song_Track", s => s.Track),
            (SongColumns.SongType, "Song_SongType", s => s.SongType),
            (SongColumns.Volume, "Song_Volume", s => s.Volume),
            (SongColumns.WordCount, "Song_WordCount", s => s.WordCount),
            (SongColumns.PlayCount, "Song_PlayCount", s => s.PlayCount),
            (SongColumns.MB, "Song_MB", s => s.MB),
            (SongColumns.CreatDate, "Song_CreatDate", s => s.CreatDate),
            (SongColumns.FileName, "Song_FileName", s => s.FileName),
            (SongColumns.Path, "Song_Path", s => s.Path),
            (SongColumns.Spell, "Song_Spell", s => s.Spell),
            (SongColumns.SpellNum, "Song_SpellNum", s => s.SpellNum),
            (SongColumns.SongStroke, "Song_SongStroke", s => s.SongStroke),
            (SongColumns.PenStyle, "Song_PenStyle", s => s.PenStyle),
            (SongColumns.PlayState, "Song_PlayState", s => s.PlayState),
            (SongColumns.ReplayGain, "Song_ReplayGain", s => s.ReplayGain),
        };

        private readonly OleDbConnection conn;
        private readonly Dictionary<string, OleDbCommand> commands = new Dictionary<string, OleDbCommand>();
        private OleDbTransaction transaction;

        private SongLibraryStore(string databaseFile, bool throwIfUnavailable = false)
        {
            if (throwIfUnavailable)
            {
                conn = new OleDbConnection(ConnectionString(databaseFile));
                conn.Open();
            }
            else
            {
                conn = CommonFunc.OleDbOpenConn(databaseFile, "");
            }
        }

        /// <summary>CrazySong.mdb: ktv_Song, ktv_Favorite, ktv_Singer.</summary>
        public static SongLibraryStore Open() => new SongLibraryStore(Global.CrazyktvDatabaseFile);

        /// <summary>Like <see cref="Open()"/>, but throws when the database cannot be opened instead of failing on the first write.</summary>
        public static SongLibraryStore OpenOrThrow() => new SongLibraryStore(Global.CrazyktvDatabaseFile, true);

        /// <summary>Another CrazyKTV database, e.g. the target of a 歌庫轉換.</summary>
        public static SongLibraryStore OpenFile(string databaseFile) => new SongLibraryStore(databaseFile);

        /// <summary>The database file holding <paramref name="table"/>.</summary>
        public static SongLibraryStore Open(SingerTable table) => new SongLibraryStore(table.DatabaseFile);

        /// <summary>Jet connection string for a database file.</summary>
        public static string ConnectionString(string databaseFile) => "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databaseFile + ";";

        /// <summary>Following writes join one transaction until <see cref="Commit"/> or <see cref="Rollback"/>.</summary>
        public void BeginTransaction()
        {
            lock (commands) { transaction = conn.BeginTransaction(); }
        }

        public void Commit()
        {
            lock (commands) { transaction.Commit(); transaction = null; }
        }

        public void Rollback()
        {
            lock (commands) { transaction.Rollback(); transaction = null; }
        }

        public void Dispose()
        {
            foreach (OleDbCommand cmd in commands.Values) cmd.Dispose();
            conn.Dispose();
        }

        #region --- ktv_Song ---

        public int InsertSong(SongRecord song, SongColumns columns)
        {
            var fields = Fields(columns);
            string sql = "insert into ktv_Song ( " + string.Join(", ", fields.ConvertAll(f => f.Name)) + " ) values ( " + string.Join(", ", fields.ConvertAll(f => "@" + f.Name.Replace("_", ""))) + " )";
            return Execute(sql, cmd => { foreach (var f in fields) cmd.Parameters.AddWithValue("@" + f.Name.Replace("_", ""), f.Value(song)); });
        }

        /// <summary>
        /// Writes the chosen columns of <paramref name="song"/> to the row whose Song_Id is
        /// <see cref="SongRecord.OldId"/>, or <see cref="SongRecord.Id"/> when OldId is empty.
        /// Returns the number of rows changed.
        /// </summary>
        public int UpdateSong(SongRecord song, SongColumns columns)
        {
            var fields = Fields(columns);
            string sql = "update ktv_Song set " + string.Join(", ", fields.ConvertAll(f => f.Name + " = @" + f.Name.Replace("_", ""))) + " where Song_Id = @WhereSongId";
            string whereId = (song.OldId != "") ? song.OldId : song.Id;
            return Execute(sql, cmd =>
            {
                foreach (var f in fields) cmd.Parameters.AddWithValue("@" + f.Name.Replace("_", ""), f.Value(song));
                cmd.Parameters.AddWithValue("@WhereSongId", whereId);
            });
        }

        /// <summary>歌曲編號 change guarded by name and 語系, so a row renumbered meanwhile is left alone.</summary>
        public int RenumberSong(string oldId, string newId, string songName, string lang)
        {
            return Execute("update ktv_Song set Song_Id = @SongId where Song_Id = @OldSongId and Song_SongName = @SongSongName and Song_Lang = @SongLang", cmd =>
            {
                cmd.Parameters.AddWithValue("@SongId", newId);
                cmd.Parameters.AddWithValue("@OldSongId", oldId);
                cmd.Parameters.AddWithValue("@SongSongName", songName);
                cmd.Parameters.AddWithValue("@SongLang", lang);
            });
        }

        /// <summary>Forgets every measured 重播增益.</summary>
        public int ClearAllReplayGain()
        {
            return Execute("update ktv_Song set Song_ReplayGain = null", cmd => { });
        }

        public int DeleteSong(string songId)
        {
            return Execute("delete from ktv_Song where Song_Id = @SongId", cmd => cmd.Parameters.AddWithValue("@SongId", songId));
        }

        #endregion

        #region --- ktv_Favorite ---

        public int RenumberFavorite(string oldId, string newId)
        {
            return Execute("update ktv_Favorite set Song_Id = @SongId where Song_Id = @OldSongId", cmd =>
            {
                cmd.Parameters.AddWithValue("@SongId", newId);
                cmd.Parameters.AddWithValue("@OldSongId", oldId);
            });
        }

        public int DeleteFavorite(string songId)
        {
            return Execute("delete from ktv_Favorite where Song_Id = @SongId", cmd => cmd.Parameters.AddWithValue("@SongId", songId));
        }

        #endregion

        #region --- 歌手資料表 ---

        /// <summary>
        /// Adds a singer with 拼音 / 筆劃 / 筆形 taken from <see cref="CommonFunc.GetSongNameSpell"/>
        /// (whose list order is Spell, SpellNum, Strokes, PenStyle).
        /// </summary>
        public int InsertSinger(SingerTable table, string singerId, string name, string singerType)
        {
            List<string> spell = CommonFunc.GetSongNameSpell(name);
            return InsertSinger(table, singerId, name, singerType, spell[0], spell[2], spell[1], spell[3]);
        }

        /// <summary>Rewrites the singer whose Singer_Id is <paramref name="oldId"/>.</summary>
        public int UpdateSinger(SingerTable table, string oldId, string singerId, string name, string singerType, string spell, string strokes, string spellNum, string penStyle)
        {
            string sql = "update " + table.Name + " set Singer_Id = @SingerId, Singer_Name = @SingerName, Singer_Type = @SingerType, Singer_Spell = @SingerSpell, Singer_Strokes = @SingerStrokes, Singer_SpellNum = @SingerSpellNum, Singer_PenStyle = @SingerPenStyle where Singer_Id = @OldSingerId";
            return Execute(sql, cmd =>
            {
                cmd.Parameters.AddWithValue("@SingerId", singerId);
                cmd.Parameters.AddWithValue("@SingerName", name);
                cmd.Parameters.AddWithValue("@SingerType", singerType);
                cmd.Parameters.AddWithValue("@SingerSpell", spell);
                cmd.Parameters.AddWithValue("@SingerStrokes", strokes);
                cmd.Parameters.AddWithValue("@SingerSpellNum", spellNum);
                cmd.Parameters.AddWithValue("@SingerPenStyle", penStyle);
                cmd.Parameters.AddWithValue("@OldSingerId", oldId);
            });
        }

        /// <summary>Empties the table before a 歌手資料 import.</summary>
        public int DeleteAllSingers(SingerTable table)
        {
            return Execute("delete * from " + table.Name, cmd => { });
        }

        public int DeleteSinger(SingerTable table, string singerId)
        {
            return Execute("delete from " + table.Name + " where Singer_Id = @SingerId", cmd => cmd.Parameters.AddWithValue("@SingerId", singerId));
        }

        public int InsertSinger(SingerTable table, string singerId, string name, string singerType, string spell, string strokes, string spellNum, string penStyle)
        {
            string sql = "insert into " + table.Name + " ( Singer_Id, Singer_Name, Singer_Type, Singer_Spell, Singer_Strokes, Singer_SpellNum, Singer_PenStyle ) values ( @SingerId, @SingerName, @SingerType, @SingerSpell, @SingerStrokes, @SingerSpellNum, @SingerPenStyle )";
            return Execute(sql, cmd =>
            {
                cmd.Parameters.AddWithValue("@SingerId", singerId);
                cmd.Parameters.AddWithValue("@SingerName", name);
                cmd.Parameters.AddWithValue("@SingerType", singerType);
                cmd.Parameters.AddWithValue("@SingerSpell", spell);
                cmd.Parameters.AddWithValue("@SingerStrokes", strokes);
                cmd.Parameters.AddWithValue("@SingerSpellNum", spellNum);
                cmd.Parameters.AddWithValue("@SingerPenStyle", penStyle);
            });
        }

        #endregion

        private static List<(SongColumns Column, string Name, Func<SongRecord, string> Value)> Fields(SongColumns columns)
        {
            var list = new List<(SongColumns, string, Func<SongRecord, string>)>();
            foreach (var f in SongFields)
            {
                if ((columns & f.Column) != 0) list.Add(f);
            }
            if (list.Count == 0) throw new ArgumentException("No columns chosen.", nameof(columns));
            return list;
        }

        /// <summary>One statement at a time: an OleDbConnection is not safe to share between threads.</summary>
        private int Execute(string sql, Action<OleDbCommand> bind)
        {
            lock (commands)
            {
                OleDbCommand cmd;
                if (!commands.TryGetValue(sql, out cmd))
                {
                    cmd = new OleDbCommand(sql, conn);
                    commands[sql] = cmd;
                }
                cmd.Parameters.Clear();
                cmd.Transaction = transaction;
                try
                {
                    bind(cmd);
                    return cmd.ExecuteNonQuery();
                }
                finally
                {
                    cmd.Parameters.Clear();
                }
            }
        }
    }
}

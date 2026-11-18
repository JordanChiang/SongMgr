using System;
using System.Collections.Generic;
using System.Data.Common;

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
    /// 歌手資料表: ktv_Singer lives in the 歌庫 database (the 歌庫's own singers), ktv_AllSinger in
    /// CrazySongMgr, the 參考資料庫 (the 預設 singer list that 歌曲分析 recognises).
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
    /// A name's 拼音 (first 注音 or letter of each word), 手機輸入 digits, 筆劃 and 筆形, as the 歌手資料表 store them
    /// in Singer_Spell, Singer_SpellNum, Singer_Strokes and Singer_PenStyle.
    /// </summary>
    sealed class NameSpell
    {
        public string Spell { get; }
        public string SpellNum { get; }
        public string Strokes { get; }
        public string PenStyle { get; }

        public NameSpell(string spell, string spellNum, string strokes, string penStyle)
        {
            Spell = spell;
            SpellNum = spellNum;
            Strokes = strokes;
            PenStyle = penStyle;
        }

        /// <summary>Looks the name up in the 參考資料庫's 注音 table (<see cref="CommonFunc.GetSongNameSpell"/>).</summary>
        public static NameSpell Of(string name)
        {
            List<string> spell = CommonFunc.GetSongNameSpell(name);
            return new NameSpell(spell[0], spell[1], spell[2], spell[3]);
        }
    }

    /// <summary>
    /// 歌庫存取: every write to ktv_Song, ktv_User, ktv_Favorite and the 歌手資料表 goes through here, so the SQL
    /// and the column lists live in one place; the connection comes from <see cref="Db"/>. One instance holds one connection;
    /// dispose it when the batch is done. Writes throw like DbCommand.ExecuteNonQuery does.
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

        private readonly DbConnection conn;
        private readonly Dictionary<string, DbCommand> commands = new Dictionary<string, DbCommand>();
        private DbTransaction transaction;

        private SongLibraryStore(string databaseFile, bool throwIfUnavailable = false)
        {
            conn = throwIfUnavailable ? Db.Open(databaseFile) : Db.OpenConn(databaseFile);
        }

        /// <summary>CrazySong.mdb: ktv_Song, ktv_Favorite, ktv_Singer.</summary>
        public static SongLibraryStore Open() => new SongLibraryStore(Global.CrazyktvDatabaseFile);

        /// <summary>Like <see cref="Open()"/>, but throws when the database cannot be opened instead of failing on the first write.</summary>
        public static SongLibraryStore OpenOrThrow() => new SongLibraryStore(Global.CrazyktvDatabaseFile, true);

        /// <summary>Another CrazyKTV database, e.g. the target of a 歌庫轉換.</summary>
        public static SongLibraryStore OpenFile(string databaseFile) => new SongLibraryStore(databaseFile);

        /// <summary>The database file holding <paramref name="table"/>.</summary>
        public static SongLibraryStore Open(SingerTable table) => new SongLibraryStore(table.DatabaseFile);

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
            foreach (DbCommand cmd in commands.Values) cmd.Dispose();
            conn.Dispose();
        }

        #region --- ktv_Song ---

        public int InsertSong(SongRecord song, SongColumns columns)
        {
            var fields = Fields(columns);
            string sql = "insert into ktv_Song ( " + string.Join(", ", fields.ConvertAll(f => f.Name)) + " ) values ( " + string.Join(", ", fields.ConvertAll(f => "@" + f.Name.Replace("_", ""))) + " )";
            return Execute(sql, cmd => { foreach (var f in fields) cmd.Parameters.AddWithValue("@" + f.Name.Replace("_", ""), ParamValue(f.Column, f.Value(song))); });
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
                foreach (var f in fields) cmd.Parameters.AddWithValue("@" + f.Name.Replace("_", ""), ParamValue(f.Column, f.Value(song)));
                cmd.Parameters.AddWithValue("@WhereSongId", whereId);
            });
        }

        /// <summary>
        /// 歌曲編號 change guarded by name and 語系, so a row renumbered meanwhile is left alone.
        /// The song's 我的最愛 move to the new number with it.
        /// </summary>
        public int RenumberSong(string oldId, string newId, string songName, string lang)
        {
            int rows = Execute("update ktv_Song set Song_Id = @SongId where Song_Id = @OldSongId and Song_SongName = @SongSongName and Song_Lang = @SongLang", cmd =>
            {
                cmd.Parameters.AddWithValue("@SongId", newId);
                cmd.Parameters.AddWithValue("@OldSongId", oldId);
                cmd.Parameters.AddWithValue("@SongSongName", songName);
                cmd.Parameters.AddWithValue("@SongLang", lang);
            });
            if (rows > 0) RenumberFavorite(oldId, newId);
            return rows;
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

        #region --- 我的最愛 (ktv_User, ktv_Favorite) ---

        /// <summary>Every 我的最愛 that points at <paramref name="oldId"/> points at <paramref name="newId"/> instead.</summary>
        public int RenumberFavorite(string oldId, string newId)
        {
            return Execute("update ktv_Favorite set Song_Id = @SongId where Song_Id = @OldSongId", cmd =>
            {
                cmd.Parameters.AddWithValue("@SongId", newId);
                cmd.Parameters.AddWithValue("@OldSongId", oldId);
            });
        }

        /// <summary>Takes a deleted song out of every 最愛用戶's 我的最愛.</summary>
        public int DeleteFavorite(string songId)
        {
            return Execute("delete from ktv_Favorite where Song_Id = @SongId", cmd => cmd.Parameters.AddWithValue("@SongId", songId));
        }

        public int AddFavoriteUser(string userId, string name)
        {
            return Execute("insert into ktv_User ( User_Id, User_Name ) values ( @UserId, @UserName )", cmd =>
            {
                cmd.Parameters.AddWithValue("@UserId", userId);
                cmd.Parameters.AddWithValue("@UserName", name);
            });
        }

        /// <summary>Removes the 最愛用戶 and all of their 我的最愛. Returns the number of ktv_User rows removed.</summary>
        public int RemoveFavoriteUser(string userId)
        {
            Execute("delete from ktv_Favorite where User_Id = @UserId", cmd => cmd.Parameters.AddWithValue("@UserId", userId));
            return Execute("delete from ktv_User where User_Id = @UserId", cmd => cmd.Parameters.AddWithValue("@UserId", userId));
        }

        /// <summary>
        /// The 最愛用戶 ends up named <paramref name="name"/> with exactly <paramref name="songIds"/> as 我的最愛
        /// (錢櫃排行 lists). Returns the number of songs written.
        /// </summary>
        public int ReplaceFavoriteUser(string userId, string name, IEnumerable<string> songIds)
        {
            RemoveFavoriteUser(userId);
            AddFavoriteUser(userId, name);
            return AddFavorites(userId, songIds);
        }

        /// <summary>
        /// Adds songs to the user's 我的最愛, skipping any already there or repeated in <paramref name="songIds"/>
        /// (ktv_Favorite's key is User_Id + Song_Id). Returns the number added.
        /// </summary>
        public int AddFavorites(string userId, IEnumerable<string> songIds)
        {
            HashSet<string> present = FavoriteSongIds(userId);
            int added = 0;
            foreach (string songId in songIds)
            {
                if (!present.Add(songId)) continue;
                added += Execute("insert into ktv_Favorite ( User_Id, Song_Id ) values ( @UserId, @SongId )", cmd =>
                {
                    cmd.Parameters.AddWithValue("@UserId", userId);
                    cmd.Parameters.AddWithValue("@SongId", songId);
                });
            }
            return added;
        }

        /// <summary>Takes songs out of one user's 我的最愛. Returns the number removed.</summary>
        public int RemoveFavorites(string userId, IEnumerable<string> songIds)
        {
            int removed = 0;
            foreach (string songId in songIds)
            {
                removed += Execute("delete from ktv_Favorite where User_Id = @UserId and Song_Id = @SongId", cmd =>
                {
                    cmd.Parameters.AddWithValue("@UserId", userId);
                    cmd.Parameters.AddWithValue("@SongId", songId);
                });
            }
            return removed;
        }

        /// <summary>Empties ktv_User and ktv_Favorite before a 我的最愛 import.</summary>
        public void ClearFavoriteUsers()
        {
            Execute("delete * from ktv_User", cmd => { });
            Execute("delete * from ktv_Favorite", cmd => { });
        }

        private HashSet<string> FavoriteSongIds(string userId)
        {
            var ids = new HashSet<string>();
            lock (commands)
            {
                using (DbCommand cmd = Db.Command("select Song_Id from ktv_Favorite where User_Id = @UserId", conn))
                {
                    cmd.Transaction = transaction;
                    cmd.Parameters.AddWithValue("@UserId", userId);
                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read()) ids.Add(Convert.ToString(reader.GetValue(0)));
                    }
                }
            }
            return ids;
        }

        #endregion

        #region --- 歌手資料表 ---

        public int InsertSinger(SingerTable table, string singerId, string name, string singerType, NameSpell spell)
        {
            string sql = "insert into " + table.Name + " ( Singer_Id, Singer_Name, Singer_Type, Singer_Spell, Singer_Strokes, Singer_SpellNum, Singer_PenStyle ) values ( @SingerId, @SingerName, @SingerType, @SingerSpell, @SingerStrokes, @SingerSpellNum, @SingerPenStyle )";
            return Execute(sql, cmd =>
            {
                cmd.Parameters.AddWithValue("@SingerId", singerId);
                cmd.Parameters.AddWithValue("@SingerName", name);
                cmd.Parameters.AddWithValue("@SingerType", singerType);
                AddSpell(cmd, spell);
            });
        }

        /// <summary>Rewrites the singer whose Singer_Id is <paramref name="oldId"/>.</summary>
        public int UpdateSinger(SingerTable table, string oldId, string singerId, string name, string singerType, NameSpell spell)
        {
            string sql = "update " + table.Name + " set Singer_Id = @SingerId, Singer_Name = @SingerName, Singer_Type = @SingerType, Singer_Spell = @SingerSpell, Singer_Strokes = @SingerStrokes, Singer_SpellNum = @SingerSpellNum, Singer_PenStyle = @SingerPenStyle where Singer_Id = @OldSingerId";
            return Execute(sql, cmd =>
            {
                cmd.Parameters.AddWithValue("@SingerId", singerId);
                cmd.Parameters.AddWithValue("@SingerName", name);
                cmd.Parameters.AddWithValue("@SingerType", singerType);
                AddSpell(cmd, spell);
                cmd.Parameters.AddWithValue("@OldSingerId", oldId);
            });
        }

        /// <summary>Rewrites only the 拼音 / 筆劃 / 筆形 columns (拼音校正).</summary>
        public int UpdateSingerSpell(SingerTable table, string singerId, NameSpell spell)
        {
            string sql = "update " + table.Name + " set Singer_Spell = @SingerSpell, Singer_Strokes = @SingerStrokes, Singer_SpellNum = @SingerSpellNum, Singer_PenStyle = @SingerPenStyle where Singer_Id = @SingerId";
            return Execute(sql, cmd =>
            {
                AddSpell(cmd, spell);
                cmd.Parameters.AddWithValue("@SingerId", singerId);
            });
        }

        /// <summary>Empties the table before a 歌手資料 import or a 重建歌庫歌手.</summary>
        public int DeleteAllSingers(SingerTable table)
        {
            return Execute("delete * from " + table.Name, cmd => { });
        }

        public int DeleteSinger(SingerTable table, string singerId)
        {
            return Execute("delete from " + table.Name + " where Singer_Id = @SingerId", cmd => cmd.Parameters.AddWithValue("@SingerId", singerId));
        }

        /// <summary>Binds in the statements' column order: Spell, Strokes, SpellNum, PenStyle.</summary>
        private static void AddSpell(DbCommand cmd, NameSpell spell)
        {
            cmd.Parameters.AddWithValue("@SingerSpell", spell.Spell);
            cmd.Parameters.AddWithValue("@SingerStrokes", spell.Strokes);
            cmd.Parameters.AddWithValue("@SingerSpellNum", spell.SpellNum);
            cmd.Parameters.AddWithValue("@SingerPenStyle", spell.PenStyle);
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

        /// <summary>The value bound for a ktv_Song column: text, except Song_CreatDate (see <see cref="Db.DateParam"/>).</summary>
        private object ParamValue(SongColumns column, string value)
        {
            return (column == SongColumns.CreatDate) ? Db.DateParam(conn, value) : value;
        }

        /// <summary>One statement at a time: a DbConnection is not safe to share between threads.</summary>
        private int Execute(string sql, Action<DbCommand> bind)
        {
            lock (commands)
            {
                DbCommand cmd;
                if (!commands.TryGetValue(sql, out cmd))
                {
                    cmd = Db.Command(sql, conn);
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

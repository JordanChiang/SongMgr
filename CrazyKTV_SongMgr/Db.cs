using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CrazyKTV_SongMgr
{
    enum DbKind
    {
        Access,
        SQLite
    }

    /// <summary>
    /// 資料庫存取: the single entry point for CrazySong / CrazySongMgr, whichever file holds them.
    /// .db / .sqlite / .sqlite3 files are SQLite (the format 匯出 SQLite 資料庫 writes and UltimateKtv reads),
    /// anything else is Access. Callers write Access SQL with "@Name" parameters added in the order they
    /// appear; on SQLite the few Access-only forms SongMgr uses are rewritten by <see cref="ToSQLite"/>.
    /// </summary>
    static class Db
    {
        public static readonly string[] SQLiteExtensions = { ".db", ".sqlite", ".sqlite3" };

        /// <summary>The 開啟歌庫 file dialog filter: every format <see cref="Db"/> can open, then each on its own.</summary>
        public const string FileDialogFilter = "資料庫檔案 (*.mdb;*.db;*.sqlite;*.sqlite3)|*.mdb;*.db;*.sqlite;*.sqlite3|Access 資料庫 (*.mdb)|*.mdb|SQLite 資料庫 (*.db;*.sqlite;*.sqlite3)|*.db;*.sqlite;*.sqlite3";

        public static DbKind KindOf(string databaseFile)
        {
            string ext = Path.GetExtension(databaseFile ?? "");
            return SQLiteExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)) ? DbKind.SQLite : DbKind.Access;
        }

        public static bool IsSQLite(string databaseFile) => KindOf(databaseFile) == DbKind.SQLite;

        public static bool IsSQLite(DbConnection conn) => conn is SQLiteConnection;

        /// <summary>
        /// The SQLite copy of <paramref name="accessFile"/> (same name, .db) when it exists, otherwise the Access
        /// file itself: putting CrazySongMgr.db into the SongMgr folder makes it take over from CrazySongMgr.mdb.
        /// </summary>
        public static string PreferSQLite(string accessFile)
        {
            string sqliteFile = Path.ChangeExtension(accessFile, ".db");
            return File.Exists(sqliteFile) ? sqliteFile : accessFile;
        }

        /// <summary>
        /// The value bound for a date column given as text. Access converts the text itself; SQLite keeps text as
        /// given, and a DATETIME column holding "2024/11/1 下午 03:04:05" cannot be read back, so on SQLite the
        /// date goes in as a DateTime.
        /// </summary>
        public static object DateParam(DbConnection conn, string text)
        {
            if (IsSQLite(conn) && DateTime.TryParse(text, out DateTime date)) return date;
            return text;
        }

        /// <summary>Opens the database and throws when it cannot be opened.</summary>
        public static DbConnection Open(string databaseFile, string password = "")
        {
            if (IsSQLite(databaseFile)) return DbHelper.SQLite.OpenConn(databaseFile);
            return DbHelper.Access.OpenConn(databaseFile, password);
        }

        /// <summary>
        /// Like <see cref="Open"/>, but a database that cannot be opened comes back closed instead of throwing,
        /// so the failure shows on the first command (the old CommonFunc.OleDbOpenConn behaviour).
        /// </summary>
        public static DbConnection OpenConn(string databaseFile, string password = "")
        {
            try
            {
                return Open(databaseFile, password);
            }
            catch
            {
                if (IsSQLite(databaseFile)) return new SQLiteConnection(DbHelper.SQLite.ConnectionString(databaseFile));
                return new OleDbConnection(DbHelper.Access.ConnectionString(databaseFile, password));
            }
        }

        /// <summary>A command on <paramref name="conn"/>, with Access SQL rewritten for SQLite when needed.</summary>
        public static DbCommand Command(string sql, DbConnection conn)
        {
            DbCommand cmd = conn.CreateCommand();
            cmd.CommandText = IsSQLite(conn) ? ToSQLite(sql) : sql;
            return cmd;
        }

        /// <summary>An adapter over a command the caller owns and disposes.</summary>
        public static DbDataAdapter Adapter(DbCommand cmd)
        {
            if (cmd is SQLiteCommand sqliteCmd) return new SQLiteDataAdapter(sqliteCmd);
            return new OleDbDataAdapter((OleDbCommand)cmd);
        }

        /// <summary>
        /// An adapter that owns its select command and disposes it with itself. On SQLite an undisposed
        /// command keeps the database file open after the connection is closed.
        /// </summary>
        public static DbDataAdapter Adapter(string sql, DbConnection conn)
        {
            if (conn is SQLiteConnection sqliteConn) return new SQLiteDataAdapter(ToSQLite(sql), sqliteConn);
            return new OleDbDataAdapter(sql, (OleDbConnection)conn);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:必須檢視 SQL 查詢中是否有安全性弱點")]
        public static DataTable GetDataTable(string databaseFile, string sql, string password = "")
        {
            using (DbConnection conn = OpenConn(databaseFile, password))
            {
                using (DbDataAdapter da = Adapter(sql, conn))
                {
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    return dt;
                }
            }
        }

        public static List<string> GetTableList(string databaseFile, string password = "")
        {
            if (!File.Exists(databaseFile)) return new List<string>();
            using (DbConnection conn = Open(databaseFile, password))
            {
                return GetTableList(conn);
            }
        }

        public static List<string> GetTableList(DbConnection conn)
        {
            if (conn is SQLiteConnection sqliteConn) return DbHelper.SQLite.GetTableList(sqliteConn);
            return DbHelper.Access.GetTableList((OleDbConnection)conn);
        }

        public static List<string> GetColumnList(string databaseFile, string password, string tableName)
        {
            using (DbConnection conn = Open(databaseFile, password))
            {
                if (conn is SQLiteConnection sqliteConn) return DbHelper.SQLite.GetColumnList(sqliteConn, tableName);
                return DbHelper.Access.GetColumnList((OleDbConnection)conn, tableName);
            }
        }

        /// <summary>壓縮並修復資料庫: JRO CompactDatabase for Access, VACUUM for SQLite.</summary>
        public static void Compact(string databaseFile)
        {
            if (IsSQLite(databaseFile)) DbHelper.SQLite.Compact(databaseFile);
            else DbHelper.Access.Compact(databaseFile);
        }

        /// <summary>
        /// Same as OleDbParameterCollection.AddWithValue, for either provider. Both bind the parameters SongMgr
        /// adds in the order they appear in the SQL, named "@Name" as in the SQL text.
        /// </summary>
        public static DbParameter AddWithValue(this DbParameterCollection parameters, string name, object value)
        {
            DbParameter p = (parameters is SQLiteParameterCollection) ? (DbParameter)new SQLiteParameter() : new OleDbParameter();
            p.ParameterName = name;
            p.Value = value ?? DBNull.Value;
            parameters.Add(p);
            return p;
        }

        // "InStr(1,LCase(Col),LCase('text'),0)": case-insensitive contains, as 歌庫查詢 and 歌手管理 build it.
        private static readonly Regex InStrLCase = new Regex(@"InStr\(1,\s*LCase\(([^()]+)\),\s*LCase\(('(?:[^']|'')*')\),\s*0\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex DeleteStar = new Regex(@"^(\s*delete)\s+\*\s+(from\b)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SelectTop = new Regex(@"^(\s*select)\s+top\s+(\d+)\s+(.*)$", RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
        // "First(Col)" in a GROUP BY query: SQLite has no First aggregate, but a bare column there takes its value
        // from one row of the group, which is all Access promises for First too.
        private static readonly Regex FirstAggregate = new Regex(@"\bFirst\(([^()]+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Rewrites the Access-only SQL SongMgr uses into SQLite: "delete * from", "select top N",
        /// InStr / LCase contains tests and the First() aggregate. Everything else (brackets, "@Name"
        /// parameters, like '%x%', Count, In (select ...)) already means the same in both.
        /// </summary>
        public static string ToSQLite(string sql)
        {
            sql = DeleteStar.Replace(sql, "$1 $2");
            sql = InStrLCase.Replace(sql, "instr(lower($1), lower($2))");
            sql = FirstAggregate.Replace(sql, "$1");
            Match top = SelectTop.Match(sql);
            if (top.Success) sql = top.Groups[1].Value + " " + top.Groups[3].Value.TrimEnd().TrimEnd(';') + " limit " + top.Groups[2].Value;
            return sql;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SQLite;
using System.IO;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// The two database providers under <see cref="Db"/>: Access (.mdb, Jet OLE DB) and SQLite (.db / .sqlite).
    /// Callers go through <see cref="Db"/>, which picks one by file extension.
    /// </summary>
    static class DbHelper
    {
        public static class Access
        {
            /// <summary>Jet connection string for a database file.</summary>
            public static string ConnectionString(string databaseFile, string password = "")
            {
                string cnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databaseFile + ";";
                if (!string.IsNullOrEmpty(password)) cnstr += "Jet OLEDB:Database Password=" + password + ";";
                return cnstr;
            }

            public static OleDbConnection OpenConn(string databaseFile, string password = "")
            {
                OleDbConnection conn = new OleDbConnection(ConnectionString(databaseFile, password));
                try
                {
                    conn.Open();
                }
                catch
                {
                    conn.Dispose();
                    throw;
                }
                return conn;
            }

            public static List<string> GetTableList(OleDbConnection conn)
            {
                List<string> list = new List<string>();
                using (DataTable dt = conn.GetSchema("Tables"))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["TABLE_TYPE"].ToString() == "TABLE") list.Add(row["TABLE_NAME"].ToString());
                    }
                }
                return list;
            }

            public static List<string> GetColumnList(OleDbConnection conn, string tableName)
            {
                List<string> list = new List<string>();
                using (DataTable dt = conn.GetSchema("Columns", new string[] { null, null, tableName, null }))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        list.Add(row["COLUMN_NAME"].ToString());
                    }
                }
                return list;
            }

            public static void Compact(string databaseFile)
            {
                CommonFunc.CompactAccessDB(ConnectionString(databaseFile), databaseFile);
            }
        }

        public static class SQLite
        {
            public static string ConnectionString(string databaseFile)
            {
                return new SQLiteConnectionStringBuilder { DataSource = databaseFile, Version = 3 }.ConnectionString;
            }

            /// <summary>Opens an existing file; a missing file is an error rather than a new empty database.</summary>
            public static SQLiteConnection OpenConn(string databaseFile)
            {
                if (!File.Exists(databaseFile)) throw new FileNotFoundException("找不到資料庫檔案。", databaseFile);

                SQLiteConnection conn = new SQLiteConnection(ConnectionString(databaseFile));
                try
                {
                    conn.Open();
                }
                catch
                {
                    conn.Dispose();
                    throw;
                }
                return conn;
            }

            /// <summary>GetSchema("Tables") does not report TABLE_TYPE "TABLE" for SQLite, so ask sqlite_master.</summary>
            public static List<string> GetTableList(SQLiteConnection conn)
            {
                List<string> list = new List<string>();
                using (SQLiteCommand cmd = new SQLiteCommand("select name from sqlite_master where type = 'table' and name not like 'sqlite_%'", conn))
                {
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read()) list.Add(reader.GetString(0));
                    }
                }
                return list;
            }

            public static List<string> GetColumnList(SQLiteConnection conn, string tableName)
            {
                List<string> list = new List<string>();
                using (SQLiteCommand cmd = new SQLiteCommand("pragma table_info(\"" + tableName.Replace("\"", "\"\"") + "\")", conn))
                {
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read()) list.Add(reader["name"].ToString());
                    }
                }
                return list;
            }

            public static void Compact(string databaseFile)
            {
                using (SQLiteConnection conn = OpenConn(databaseFile))
                {
                    using (SQLiteCommand cmd = new SQLiteCommand("vacuum", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}

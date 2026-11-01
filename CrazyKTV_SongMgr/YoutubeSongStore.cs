using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;

namespace CrazyKTV_SongMgr
{
    /// <summary>YoutubeSongList columns 歌庫查詢 lets the user change.</summary>
    enum YoutubeColumn
    {
        Name,
        Track,
        Volume,
        PlayCount,
    }

    /// <summary>
    /// YouTube 歌曲: UltimateKtv records each YouTube video it downloads as a row of YoutubeSongList in
    /// the 歌庫 database (CrazySong.mdb or CrazySong.db). SongMgr lists, edits and deletes those rows; only UltimateKtv adds them.
    /// Every method opens its own connection and throws when the database cannot be read or written.
    /// </summary>
    static class YoutubeSongStore
    {
        public const string TableName = "YoutubeSongList";

        /// <summary>
        /// The only 聲道 values a YouTube 歌曲 may have: the two AudioChannel states UltimateKtv's
        /// 人聲 / 伴唱 buttons switch a YouTube video between (MainWindow.Player.cs).
        /// </summary>
        public static readonly (int Value, string Display)[] Tracks = { (0, "立體聲 (人聲)"), (3, "去人聲 (伴唱)") };

        /// <summary>
        /// UltimateKtv's CheckAndCreateYoutubeTable (Library/SongDatas.cs): Url is MEMO on Access and TEXT on
        /// SQLite, where MEMO would give the column numeric affinity.
        /// </summary>
        private static string CreateTableSql(string databaseFile) => "create table YoutubeSongList ( Id TEXT(5) PRIMARY KEY, Name TEXT(255), CreatDate DATETIME, Url " + (Db.IsSQLite(databaseFile) ? "TEXT" : "MEMO") + ", Singer TEXT(50), Lang TEXT(16), PlayCount INTEGER, Length INTEGER, FileName TEXT(255), FileSize INTEGER, Track INTEGER DEFAULT 3, Volume INTEGER, WordCount INTEGER )";

        /// <summary>
        /// The columns 歌庫查詢 shows, in display order. Singer, Lang and WordCount are left out: UltimateKtv
        /// always writes them empty (0) and never reads them for a YouTube 歌曲.
        /// </summary>
        private const string QuerySql = "select Id, Name, Url, Length, PlayCount, Track, Volume, FileName, FileSize, CreatDate from YoutubeSongList order by Id";

        public static bool TableExists(string databaseFile)
        {
            using (DbConnection conn = Open(databaseFile))
            {
                return Db.GetTableList(conn).Contains(TableName);
            }
        }

        public static void CreateTable(string databaseFile)
        {
            using (DbConnection conn = Open(databaseFile))
            {
                using (DbCommand cmd = Db.Command(CreateTableSql(databaseFile), conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>Every row, or only those whose Name contains <paramref name="keyword"/> (ignoring case), ordered by Id.</summary>
        public static DataTable Query(string databaseFile, string keyword)
        {
            DataTable dt = new DataTable();
            using (DbConnection conn = Open(databaseFile))
            {
                using (DbDataAdapter da = Db.Adapter(QuerySql, conn))
                {
                    da.Fill(dt);
                }
            }

            if (!string.IsNullOrWhiteSpace(keyword))
            {
                string word = keyword.Trim();
                var misses = dt.AsEnumerable().Where(row => row["Name"].ToString().IndexOf(word, StringComparison.OrdinalIgnoreCase) < 0).ToList();
                foreach (DataRow row in misses) dt.Rows.Remove(row);
            }
            dt.AcceptChanges();
            return dt;
        }

        /// <summary>
        /// Checks <paramref name="input"/> for <paramref name="column"/>. Returns why it is not allowed, or ""
        /// with the value to store in <paramref name="value"/>.
        /// </summary>
        public static string Validate(YoutubeColumn column, object input, out object value)
        {
            string text = (input == null || input == DBNull.Value) ? "" : input.ToString().Trim();
            value = null;
            switch (column)
            {
                case YoutubeColumn.Name:
                    if (text == "") return "名稱不可空白";
                    if (text.Length > 255) return "名稱不可超過 255 個字";
                    value = text;
                    return "";
                case YoutubeColumn.Track:
                    if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int track) || !Tracks.Any(t => t.Value == track)) return "聲道只能是 " + string.Join(" 或 ", Tracks.Select(t => t.Value));
                    value = track;
                    return "";
                case YoutubeColumn.Volume:
                    return ValidateInt(text, 0, 100, "音量", out value);
                case YoutubeColumn.PlayCount:
                    return ValidateInt(text, 0, int.MaxValue, "點播次數", out value);
                default:
                    throw new ArgumentOutOfRangeException(nameof(column));
            }
        }

        /// <summary>Writes one column of the row <paramref name="id"/>. Returns the number of rows changed.</summary>
        public static int Update(string databaseFile, string id, YoutubeColumn column, object input)
        {
            string error = Validate(column, input, out object value);
            if (error != "") throw new ArgumentException(error, nameof(input));

            using (DbConnection conn = Open(databaseFile))
            {
                using (DbCommand cmd = Db.Command("update YoutubeSongList set [" + column + "] = @Value where Id = @Id", conn))
                {
                    cmd.Parameters.AddWithValue("@Value", value);
                    cmd.Parameters.AddWithValue("@Id", id);
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>Removes the rows; the downloaded video files stay where UltimateKtv put them.</summary>
        public static int Delete(string databaseFile, IEnumerable<string> ids)
        {
            int count = 0;
            using (DbConnection conn = Open(databaseFile))
            {
                using (DbCommand cmd = Db.Command("delete from YoutubeSongList where Id = @Id", conn))
                {
                    foreach (string id in ids)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@Id", id);
                        count += cmd.ExecuteNonQuery();
                    }
                }
            }
            return count;
        }

        /// <summary>Length (seconds) as 分:秒, e.g. 221 → 3:41.</summary>
        public static string FormatLength(object seconds)
        {
            if (seconds == null || seconds == DBNull.Value) return "";
            int total = Convert.ToInt32(seconds);
            return (total / 60) + ":" + (total % 60).ToString("D2");
        }

        /// <summary>FileSize (bytes) as whole MB, rounded.</summary>
        public static string FormatMB(object bytes)
        {
            if (bytes == null || bytes == DBNull.Value) return "";
            return Math.Round(Convert.ToInt64(bytes) / 1024m / 1024m, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>The address UltimateKtv downloaded the video from; "" when Url holds no video ID.</summary>
        public static string WatchUrl(object videoId)
        {
            string id = (videoId == null || videoId == DBNull.Value) ? "" : videoId.ToString().Trim();
            return (id == "") ? "" : "https://www.youtube.com/watch?v=" + Uri.EscapeDataString(id);
        }

        private static string ValidateInt(string text, int min, int max, string label, out object value)
        {
            value = null;
            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number) || number < min || number > max)
            {
                return (max == int.MaxValue) ? label + "必須是 " + min + " 以上的整數" : label + "必須是 " + min + " 到 " + max + " 的整數";
            }
            value = number;
            return "";
        }

        private static DbConnection Open(string databaseFile) => Db.Open(databaseFile);
    }
}

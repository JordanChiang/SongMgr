using System.Data;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// A "Display / Value" list shown in a ListBox: 操作記錄, 加歌結果 failures, 加歌結果 duplicates.
    /// Every write goes through one lock, so background tasks can log without corrupting the table.
    /// </summary>
    class OperationLog
    {
        private readonly object gate = new object();
        private DataTable table = NewTable();

        /// <summary>The table the ListBox binds to. Replaced by <see cref="Reset"/>.</summary>
        public DataTable Table
        {
            get { lock (gate) { return table; } }
        }

        public int Count
        {
            get { lock (gate) { return table.Rows.Count; } }
        }

        /// <summary>Appends a line; its Value is its 1-based position.</summary>
        public void Add(string message)
        {
            lock (gate)
            {
                DataRow row = table.NewRow();
                row["Display"] = message;
                row["Value"] = table.Rows.Count + 1;
                table.Rows.Add(row);
            }
        }

        public void Clear()
        {
            lock (gate) { table.Rows.Clear(); }
        }

        /// <summary>Starts a fresh table, leaving any ListBox bound to the old one untouched.</summary>
        public void Reset()
        {
            lock (gate) { table = NewTable(); }
        }

        private static DataTable NewTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Display", typeof(string)));
            dt.Columns.Add(new DataColumn("Value", typeof(int)));
            return dt;
        }
    }
}

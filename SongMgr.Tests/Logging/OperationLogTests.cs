using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Logging
{
    public class OperationLogTests
    {
        /// <summary>The three-line idiom every log write used before the module existed.</summary>
        static void LegacyAdd(DataTable dt, string message)
        {
            dt.Rows.Add(dt.NewRow());
            dt.Rows[dt.Rows.Count - 1][0] = message;
            dt.Rows[dt.Rows.Count - 1][1] = dt.Rows.Count;
        }

        static DataTable LegacyTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Display", typeof(string)));
            dt.Columns.Add(new DataColumn("Value", typeof(int)));
            return dt;
        }

        static List<string> Dump(DataTable dt) =>
            dt.AsEnumerable().Select(r => r[0] + "#" + r[1] + "#" + r[0].GetType().Name + "#" + r[1].GetType().Name).ToList();

        [Fact]
        public void Sequential_writes_produce_the_same_table_as_the_legacy_idiom()
        {
            string[] messages = { "【歌庫查詢】a", "", "【加歌頁面】c|d", "x" };
            DataTable legacy = LegacyTable();
            OperationLog log = new OperationLog();

            foreach (string m in messages)
            {
                LegacyAdd(legacy, m);
                log.Add(m);
            }

            Assert.Equal(Dump(legacy), Dump(log.Table));
            Assert.Equal(legacy.Columns.Cast<DataColumn>().Select(c => c.ColumnName + c.DataType), log.Table.Columns.Cast<DataColumn>().Select(c => c.ColumnName + c.DataType));
            Assert.Equal(4, log.Count);
        }

        [Fact]
        public void Clear_then_add_restarts_numbering_like_the_legacy_table()
        {
            DataTable legacy = LegacyTable();
            OperationLog log = new OperationLog();
            LegacyAdd(legacy, "a"); log.Add("a");
            legacy.Rows.Clear(); log.Clear();
            LegacyAdd(legacy, "b"); log.Add("b");

            Assert.Equal(Dump(legacy), Dump(log.Table));
        }

        [Fact]
        public void Reset_starts_a_new_table_and_leaves_the_old_one_intact()
        {
            OperationLog log = new OperationLog();
            log.Add("old");
            DataTable before = log.Table;

            log.Reset();
            log.Add("new");

            Assert.NotSame(before, log.Table);
            Assert.Equal(1, before.Rows.Count);
            Assert.Equal("new", log.Table.Rows[0][0]);
        }

        [Fact]
        public void Concurrent_writes_keep_every_line_and_number_them_1_to_n()
        {
            OperationLog log = new OperationLog();
            const int n = 20000;

            Parallel.For(0, n, i => log.Add("m" + i));

            Assert.Equal(n, log.Count);
            Assert.Equal(Enumerable.Range(1, n), log.Table.AsEnumerable().Select(r => (int)r[1]));
            Assert.Equal(n, log.Table.AsEnumerable().Select(r => (string)r[0]).Distinct().Count());
        }
    }
}

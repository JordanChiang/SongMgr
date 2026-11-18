using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Compare
{
    /// <summary>
    /// 新舊版比對的資料, made once from a real Access 歌庫 (SONGMGR_COMPARE_SOURCE) into SONGMGR_COMPARE_DATA, so both
    /// versions start from the same files:
    /// <list type="bullet">
    /// <item>query.mdb: the 歌庫 with every Song_Path moved under lib\ (k:\國語\ becomes lib\k\國語\).</item>
    /// <item>lib\: an empty file for every song but one in ten, so 檔案不存在 has something to find.</item>
    /// <item>add.mdb: query.mdb without one song in ten (the 新歌), for 加歌.</item>
    /// <item>add_new\: the 新歌' files; add_all\: every song's file; add_mixed\: the 新歌 plus up to 2000 songs already
    /// in add.mdb, for the 重複歌曲 check. Songs whose folders name no 語系 are left out of these, since 加歌 refuses a
    /// list with a 語系 of 未知.</item>
    /// </list>
    /// </summary>
    [Trait("Category", "Compare")]
    public class CompareData
    {
        public const string NewSongs = "add_new";
        public const string AllSongs = "add_all";
        public const string Mixed = "add_mixed";

        sealed class Song
        {
            public string Id, Path, FileName;
        }

        [CompareFact("SONGMGR_COMPARE_SOURCE")]
        public void Prepare()
        {
            string source = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_SOURCE");
            string data = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_DATA");
            string cfg = Environment.GetEnvironmentVariable("SONGMGR_COMPARE_CFG");
            Assert.False(string.IsNullOrEmpty(data), "SONGMGR_COMPARE_DATA is not set");
            Assert.True(File.Exists(source), source + " does not exist");
            Assert.Equal(".mdb", Path.GetExtension(source).ToLowerInvariant());

            List<string> langs = new List<string> { "國語", "台語", "粵語", "日語", "英語", "客語", "原住民語", "韓語", "兒歌", "其它" };
            if (!string.IsNullOrEmpty(cfg) && File.Exists(cfg))
            {
                string configured = SettingsFile.Load(cfg)["CrazyktvSongLangStr"];
                if (configured.Split(',').Length >= 10) langs = configured.Split(',').ToList();
            }

            if (Directory.Exists(data)) Directory.Delete(data, true);
            Directory.CreateDirectory(data);
            string query = Path.Combine(data, "query.mdb");
            File.Copy(source, query);
            File.SetAttributes(query, FileAttributes.Normal);

            List<Song> songs;
            using (DataTable dt = Db.GetDataTable(query, "select Song_Id, Song_Path, Song_FileName from ktv_Song order by Song_Id, Song_Path, Song_FileName", ""))
            {
                songs = dt.AsEnumerable().Select(r => new Song { Id = r[0].ToString(), Path = r[1].ToString(), FileName = r[2].ToString() }).ToList();
            }

            var newSongs = new List<Song>();
            int noFile = 0, noLang = 0, mixedOld = 0;
            for (int i = 0; i < songs.Count; i++)
            {
                Song s = songs[i];
                string rel = Relative(s.Path);
                bool isNew = i % 10 == 3;
                if (i % 10 == 7) noFile++;
                else Touch(Path.Combine(data, "lib", rel, s.FileName));

                if (!rel.Split('\\').Any(langs.Contains))
                {
                    noLang++;
                }
                else
                {
                    Touch(Path.Combine(data, AllSongs, rel, s.FileName));
                    if (isNew)
                    {
                        Touch(Path.Combine(data, NewSongs, rel, s.FileName));
                        Touch(Path.Combine(data, Mixed, rel, s.FileName));
                    }
                    else if (i % 16 == 5 && mixedOld < 2000)
                    {
                        Touch(Path.Combine(data, Mixed, rel, s.FileName));
                        mixedOld++;
                    }
                }
                if (isNew) newSongs.Add(s);
            }

            using (var conn = DbHelper.Access.OpenConn(query))
            using (var tx = conn.BeginTransaction())
            {
                foreach (Song s in songs)
                {
                    string moved = Path.Combine(data, "lib", Relative(s.Path)).TrimEnd('\\') + "\\";
                    Execute(conn, tx, "update ktv_Song set Song_Path = @a where Song_Id = @b and Song_Path = @c and Song_FileName = @d",
                        moved, s.Id, s.Path, s.FileName);
                    s.Path = moved;
                }
                tx.Commit();
            }

            string add = Path.Combine(data, "add.mdb");
            File.Copy(query, add);
            using (var conn = DbHelper.Access.OpenConn(add))
            using (var tx = conn.BeginTransaction())
            {
                foreach (Song s in newSongs)
                {
                    Execute(conn, tx, "delete from ktv_Song where Song_Id = @a and Song_Path = @b and Song_FileName = @c", s.Id, s.Path, s.FileName);
                }
                tx.Commit();
            }

            File.WriteAllText(Path.Combine(data, "README.txt"),
                "source: " + source + "\r\n" +
                "songs: " + songs.Count + "\r\n" +
                "without a file under lib: " + noFile + "\r\n" +
                "新歌 (not in add.mdb): " + newSongs.Count + "\r\n" +
                "songs already in add.mdb in " + Mixed + ": " + mixedOld + "\r\n" +
                "left out of the add folders (no 語系 folder): " + noLang + "\r\n", new UTF8Encoding(false));
        }

        /// <summary>k:\國語\男\ becomes k\國語\男\; \\server\share\x\ becomes unc\server\share\x\.</summary>
        static string Relative(string path)
        {
            if (path.Length >= 3 && path[1] == ':' && path[2] == '\\') return path[0] + "\\" + path.Substring(3);
            if (path.StartsWith("\\\\")) return "unc\\" + path.Substring(2);
            return "other\\" + path.TrimStart('\\');
        }

        static void Touch(string file)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(file));
            File.WriteAllBytes(file, new byte[0]);
        }

        static void Execute(OleDbConnection conn, OleDbTransaction tx, string sql, params string[] values)
        {
            using (var cmd = new OleDbCommand(sql, conn, tx))
            {
                foreach (string v in values) cmd.Parameters.AddWithValue("@p", v);
                cmd.ExecuteNonQuery();
            }
        }
    }
}

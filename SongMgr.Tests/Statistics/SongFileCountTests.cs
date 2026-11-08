using System;
using System.Data;
using System.IO;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Statistics
{
    /// <summary>歌庫檔案統計: counts the songs whose file exists, listing each folder once.</summary>
    public class SongFileCountTests : IDisposable
    {
        readonly string root = TestEnv.NewTempDir("filecount");

        public void Dispose() => TestEnv.DeleteDir(root);

        static DataTable Songs(params (string Path, string FileName)[] rows)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Song_Path", typeof(string));
            dt.Columns.Add("Song_FileName", typeof(string));
            foreach (var r in rows) dt.Rows.Add(r.Path, r.FileName);
            return dt;
        }

        string Touch(string folder, string name)
        {
            string dir = Path.Combine(root, folder);
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, name), "");
            return dir;
        }

        [Fact]
        public void Counts_only_files_that_exist()
        {
            string a = Touch("A", "1.mpg");
            Touch("A", "2.mpg");
            string b = Touch("B", "3.mkv");
            Directory.CreateDirectory(Path.Combine(a, "sub.mpg")); // a folder named like a song is not a file

            DataTable dt = Songs(
                (a, "1.mpg"), (a, "2.mpg"), (a, "missing.mpg"), (a, "sub.mpg"),
                (b, "3.mkv"),
                (Path.Combine(root, "NoSuchFolder"), "4.mpg"));

            Assert.Equal(3, CommonFunc.GetSongFileCount(dt));
        }

        [Fact]
        public void Ignores_case_and_trailing_backslash_in_the_path()
        {
            string a = Touch("Songs", "Hello.MPG");

            DataTable dt = Songs(
                (a.ToUpperInvariant() + "\\", "hello.mpg"),
                (a, "HELLO.mpg"));

            Assert.Equal(2, CommonFunc.GetSongFileCount(dt));
        }

        [Fact]
        public void File_name_with_a_subfolder_is_still_found()
        {
            Touch(Path.Combine("Root", "Sub"), "5.mpg");

            DataTable dt = Songs((Path.Combine(root, "Root"), Path.Combine("Sub", "5.mpg")));

            Assert.Equal(1, CommonFunc.GetSongFileCount(dt));
        }

        [Fact]
        public void Empty_library_counts_zero()
        {
            Assert.Equal(0, CommonFunc.GetSongFileCount(Songs()));
        }
    }
}

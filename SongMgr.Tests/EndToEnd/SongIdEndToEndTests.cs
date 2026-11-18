using System.IO;
using System.Linq;
using Xunit;

namespace SongMgr.Tests.EndToEnd
{
    /// <summary>A theory that needs the desktop; skipped when <see cref="AppSession.SkipReason"/> says so.</summary>
    public sealed class DesktopTheoryAttribute : TheoryAttribute
    {
        public DesktopTheoryAttribute()
        {
            if (AppSession.SkipReason != null) Skip = AppSession.SkipReason;
        }
    }

    public sealed class DesktopFactAttribute : FactAttribute
    {
        public DesktopFactAttribute()
        {
            if (AppSession.SkipReason != null) Skip = AppSession.SkipReason;
        }
    }

    /// <summary>
    /// 歌曲編號 in the running program, for a 5-digit and a 6-digit 歌庫: 歌庫監視 adding songs, 歌庫查詢 moving a song
    /// to another 語系, 轉換為 6 位數, and the 資料庫結構更新 of an old 歌庫 at start-up, each checked in the sandbox's
    /// database and 設定檔. (A 歌庫 with both digit
    /// counts is not usable until 校正編碼位數, so it is not driven here.)
    /// </summary>
    [Trait("Category", "EndToEnd")]
    public class SongIdEndToEndTests
    {
        static string Id(int value, int digits) => value.ToString("D" + digits);
        static int Start(string lang, int digits) => int.Parse((digits == 5 ? AppSession.Code5 : AppSession.Code6).Split(',')[AppSession.Langs.IndexOf(lang)]);

        [DesktopTheory]
        [InlineData(5)]
        [InlineData(6)]
        public void 歌庫監視加歌_then_歌庫查詢_moves_a_song_to_another_語系(int digits)
        {
            int mandarin = Start("國語", digits), taiwanese = Start("台語", digits);
            using (var app = new AppSession("monitor" + digits))
            {
                app.WatchMonitorFolder(digits);
                app.InsertSong(Id(mandarin, digits), "國語", "周杰倫", "晴天", app.MakeSongFile(@"國語\周杰倫_晴天"));
                app.InsertSong(Id(mandarin + 1, digits), "國語", "王菲", "紅豆", app.MakeSongFile(@"國語\王菲_紅豆"));
                app.MakeSongFile(@"國語\蔡依林_日不落");
                app.MakeSongFile(@"台語\江蕙_家後");
                app.MakeSongFile(@"台語\陳雷_歡喜就好");
                app.MakeSongFile(@"台語\黃乙玲_講你也不知");

                app.Start();
                AppSession.WaitUntil(() => app.Songs().Count == 6, 120, "歌庫監視 to add 4 songs");
                // A new 設定檔 on an up-to-date 歌庫 asks nothing (it used to bring up 偵測到歌庫版本更新).
                Assert.Empty(app.Dialogs());

                var songs = app.Songs();
                // 台語 had no songs: it gets its start number once, then the next ones.
                Assert.Equal(new[] { Id(taiwanese, digits), Id(taiwanese + 1, digits), Id(taiwanese + 2, digits) },
                    songs.Where(s => s.Lang == "台語").Select(s => s.Id).OrderBy(x => x));
                Assert.Equal(Id(mandarin + 2, digits), songs.Single(s => s.Name == "日不落").Id);

                // 歌庫查詢: 日不落 to 台語 shows the next 台語 number, and keeps it.
                app.Click("編輯模式");
                app.ClickRow("SongQuery_DataGridView", Id(mandarin + 2, digits));
                AppSession.WaitUntil(() => app.Value("SongQuery_EditSongId_TextBox") == Id(mandarin + 2, digits), 20, "日不落 in the edit panel");
                app.Choose("SongQuery_EditSongLang_ComboBox", "台語");
                AppSession.WaitUntil(() => app.Value("SongQuery_EditSongId_TextBox") == Id(taiwanese + 3, digits), 20, "the next 台語 number in 歌曲編號");
                app.Click("套用變更(A)");
                AppSession.WaitUntil(() => app.Songs().Any(s => s.Name == "日不落" && s.Lang == "台語"), 30, "日不落 saved as 台語");

                Assert.Equal(Id(taiwanese + 3, digits), app.Songs().Single(s => s.Name == "日不落").Id);
                Assert.Equal(6, app.Songs().Select(s => s.Id).Distinct().Count());
            }
        }

        [DesktopFact]
        public void An_old_歌庫_gets_its_資料庫結構更新_when_SongMgr_starts()
        {
            using (var app = new AppSession("schema"))
            {
                app.WriteSettings(6);
                app.Execute("alter table ktv_Song drop column Song_ReplayGain");
                app.Execute("alter table ktv_Song add column Song_MeanVolume DOUBLE");
                app.InsertSong("100001", "國語", "周杰倫", "晴天");
                app.InsertFavorite("0001", "100001");
                Assert.NotEmpty(CrazyKTV_SongMgr.SchemaUpgrade.Check(app.Database));

                app.Start();
                AppSession.WaitUntil(() => CrazyKTV_SongMgr.SchemaUpgrade.Check(app.Database).Count == 0, 60, "the 資料庫結構更新");
                AppSession.WaitUntil(() => app.Value("SongMaintenance_DBVer1Value_Label") == "最新", 20, "資料庫結構: 最新");

                Assert.Empty(app.Dialogs());
                Assert.Equal(new[] { ("100001", "國語", "周杰倫", "晴天") }, app.Songs());
                Assert.Equal(new[] { "0001:100001" }, app.Favorites());
                Assert.Single(Directory.GetFiles(Path.Combine(app.AppDir, "SongMgr", "Backup"), "*_CrazySong.mdb"));
            }
        }

        [DesktopFact]
        public void 轉換為6位數_renumbers_every_song_keeps_我的最愛_and_saves_the_digit_count()
        {
            using (var app = new AppSession("to6"))
            {
                app.WriteSettings(5);
                app.InsertSong("10001", "國語", "周杰倫", "晴天");
                app.InsertSong("10002", "國語", "王菲", "紅豆");
                app.InsertSong("20005", "台語", "江蕙", "家後");
                app.InsertFavorite("0001", "10002");
                app.InsertFavorite("0001", "20005");

                app.Start();
                app.Click("轉換為 6 位數", onTab: "歌庫維護");
                app.Answer("確認提示", "是(Y)");
                AppSession.WaitUntil(() => app.Songs().All(s => s.Id.Length == 6), 60, "轉換為 6 位數 to finish");

                Assert.Equal(new[] { ("100001", "晴天"), ("100002", "紅豆"), ("200001", "家後") }, app.Songs().Select(s => (s.Id, s.Name)));
                Assert.Equal(new[] { "0001:100002", "0001:200001" }, app.Favorites());
                AppSession.WaitUntil(() => app.Setting("SongMgrMaxDigitCode") == "2", 20, "the 設定檔 to say 6 digits");
                Assert.Equal(AppSession.Code6, app.Setting("SongMgrLangCode"));
                Assert.False(File.Exists(Path.Combine(app.AppDir, "SongMgr", "Backup", "Favorite.txt")), "我的最愛 are no longer exported to Favorite.txt");
            }
        }
    }
}

# 歌庫監視加歌

加歌模式選「歌庫監視」（`SongMgrSongAddMode=4`）並啟用監視目錄後，程式啟動時和執行中會掃描監視目錄，把還不在歌庫裡的歌曲檔分析出語系、歌手、歌名，配上歌曲編號寫進 `ktv_Song`。歌曲編號由每個語系的起始編號（`SongMgrLangCode`）往後接，5 位數或 6 位數依 `SongMgrMaxDigitCode`。

## Sub-features

- 啟動時掃描監視目錄，補加新檔案。
- 依資料夾名稱（語系）和檔名（`歌手_歌名`）辨識歌曲。
- 語系第一次有歌時從起始編號開始，之後接著最大編號往下編；已經在歌庫裡的檔案不重加。
- 5 位數與 6 位數歌庫各自正確。
- 手動加歌（把檔案拖曳到加歌頁面的「請將要加入的歌曲檔案或資料夾拖曳至此處」，再按「加入歌庫」）：目前沒有端對端測試，見 Gotchas。

## How to get to it (user POV)

1. 程式設定裡選加歌模式「歌庫監視」：歌庫設定 › 監視目錄 分頁，勾選啟用並填入監視資料夾，按儲存。
2. 把歌曲檔放進 `<監視資料夾>\<語系>\<歌手>_<歌名>.mpg`。
3. 重新開程式（或等執行中的監視觸發），到歌庫查詢看新歌和它的歌曲編號。

## Driving it with AppSession

現有測試：`SongIdEndToEndTests.歌庫監視加歌_then_歌庫查詢_moves_a_song_to_another_語系`（5 位數、6 位數各一次）。

```csharp
using (var app = new AppSession("monitor6"))
{
    app.WatchMonitorFolder(6);                                   // 設定檔：加歌模式 4、啟用監視、監視 app.MonitorDir
    app.InsertSong("100001", "國語", "周杰倫", "晴天", app.MakeSongFile(@"國語\周杰倫_晴天")); // 已在歌庫的歌
    app.MakeSongFile(@"國語\蔡依林_日不落");                       // 新檔案
    app.MakeSongFile(@"台語\江蕙_家後");
    app.Start();
    AppSession.WaitUntil(() => app.Songs().Count == 3, 120, "歌庫監視 to add 2 songs");
    Assert.Empty(app.Dialogs());
}
```

證明它能用的結果：

- `app.Songs()` 剛好多出新檔案那幾首，已在歌庫的歌沒有重複。
- 新歌的編號：國語接在 `100001` 後面是 `100002`；台語第一首是台語的起始編號 `200001`。
- 新設定檔、最新版歌庫啟動時沒有任何對話框（`Dialogs()` 是空的）。
- 證據資料夾裡的 `CrazySong.mdb` 有這些列，`end.png` 是主視窗。

## Gotchas

- `MakeSongFile` 用沙盒 `Tools\ffmpeg.exe` 做 3 秒的 mpeg2 檔，每個約 1 秒；檔案太多會拖慢測試。
- 歌庫監視在背景執行，主視窗出現時還沒加完：一定要 `WaitUntil` 等 `Songs().Count`，給 120 秒。
- 手動加歌要把檔案**拖曳**到加歌頁面，UI Automation 做不到拖曳，`AppSession` 也沒有對應的操作。要驗證手動加歌，得先在 `AppSession` 加一個用 `WM_DROPFILES` 或剪貼簿模擬拖放的方法；在那之前，手動加歌的修改在回報裡列為「端對端未驗證」，只靠 `SongMgr.Tests\Analysis` 等單元測試。

# 歌曲編號維護

歌庫維護 › 維護雜項 分頁的兩個按鈕，整批改歌曲編號：

- 「轉換為 6 位數」：5 位數歌庫的每首歌改成 6 位數編號，我的最愛跟著換，設定檔改成 6 位數（`SongMgrMaxDigitCode=2`、`SongMgrLangCode` 換成 6 位數的起始編號）。
- 「校正編碼位數」：5、6 位數混在一起的歌庫，把位數不對的歌改成設定檔的位數。

## Sub-features

- 轉換為 6 位數：重新編號、保留我的最愛、存設定檔。
- 校正編碼位數：只改位數不對的歌。
- 語系是「未知」的歌兩者都算失敗、不改，其他照做；結果顯示在 `SongMaintenance_Tooltip_Label`（「總共轉換 N 首歌曲的歌曲編號,失敗 M 首」/「總共校正 ...」）。
- 不再匯出 `SongMgr\Backup\Favorite.txt`。
- 按鈕啟用依歌庫實際的位數：5 位數歌庫配上寫 6 位數的設定檔，第一次啟動「轉換為 6 位數」就能按。

## How to get to it (user POV)

1. 切到「歌庫維護」分頁（預設在「維護雜項」子分頁）。
2. 按「轉換為 6 位數」或「校正編碼位數」。
3. 確認對話框「確認提示」按「是(Y)」。
4. 看分頁下方的提示文字，再到歌庫查詢看編號。

## Driving it with AppSession

現有測試：`轉換為6位數_renumbers_every_song_keeps_我的最愛_and_saves_the_digit_count`、`轉換為6位數_counts_a_song_whose_語系_is_未知_as_failed_and_converts_the_rest`、`轉換為6位數_works_on_a_5_digit_歌庫_whose_設定檔_says_6_digits`、`校正編碼位數_counts_a_song_whose_語系_is_未知_as_failed_and_corrects_the_rest`。

```csharp
using (var app = new AppSession("to6"))
{
    app.WriteSettings(5);
    app.InsertSong("10001", "國語", "周杰倫", "晴天");
    app.InsertSong("20005", "台語", "江蕙", "家後");
    app.InsertFavorite("0001", "20005");
    app.Start();
    app.Click("轉換為 6 位數", onTab: "歌庫維護");
    app.Answer("確認提示", "是(Y)");
    AppSession.WaitUntil(() => app.Songs().All(s => s.Id.Length == 6), 60, "轉換為 6 位數 to finish");
    AppSession.WaitUntil(() => app.Setting("SongMgrMaxDigitCode") == "2", 20, "the 設定檔");
}
```

證明它能用的結果：

- `Songs()`：`10001 → 100001`、`20005 → 200001`（每個語系從起始編號重新編）。
- `Favorites()`：`0001:200001`，跟著新編號。
- 設定檔：`SongMgrMaxDigitCode=2`、`SongMgrLangCode` 等於 `AppSession.Code6`。
- 有「未知」語系時：提示文字的失敗數正確，那首歌的編號沒變。
- 證據：`click-轉換為 6 位數.png`（按下前的維護雜項）、`end.png`、證據資料夾裡的 `CrazyKTV_SongMgr.cfg`。

## Gotchas

- 5、6 位數混在一起的歌庫，在校正編碼位數之前無法正常使用，所以「先加歌再校正」的流程不驅動，直接用 `InsertSong` 佈置混合歌庫。
- 設定檔是程式結束前或背景寫入的，判定前用 `WaitUntil` 等 `Setting(...)`，不要直接讀。
- 啟動時 `SongDBUpdate_UpdateFinish` 在背景執行緒跑，改位數設定是 `BeginInvoke` 排進 UI 執行緒的，同一段裡讀 `Global.SongMgrMaxDigitCode` 會讀到舊值。要判斷位數就用歌庫算出來的 `MaxDigitCode`。

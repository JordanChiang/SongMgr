# 資料庫結構更新

程式啟動時檢查 `CrazySong.mdb` 的結構（`SchemaUpgrade.Check`），舊版歌庫缺欄位或多了舊欄位時，先備份到 `SongMgr\Backup\<時間>_CrazySong.mdb`，再自動更新結構，不問使用者、不掉資料。歌庫維護 › 資料庫結構 分頁顯示目前狀態。

## Sub-features

- 啟動時自動更新：加回缺的欄位（例如 `Song_ReplayGain`）、移除舊欄位（例如 `Song_MeanVolume`）。
- 更新前備份一份歌庫。
- 歌曲和我的最愛原封不動。
- 歌庫維護 › 資料庫結構 分頁：`SongMaintenance_DBVer1Value_Label` 顯示「最新」。
- 已是最新結構的歌庫啟動時不跳任何對話框（以前會跳「偵測到歌庫版本更新」）。

## How to get to it (user POV)

1. 把舊版程式用過的 `CrazySong.mdb` 放在程式旁邊，開程式。
2. 不用做任何事；到「歌庫維護」›「資料庫結構」看狀態是「最新」。
3. `SongMgr\Backup\` 裡多一份備份。

## Driving it with AppSession

現有測試：`An_old_歌庫_gets_its_資料庫結構更新_when_SongMgr_starts`。

```csharp
using (var app = new AppSession("schema"))
{
    app.WriteSettings(6);
    app.Execute("alter table ktv_Song drop column Song_ReplayGain");      // 做出舊版結構
    app.Execute("alter table ktv_Song add column Song_MeanVolume DOUBLE");
    app.InsertSong("100001", "國語", "周杰倫", "晴天");
    app.InsertFavorite("0001", "100001");
    app.Start();
    AppSession.WaitUntil(() => CrazyKTV_SongMgr.SchemaUpgrade.Check(app.Database).Count == 0, 60, "the 資料庫結構更新");
    AppSession.WaitUntil(() => app.Value("SongMaintenance_DBVer1Value_Label") == "最新", 20, "資料庫結構: 最新");
}
```

證明它能用的結果：`SchemaUpgrade.Check` 回空清單（這是讀檔判定，不是操作）、標籤是「最新」、`Dialogs()` 是空的、`Songs()` 和 `Favorites()` 沒變、`SongMgr\Backup\` 剛好一個 `*_CrazySong.mdb`（證據資料夾的 `Backup\` 也有）。

## Gotchas

- 舊版結構是用 `alter table` 在空白歌庫上做出來的，只涵蓋這兩個欄位差異；`SchemaUpgrade` 新增別的升級步驟時，佈置也要做出對應的舊結構，否則測試會 green 但沒驗到新步驟。
- 要用真的舊歌庫驗證，用第 5 層新舊版比對（專案根目錄的 `CrazySong.mdb`），不要把真實歌庫放進沙盒以外的地方。

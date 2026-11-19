# 套用錢櫃編號

錢櫃資料分頁的「套用錢櫃編號」：把歌庫裡在錢櫃資料（`SongMgr\CrazySongMgr.mdb`）找得到的歌，改成錢櫃的歌曲編號。目標編號已被別首占用時，被占的歌會連鎖搬走（也在錢櫃裡就搬到它的錢櫃編號，不在就換一個該語系的新編號），我的最愛跟著換。

## Sub-features

- 歌名＋歌手比對到錢櫃資料就改用錢櫃編號。
- 連鎖搬移：A 要去的編號上有 B，B 先搬到它自己的錢櫃編號。
- 讓位：占著錢櫃編號、但不在錢櫃的歌，換成語系的下一個新編號。
- 我的最愛跟著新編號。
- 錢櫃資料的查詢／其它查詢子分頁（`Cashbox_Query_TabPage`、`Cashbox_OtherQuery_TabPage`）：目前沒有端對端測試。

## How to get to it (user POV)

1. 按鈕預設是隱藏的：程式設定裡取消「隱藏套用錢櫃編號按鈕」（設定 `MainCfgHideApplyCashboxIdButton=False`）。
2. 切到「錢櫃資料」分頁，按「套用錢櫃編號」。
3. 「確認提示」按「是(Y)」，等完成，到歌庫查詢看編號。

## Driving it with AppSession

現有測試：`套用錢櫃編號_moves_songs_in_a_chain_makes_room_and_keeps_我的最愛`。

```csharp
using (var app = new AppSession("cashbox"))
{
    app.WriteSettings(6, ("MainCfgHideApplyCashboxIdButton", "False"));
    app.InsertSong("100001", "國語", "派偉俊", "去年下的那場雪"); // 錢櫃 002034
    app.InsertSong("100002", "國語", "派偉俊", "別戀");           // 錢櫃 002035，上面有 Want Me Back
    app.InsertSong("002035", "國語", "郁欣Ivy", "Want Me Back");  // 錢櫃 002036
    app.InsertSong("002034", "國語", "歌手", "不在錢櫃的歌");      // 要讓位
    app.InsertFavorite("0001", "100001");
    app.Start();
    app.Click("套用錢櫃編號", onTab: "錢櫃資料");
    app.Answer("確認提示", "是(Y)");
    AppSession.WaitUntil(() => app.Songs().Any(s => s.Name == "去年下的那場雪" && s.Id == "002034"), 60, "套用錢櫃編號");
}
```

證明它能用的結果：`Songs()` 是 `002034 去年下的那場雪`、`002035 別戀`、`002036 Want Me Back`、`100003 不在錢櫃的歌`；`Favorites()` 指向新編號；沒有產生 `SongMgr\Backup\Favorite.txt`。

## Gotchas

- 範例用的 002034～002036 是隨程式附的 `CrazySongMgr.mdb` 裡的真實錢櫃資料；更新錢櫃資料後這些歌可能換號，測試會 red，要先查新資料再改佈置。
- 沒有設 `MainCfgHideApplyCashboxIdButton=False` 時找不到按鈕，`Click` 會在 30 秒後逾時，訊息裡的 `On screen:` 列出看得到的按鈕。

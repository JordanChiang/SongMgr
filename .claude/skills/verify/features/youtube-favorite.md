# YouTube 我的最愛

UltimateKtv 把 YouTube 影片的我的最愛存在 `YoutubeFavorite`（User_Id + VideoId，自帶名稱、長度、加入日期），不放 `ktv_Favorite`。它的「今日點唱」（`****`）和每次開機的點唱紀錄（`MM/dd HH:mm`）也記在同一張表。加歌程式讀寫這張表：

- 歌庫查詢 › 類型「YouTube 歌曲」：右鍵「加入我的最愛 › 用戶」。
- 歌庫查詢 › 其它查詢 › 最愛：選用戶後先列本機歌曲，再列他的 YouTube 最愛（語系顯示 `YouTube`）。
- 我的最愛清單裡的 YouTube 列：右鍵「在瀏覽器開啟」「複製影片網址」「從我的最愛移除」，雙擊開瀏覽器。
- 歌庫維護 › 我的最愛：匯出、匯入（`YoutubeFavorite|…` 行）、移除最愛用戶（連 YouTube 最愛一起刪）。

## Sub-features

- 加入：資料庫沒有這張表時照 UltimateKtv 的定義建立；已在清單、沒有影片 ID 的略過，狀態列寫出略過幾筆。
- 列出：依加入時間排序；已下載的影片帶出點播次數、音量、大小；歌手類別、聲道留白。
- 移除：只拿掉選取的影片，本機歌和其他影片不動；多選時本機歌與影片可一起移除。
- 匯出／匯入：名稱放最後（可能含 `|`）；帶 YouTube 行的備份整個換掉 YouTube 最愛，舊格式備份保留備份裡用戶的 YouTube 最愛。
- 錢櫃排行等清單更新（`ReplaceFavoriteUser`）不動 YouTube 最愛。

## How to get to it (user POV)

1. 歌庫查詢 › 類型選「YouTube 歌曲」 › 按「查詢」 › 在列上按右鍵 › 加入我的最愛 › 用戶。
2. 歌庫查詢 › 「其它查詢」子分頁 › 最愛下拉選用戶。
3. 在 YouTube 列上按右鍵。
4. 歌庫維護 › 「我的最愛」子分頁 › 匯出我的最愛／匯入我的最愛（確認提示「是」）／在清單選用戶按「移除」（確認提示「是」）。

## Driving it with AppSession

現有測試：`YoutubeFavoriteEndToEndTests`（篩選 `Category=EndToEnd&FullyQualifiedName~YoutubeFavoriteEndToEndTests`）。

```csharp
app.Choose("SongQuery_QueryType_ComboBox", "YouTube 歌曲");
app.Click("查詢", onTab: "歌庫查詢");
app.ClickRow("SongQuery_DataGridView", "Y0001");
app.RightClickRow("SongQuery_DataGridView", "Y0001");
app.ClickMenu("加入我的最愛", "小明");

app.SelectSubTab("SongQuery_TabControl", "其它查詢", "歌曲查詢", "其它查詢");
app.Focus("SongQuery_FavoriteQuery_ComboBox");
app.Choose("SongQuery_FavoriteQuery_ComboBox", "小明");

app.GoToTab("歌庫維護");
app.SelectSubTab("SongMaintenance_TabControl", "我的最愛", "維護雜項", "我的最愛", "自訂語系", "FFmpeg", "資料庫結構");
app.SelectListItem("SongMaintenance_Favorite_ListBox", "小明");
app.Click("移除", onTab: "歌庫維護");
```

證明它能用的結果：讀沙盒歌庫的 `YoutubeFavorite`（User_Id、VideoId、Name、Length、CreatDate）、`ktv_Favorite`（`app.Favorites()`），以及 `SongMgr\Backup\Favorite.txt` 的內容；`SongQuery_QueryStatus_Label` 的筆數文字；截圖 `favorite-list.png`、`menu-從我的最愛移除.png`、`end.png`。

## Gotchas

- 最愛下拉只在使用者操作（控制項有焦點）且選到不同的項目時才查詢：先 `Focus`，而且要從另一個用戶切過去。只有一個用戶時下拉一開始就停在他身上，選他不會查詢，所以測試放兩個用戶。
- 子分頁（歌庫查詢的歌曲查詢／其它查詢、歌庫維護的維護雜項／我的最愛…）收不到 Ctrl+PageDown，UI Automation 也只看得到選中的那個標籤，`SelectSubTab` 沿標籤列從左點過去。
- 右鍵選單在 UI Automation 裡掛在主視窗底下；`ClickMenu` 用滑鼠點，不用 Invoke（「從我的最愛移除」會開 MessageBox，Invoke 會卡住）。
- 未自動驗證：雙擊 YouTube 列、「在瀏覽器開啟」（會開真的瀏覽器）。

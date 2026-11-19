# 歌庫查詢與編輯

歌庫查詢分頁列出歌庫裡的歌（`SongQuery_DataGridView`），可以依條件查詢；按「編輯模式」後點一列，右邊的編輯區顯示這首歌的資料，改完按「套用變更(A)」寫回 `ktv_Song`。改語系時歌曲編號會換成新語系的下一個編號。

## Sub-features

- 歌曲查詢：`SongQuery_QueryType_ComboBox` 選查詢類型、`SongQuery_QueryValue_TextBox` 輸入、按「查詢」。
- 其它查詢分頁（`SongQuery_OtherQuery_TabPage`）：異常與重複歌曲查詢。
- 編輯模式：點列後編輯區（`SongQuery_EditSongId_TextBox`、`SongQuery_EditSongLang_ComboBox` 等）顯示該歌資料。
- 改語系 → 歌曲編號換成新語系的下一個編號，套用後寫入資料庫。

## How to get to it (user POV)

1. 開程式，預設就在「歌庫查詢」分頁，歌庫的歌會列出來。
2. 按「編輯模式」，狀態列（`SongQuery_QueryStatus_Label`）顯示「已進入編輯模式...」。
3. 點要改的歌，編輯區出現它的資料；改語系下拉，歌曲編號欄跟著變。
4. 按「套用變更(A)」。

## Driving it with AppSession

現有測試：`SongIdEndToEndTests.歌庫監視加歌_then_歌庫查詢_moves_a_song_to_another_語系` 的後半段。

```csharp
app.Start();
app.Click("編輯模式");
AppSession.WaitUntil(() => app.Value("SongQuery_QueryStatus_Label") == "已進入編輯模式...", 20, "編輯模式");
app.ClickRow("SongQuery_DataGridView", "100003");             // 第一欄（歌曲編號）是 100003 的列
AppSession.WaitUntil(() => app.Value("SongQuery_EditSongId_TextBox") == "100003", 20, "the edit panel");
app.Choose("SongQuery_EditSongLang_ComboBox", "台語");
AppSession.WaitUntil(() => app.Value("SongQuery_EditSongId_TextBox") == "200004", 20, "the next 台語 number");
app.Click("套用變更(A)");
AppSession.WaitUntil(() => app.Songs().Any(s => s.Name == "日不落" && s.Lang == "台語"), 30, "saved");
```

證明它能用的結果：資料庫裡那首歌的語系和編號都變了、編號沒有和別首重複（`Songs().Select(s => s.Id).Distinct().Count()` 等於歌數）；截圖 `click-套用變更(A).png` 看得到編輯區的新值。

## Gotchas

- 「編輯模式」的按鈕是 post 出去的：要先等狀態列變成「已進入編輯模式...」，否則版面還沒縮，`ClickRow` 會點到舊位置（commit `edbb0f7` 修過這個）。
- `ClickRow` 是真的移動滑鼠點擊（DataGridView 的列沒辦法用 UI Automation 選），期間動到滑鼠會點錯。
- 查詢（「查詢」按鈕）和其它查詢目前沒有端對端測試；改到 `SongSearch` 等查詢邏輯時，要補一段：選查詢類型、輸入值、`Click("查詢")`、再讀 `SongQuery_DataGridView` 的列。`AppSession` 目前沒有讀整個表格的方法，需要時照 `ClickRow` 的走法加一個。

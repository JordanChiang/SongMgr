---
name: verify
description: 驗證 CrazyKTV 加歌程式（CrazyKTV_SongMgr，WinForms 桌面程式）的修改：建置、單元測試、用 UI Automation 開出真正的主視窗操作並保留截圖與資料庫證據、新舊版比對。改完程式準備說完成、commit 或開 PR 前，要證明某個功能在真的程式裡能用，或使用者要求驗證、跑測試時使用。
---

# 驗證修改

由下往上一層一層跑，每層 **green** 才往下一層。任何一層 **red** 就停下來修，修完從第 1 層重跑。指令都在專案根目錄（或目前的 worktree 根目錄）執行。

程式和 Jet OLEDB 4.0 只有 32 位元，所以測試一律帶 `-s SongMgr.Tests\x86.runsettings`，在 x86 testhost 裡跑。

本 skill 的腳本在 `.claude\skills\verify\scripts\`，跟著版控走，每個 worktree 都有自己的一份。在要驗證的 checkout（或 worktree）根目錄呼叫，腳本會用 `git rev-parse --show-toplevel` 找到所在的 checkout，驅動的是那一份的建置結果。

## 程式長什麼樣子

- **使用者碰到的是**：一個主視窗 `CrazyKTV 加歌程式`，`MainTabControl` 底下七個固定分頁：歌庫查詢、加歌頁面、歌手管理、歌庫設定、歌庫維護、錢櫃資料、程式設定（另有加歌結果、操作記錄等會被隱藏的分頁）。資料在 exe 旁邊的 `CrazySong.mdb`（Access，Jet 4.0），設定在 `CrazyKTV_SongMgr.cfg`。
- **驅動方式**：`SongMgr.Tests\EndToEnd\AppSession.cs`。它把建置好的 `CrazyKTV_SongMgr.exe` 複製到 `%TEMP%\SongMgrTests\e2e_<名稱>_<亂碼>\app\` 沙盒，放一份空白歌庫（`Fixtures\CrazySongEmpty.mdb`）和自己寫的設定檔（自動更新關閉，不上網），開出主視窗，用 UI Automation 和 Win32 訊息點擊，再讀沙盒裡的資料庫和設定檔判定結果。功能地圖在 [features/README.md](features/README.md)。
- **隔離**：每個 session 有自己的沙盒資料夾，資料不會互相干擾，也不碰專案根目錄的真實歌庫。但它會移動滑鼠、送鍵盤，**同一時間只能有一個驅動**：另一個測試執行或新舊版比對在跑時不要啟動，`doctor.ps1` 會擋下來。使用者自己開著的加歌程式不用關，session 只操作自己啟動的那個 process。

## Doctor：能不能驅動

任何東西看起來不對勁、或第 4 層之前，先跑這個（唯讀，不改任何東西）：

```powershell
powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\doctor.ps1
```

每項一行 `OK` / `WARN` / `FAIL`，最後一行 `READY`（exit 0）或 `NOT READY`（exit 1）。檢查 x86 dotnet、Jet OLEDB 4.0、互動桌面且螢幕沒鎖、建置輸出是否存在且比原始碼新、沙盒需要的 Fixtures 和 ffmpeg、有沒有其他沙盒程式或 `SongMgr.Tests` 測試正在跑。

- `FAIL a SongMgr.Tests test run is in progress`：別人（使用者或另一個 session）在跑測試或新舊版比對，它也會鎖住 `bin\` 裡的 dll，連建置都會失敗。**不要砍它**，等它結束，或問使用者。
- `FAIL a sandboxed SongMgr is already running` 而沒有測試在跑：上一次留下的，跑 Cleanup。
- `WARN build is stale`：先跑第 1 層。

## 1. 建置

```powershell
dotnet build SongMgr.Tests\SongMgr.Tests.csproj -c Debug -v q -nologo
```

測試專案參考主程式，這一行會一起建置 `CrazyKTV_SongMgr`，輸出到 `SongMgr.Tests\bin\Debug\net8.0-windows7.0\`（沙盒從這裡複製）。

**完成條件**：0 個錯誤、0 個警告。目前的程式碼是 0 警告，出現的警告就是這次修改帶進來的，要處理掉。`MSB3027 ... 檔案鎖定者: "testhost.x86"` 表示有測試正在跑，見 Doctor。

## 2. 單元測試

```powershell
$env:SONGMGR_SKIP_E2E = '1'; dotnet test SongMgr.Tests\SongMgr.Tests.csproj --no-build -v q --nologo -s SongMgr.Tests\x86.runsettings
```

約 2 分鐘，超過工具預設的逾時，設長一點的 timeout 或在背景執行。測試不平行執行（`TestEnv` 關掉了平行），因為它們共用靜態的 `Global`。

**完成條件**：失敗 0；略過的只有 `EndToEnd` 和 `Compare` 兩個類別（端對端被環境變數關掉，比對只在比對腳本裡跑）。其他測試被略過要查原因。

## 3. 補上這次修改的測試

列出這次改到的每個行為，逐一指出哪個測試涵蓋它：拿掉修改後那個測試會 red。沒有測試涵蓋的行為，補一個（測試先寫、先看到 red，見 `tdd` skill），然後回到第 1 層。

測試放在 `SongMgr.Tests\` 底下對應領域的資料夾。要資料庫的測試用 `TestEnv.UseScratchDatabases` 取得空白歌庫副本；要真的操作主視窗的，用 `EndToEnd\AppSession`，寫法見下面的 Drive 和功能地圖。

**完成條件**：每個改到的行為都有對應的測試；真的無法自動測試的，在回報裡明列為「未驗證」。

## 4. 端對端：開出真的程式操作（Launch → Drive → Evidence → Cleanup）

修改碰到主視窗操作、加歌或歌庫監視、歌曲編號、設定檔讀寫、啟動時的資料庫結構更新時要跑。跑之前告訴使用者：畫面上會跳出加歌程式的視窗，期間別動滑鼠鍵盤。

### Launch + Drive：一行跑完

```powershell
powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\e2e.ps1
```

依序：建置 → doctor（不 READY 就 exit 2，什麼都不開）→ 設好 `SONGMGR_E2E_EVIDENCE` 跑 `Category=EndToEnd` → cleanup → 列出證據。全部約 1 到 2 分鐘，用背景執行或 timeout 設 600000。

- 只跑一部分：`-Filter "Category=EndToEnd&FullyQualifiedName~轉換為6位數"`（xUnit 的篩選語法，條件見功能地圖各檔）。
- 已經建置過：加 `-NoBuild`。
- exit code 就是測試的結果，0 = 全部通過。

每個 `AppSession` 的「就緒」是 `Start()` 等到 process 的頂層視窗標題以 `CrazyKTV` 開頭（最多 60 秒；沒有 VMR9 的機器會先跳「顯示環境檢查」，它會自動按確定）。主視窗出現不代表啟動完成：程式啟動時會自己切分頁、跑歌庫監視和資料庫結構更新，所以測試一律用 `WaitUntil` 等**資料庫或標籤的結果**，不要用 sleep。

### 寫一段新的操作

新功能要證明時，在 `SongMgr.Tests\EndToEnd\` 加一個 `[Trait("Category", "EndToEnd")]` 類別、`[DesktopFact]` 方法，照這個形狀：

```csharp
using (var app = new AppSession("短名稱"))           // 建沙盒
{
    app.WriteSettings(6, ("設定名", "值"));            // 設定檔；5 或 6 位數歌曲編號
    app.InsertSong("100001", "國語", "周杰倫", "晴天"); // 事先放進歌庫的資料
    app.Start();                                       // 開程式，等主視窗
    app.Click("轉換為 6 位數", onTab: "歌庫維護");      // 按鈕/分頁/核取方塊的文字
    app.Answer("確認提示", "是(Y)");                   // MessageBox 標題 + 按鈕
    AppSession.WaitUntil(() => app.Songs().All(s => s.Id.Length == 6), 60, "轉換完成");
    Assert.Equal(..., app.Songs());                   // 判定看資料庫/設定檔，不只看畫面
}
```

可用的操作（都在 `AppSession`）：`Click(文字, onTab)`、`Answer(標題, 按鈕)`、`Dialogs()`、`Value(AutomationId)` 讀標籤/文字框、`Choose(下拉 Id, 項目)`、`ClickRow(表格 Id, 第一欄的值)`、`Snap(標籤)` 額外截圖；資料面有 `WriteSettings`、`WatchMonitorFolder`、`MakeSongFile`（用 ffmpeg 做 3 秒的歌曲檔）、`InsertSong`、`InsertFavorite`、`Execute`、`Songs()`、`Favorites()`、`Setting(名稱)`。

控制項的 AutomationId 就是 `MainForm.Designer.cs` 裡的 `Name`（例如 `SongQuery_DataGridView`、`SongMaintenance_Tooltip_Label`）；`Click` 用的是畫面上的文字（例如 `套用變更(A)`，`&` 不算）。用這兩種，不要用座標或 Tab 順序。分頁只能用 `onTab:` 切換七個固定分頁（見 `AppSession.MainTabs`）。

### Evidence：證據在哪、要拿什麼當證明

`e2e.ps1` 結束時印出 `EVIDENCE <資料夾>`，預設 `%TEMP%\SongMgrEvidence\<yyyyMMdd-HHmmss>\`，**cleanup 不會刪它**：

- `e2e.trx`、`test-output.txt`：每個測試的結果與失敗訊息。
- `e2e_<名稱>_<亂碼>\`：每個 session 一個資料夾。`NN-click-<按鈕>.png` 是每次按下前的主視窗，`NN-end.png` 是結束時的主視窗，`CrazySong.mdb`、`CrazyKTV_SongMgr.cfg` 是程式結束後的歌庫和設定檔，`Backup\` 是程式做的備份。
- 失敗時另外印 `FAILURE SCREENSHOT %TEMP%\SongMgrTests\e2e-<沙盒>.png`，失敗訊息裡有當時的分頁移動和畫面上的按鈕。用 Read 開截圖看卡在哪個畫面。

證明的標準：

- 走使用者的路：點按鈕、選下拉、回答對話框。不要在測試裡呼叫程式內部的方法或改 `Global` 來「操作」，那只能當前置資料（`InsertSong` 直接寫資料庫是佈置現場，不是操作）。
- 要有動作和結果：截圖要看得到按下前的畫面和結束時的畫面，判定要讀資料庫或設定檔的實際內容（歌曲編號、我的最愛、`SongMgrMaxDigitCode`），不是只看到「沒有錯誤」。
- 要查副作用：寫了哪些列、設定檔哪一行變了、備份檔有沒有產生、不該出現的檔案（例如 `Favorite.txt`）沒出現。
- 不用 mock。沙盒裡就是真的 exe、真的 Jet、真的 ffmpeg；唯一隔開的是自動更新（設定檔關掉，不上網）。

回報時附上證據資料夾的完整路徑，至少用 Read 看過一張 `end.png` 確認畫面合理。

### Cleanup

`e2e.ps1` 結束時一定會跑。自己手動跑測試、或測試被中斷後：

```powershell
powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\cleanup.ps1
```

只停止執行檔位於 `%TEMP%\SongMgrTests\e2e_*` 底下的 process（沙盒裡的加歌程式和它叫的 ffmpeg），只刪 `e2e_*` 沙盒資料夾；有 `SongMgr.Tests` 測試還在跑時整個跳過（exit 1），免得刪掉別人正在用的沙盒。**不用名稱砍 process**（使用者可能自己開著加歌程式），也不刪 `%TEMP%\SongMgrEvidence` 和失敗截圖。

**完成條件**：失敗 0、略過 0（略過代表沒有互動桌面或設了 `SONGMGR_SKIP_E2E`，不算驗證過）。

## 5. 新舊版比對

大範圍重構，或發佈前才跑。用專案根目錄的真實歌庫 `CrazySong.mdb` 和 `CrazyKTV_SongMgr.cfg`（未納入版控，只用副本），比較目前的程式和已發佈版本（預設 `2.11.4`）使用者看到的結果。

約 15 到 20 分鐘且佔用桌面，**先問使用者**再跑，並在背景執行：

```powershell
powershell -ExecutionPolicy Bypass -File tools\compare-with-release.ps1
```

報告在 `%TEMP%\SongMgrCompare\report.txt`，第一行是相同的段落數。參數與比較範圍見 `tools\README.md`。跑的期間 doctor 會回報有測試在跑，不要同時跑第 4 層。

**完成條件**：報告列出的每一個差異都判定為刻意的修改（說得出是哪個 commit 改的）或回歸；回歸要修掉再重跑。

## 回報

每一層貼出實際的結果行（例如「失敗: 0，通過: 480，略過: 10」）。第 4 層附上證據資料夾路徑。沒跑的層寫明沒跑及原因，第 3 層的「未驗證」項目逐條列出。

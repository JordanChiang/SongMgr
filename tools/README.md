# 發佈工具

## pack-release.ps1

產生 GitHub Release 要上傳的附檔。程式的「啟用程式自動更新」會讀這些檔案，只下載有變動的部分。

### 元件

元件只定義在 `pack-release.ps1` 的 `$components`，程式端依 `update.json` 處理，不寫死元件名稱。要新增或改名元件，只改腳本即可。

| 元件 | 內容 | 多久變一次 |
|---|---|---|
| `App` | 根目錄的檔案（`CrazyKTV_SongMgr.exe`、`SQLite.Interop.dll`，不含 `.pdb`）、`SongMgr\` | 每次發佈 |
| `Codec_x86` | 影音解碼器 | 很少，換 LAV 版本時才會變 |
| `Tools` | ffmpeg / ffplay / ffprobe | 很少 |

### 自動發佈 (GitHub Actions)

`.github/workflows/release.yml` 會在推送版本號 tag 時自動發佈：

1. 修改 `CrazyKTV_SongMgr\Properties\AssemblyInfo.cs` 的 `AssemblyFileVersion` (例如 `2.12.0`) 並 commit 到 master。
2. 推送同名 tag：`git tag 2.12.0 && git push origin 2.12.0`。

workflow 會檢查 tag 與 `AssemblyFileVersion` 一致、`dotnet publish` (FolderProfile)、補上 `Codec_x86\` (發佈時 PostBuild 不會複製)、驗證產出、執行 `pack-release.ps1`、建立 Release 並上傳附檔，最後把 `CrazyKTV_SongMgr.exe` 送 VirusTotal，結果附加到 Release notes。

- 手動觸發 (Actions → Release → Run workflow) 可勾 `dry_run`：照常建置、打包、送 VirusTotal，但不建立 Release；附檔上傳為 workflow artifact，VT 結果寫在該 run 的 Summary。
- VirusTotal 需要 repo secret `VT_API_KEY`，沒設定就略過掃描，不影響發佈。VT 結果對照 `vt-baseline.json` 白名單，出現白名單外的新命中才建議送 Hybrid Analysis。
- UltimateKtvPublic 的發佈會抓本 repo 最新 release 的 `SongMgr_App_*.zip`，所以本程式要先發佈，UltimateKtv 才會帶到新版。

### 手動發佈步驟

1. 修改 `CrazyKTV_SongMgr\Properties\AssemblyInfo.cs` 裡的 `AssemblyFileVersion`，例如 `2.12.0`。
2. 在 Visual Studio 用 `FolderProfile` 發佈，輸出到 `CrazyKTV_SongMgr\Publish`。
3. 在專案根目錄執行：

   ```bash
   powershell -ExecutionPolicy Bypass -File tools\pack-release.ps1
   ```

4. 在 GitHub 建立 Release：
   - tag 要和 exe 的版本號**完全一樣**，例如 `2.12.0`
   - 不要勾選 pre-release，否則自動更新會略過這個版本
   - 上傳輸出資料夾 `CrazyKTV_SongMgr\Publish\ReleaseAssets\<版本>\` 裡的**所有檔案**

### 要上傳哪些檔案

腳本會讀最新 release 的 `update.json`。跟上一版比沒有變動的元件不會重新打包，清單直接指向舊 release 的附檔。所以輸出資料夾裡有什麼，就上傳什麼。

**第一次發佈**：5 個檔案都要上傳。

| 檔案 | 用途 | 必要 |
|---|---|---|
| `update.json` | 自動更新讀的清單 | **一定要** |
| `SongMgr_App_<版本>.zip` | 主程式 | 要 |
| `Codec_x86_<版本>.zip` | 影音解碼器 | 要 |
| `Tools_<版本>.zip` | FFmpeg 工具 | 要 |
| `CrazyKtv_SongMgr<版本>.zip` | 完整包，給手動下載及第一次安裝；沒有 `update.json` 時自動更新也只認這個檔名 | **一定要** |

**之後的版本**：通常只有 3 個，例如只改了程式的 2.12.1：

```
update.json
SongMgr_App_2.12.1.zip
CrazyKtv_SongMgr2.12.1.zip
```

更新 LAV 或 ffmpeg 時，才會多出 `Codec_x86_*.zip` 或 `Tools_*.zip`。

### 注意事項

- **每次都要上傳 `update.json` 和完整包。** 沒有 `update.json` 時，自動更新會改成下載完整包；`update.json` 讀不到或格式錯誤時，這次不更新，下次啟動再檢查。
- **不要刪除舊 release 和它的附檔。** 新的 `update.json` 可能還指向舊 release 裡的 `Tools_*.zip` 或 `Codec_x86_*.zip`。
- **tag 和版本號要一致。** 清單裡的下載網址是用 tag 組出來的，不一致就會下載失敗。
- 在 `Publish` 資料夾執行過程式留下的 `.cfg`、`.log` 不會被打包。
- 使用者電腦上已經存在的參考資料庫、自訂語系和設定檔，自動更新不會覆蓋。清單在 `AppUpdater.KeepExisting`。

### 避免防毒誤判

發佈的 exe 曾在 VirusTotal 被 ML 引擎 (APEX、Cylance) 判為惡意、Hybrid Analysis 49 分。以下設定是為了避免誤判，不要改回去：

- **原生 DLL 不包進 exe** (`FolderProfile.pubxml` 的 `IncludeNativeLibrariesForSelfExtract=false`)，`SQLite.Interop.dll` 放在 exe 旁邊。這是影響最大的一項：包進去時 exe 會符合 ChaCha20 加密演算法的 YARA 規則 (SQLite 的亂數產生器)，執行時還會解壓 DLL 到暫存資料夾再載入。實測其他修正都做了、只把 DLL 包回 exe，VirusTotal 仍有 Cylance、Hybrid Analysis 仍是 suspicious 45 分。
- **程式圖示每個尺寸都是未壓縮的 32bpp BMP。** PNG 壓縮的圖示會讓 exe 資源區段的熵接近亂數，Hybrid Analysis 會標「資源熵異常」。換圖示時用 `python tools/ico-png-to-bmp.py <原.ico> <新.ico>` 重新封裝，像素不變。
- 程式碼裡不要出現 `regedit.exe` 字串 (會被標為「檢查鑑識工具」)，匯入 `.reg` 用 `reg.exe import`。

修正後 VirusTotal 剩 APEX 1 個 (拿掉 256x256 圖示也一樣)，Hybrid Analysis 為 no specific threat。發佈後可把 exe 上傳到 [VirusTotal](https://www.virustotal.com/) 和 [Hybrid Analysis](https://hybrid-analysis.com/) 檢查。

### 參數

| 參數 | 說明 | 預設值 |
|---|---|---|
| `-PublishDir` | 發佈資料夾 | `CrazyKTV_SongMgr\Publish` |
| `-OutDir` | 輸出資料夾 | `<PublishDir>\ReleaseAssets\<版本>` |
| `-Tag` | Release 的 tag | exe 的檔案版本 |
| `-PreviousManifest` | 上一版 `update.json` 的路徑或網址 | 最新 release 的 `update.json` |
| `-NoReuse` | 不沿用上一版的元件，全部重新打包 | 不啟用 |

例如要以本機的舊清單比對，並強制指定 tag：

```bash
powershell -ExecutionPolicy Bypass -File tools\pack-release.ps1 -Tag 2.12.0 -PreviousManifest D:\old\update.json
```

### update.json 格式

```json
{
  "version": "2.12.0",
  "tag": "2.12.0",
  "components": [
    {
      "name": "App",
      "desc": "加歌程式",
      "url": "https://github.com/JordanChiang/SongMgr/releases/download/2.12.0/SongMgr_App_2.12.0.zip",
      "size": 3405285,
      "files": {
        "CrazyKTV_SongMgr.exe": "<sha256>",
        "SongMgr/CrazySongMgr.mdb": "<sha256>"
      }
    }
  ]
}
```

自動更新會逐一比對 `files` 裡每個檔案的 SHA256。某個元件有任何檔案不存在或內容不同，才會下載那個元件的 zip；解壓後也會用同一份雜湊值驗證，不符就中止更新。

### 覆蓋檔案

下載並驗證完成後，程式先把會被取代的檔案備份到暫存資料夾，再關閉自己，由批次檔覆蓋：

- 程式有背景工作（例如加歌）時，會等工作完成才詢問更新或關閉程式。
- 批次檔最多等程式結束 10 分鐘（關閉時會壓縮資料庫）。逾時仍未結束就放棄這次更新，不動任何檔案。
- 覆蓋失敗（robocopy 結束碼 8 以上）時，把備份還原回程式資料夾，再以 `--update-failed` 啟動舊版並提示使用者。
- 還原也失敗時保留備份（`%TEMP%\CrazyKTV_SongMgr_Update\backup`），提示使用者手動複製回去。

## compare-with-release.ps1

新舊版比對：用真實歌庫比較目前的程式和已發佈的版本，使用者看到的結果是否相同。發佈前、或大範圍重構後跑一次。

```bash
powershell -ExecutionPolicy Bypass -File tools\compare-with-release.ps1
```

- 預設比較 `2.11.4`，用專案根目錄的 `CrazySong.mdb` 和 `CrazyKTV_SongMgr.cfg`；可用 `-Baseline`、`-Database`、`-Settings` 指定其他的。歌庫只用副本，原檔不會被修改。
- 已發佈的版本會取到暫存的 git worktree，跑完就移除（`-KeepWorktree` 可保留）。
- 兩版都在畫面外開啟真正的主視窗，對話框自動回答。需要互動桌面，全部約 15 到 20 分鐘。
- 比較的內容：
  - 歌庫查詢：一般查詢、異常查詢、我的最愛的清單和狀態列
  - 加歌：新歌、所有歌曲檔、新歌加上已在歌庫的歌（重複歌曲判斷）的歌曲分析清單、加歌結果、操作記錄，以及加歌後的 `ktv_Song`、`ktv_Singer`、`ktv_Favorite`
- 報告在 `%TEMP%\SongMgrCompare\report.txt`，第一行是相同的段落數。列出的差異要逐一判斷是刻意的修改，還是回歸。

程式在 `SongMgr.Tests\Compare\`：`CompareData` 準備資料，`LibraryCompare` 操作主視窗並記錄結果，`CompareReport` 寫報告。一般跑測試時這三個測試會略過。

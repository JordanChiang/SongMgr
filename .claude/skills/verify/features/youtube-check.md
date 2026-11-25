# 檢查影片（YouTube 歌曲）

歌庫查詢 › 類型「YouTube 歌曲」時，狀態列「編輯模式」左邊出現「檢查影片」。按下後先用 YouTube 的 oEmbed（約 0.1 秒一筆，4 筆同時）查清單上的影片：回 200 就是「可用」；沒回應、5xx 是「無法確認」；其他狀態（不存在的 400、私人／不可嵌入的 401/403/404、被限流的 429）再交給 UltimateKtv 的 yt-dlp（程式目錄 `Tools\yt-dlp.exe`，32 位元 Windows 用 `yt-dlp_x86.exe`，約 2.5 秒一筆，一次一筆）確認（只查不下載），把結果寫在「檢查結果」欄（名稱右邊）；再到 UltimateKtv `settings.json` 的 `YoutubeDownloadDir`（相對於程式目錄，讀不到用 `YoutubeDownloads`）找 YoutubeSongList 的 `FileName`。

| YouTube | 本機檔 | 檢查結果 | 處理 |
|---|---|---|---|
| 可用 | 有／無 | 可用／可用 (未下載) | 保留 |
| 已失效 | 有 | 已失效 (本機可播)（紅） | 保留 |
| 已失效 | 無 | 已失效 (無本機檔)（紅） | 「移除提示」按「是」才刪 YoutubeSongList 的列，不刪檔 |
| 其他錯誤、沒有影片 ID、逾時 60 秒 | — | 無法確認: …（灰） | 保留 |

「已失效」只認 yt-dlp 的這幾種錯誤：This video is unavailable、Video unavailable、Private video、has been removed、帳號已終止、ID 不完整。年齡限制、要登入、地區限制、網路錯誤都算「無法確認」。

## Sub-features

- 按鈕只在 YouTube 歌曲清單出現；狀態文字讓出位置。
- 檢查中按鈕變「取消檢查」，查完目前這筆就停；離開清單也會取消。
- 找不到 yt-dlp：狀態列提示，不檢查。
- 結束時狀態列：`檢查完成: 可用 a 筆,失效但本機可播 b 筆,失效無檔已移除 c 筆,無法確認 d 筆。`（按「否」是「失效無檔未移除」，取消是「已取消 (查了 n/N 筆)」開頭）。

## How to get to it (user POV)

1. 歌庫查詢 › 類型「YouTube 歌曲」 › 查詢。
2. 按「檢查影片」，等狀態列跑完。
3. 有可移除的會跳「移除提示」，列出名稱，按「是」移除。

## Driving it with AppSession

`YoutubeCheckEndToEndTests`（篩選 `Category=EndToEnd&FullyQualifiedName~YoutubeCheckEndToEndTests`）。**要網路，並用環境變數 `SONGMGR_YTDLP` 指向一個 yt-dlp.exe**（例如 UltimateKtv 的 `Tools\yt-dlp.exe`），沒設就略過——略過不算驗證過：

```powershell
$env:SONGMGR_YTDLP = 'D:\VS2022\UltimateKtvPublic\Tools\yt-dlp.exe'
powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\e2e.ps1 -Filter "Category=EndToEnd&FullyQualifiedName~YoutubeCheckEndToEndTests"
```

測試把 yt-dlp 複製到沙盒 `app\Tools\`，寫一份帶 `//` 註解的 `settings.json`（`YoutubeDownloadDir` = `影片快取`），放四筆：`dQw4w9WgXcQ`（可用）、`xxxxxxxxxxx`（不存在，但下載目錄有檔）、`zzzzzzzzzzz`（不存在、沒檔）、沒有影片 ID。

證明它能用的結果：「移除提示」出現前四筆都還在；按「是」後只剩 Y0001、Y0002、Y0004，`影片快取` 裡的檔還在；狀態列是上面的摘要。截圖 `checking-N.png`（檢查中）、`check-results.png`（檢查結果欄與移除提示）、`end.png`。

## Gotchas

- yt-dlp 不能加 `--ignore-no-formats-error`：它會把不存在的影片也當成功。
- 不存在的影片 yt-dlp 印「This video is unavailable」，UltimateKtv 的 PermanentErrorMarkers（"Video unavailable"）比對不到；這裡兩種都認。
- 單元測試 `YoutubeAvailabilityTests` 用 `.cmd` 假 yt-dlp 和假 oEmbed 測執行、逾時、判斷，以及只有 oEmbed 沒判定的才問 yt-dlp，不上網。
- 只靠 oEmbed 不能判「已失效」：不可嵌入、年齡限制的影片也回 401/403，所以非 200 一律再問 yt-dlp。

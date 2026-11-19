# 功能地圖

加歌程式使用者看得到的功能，每個一個檔案：功能是什麼、使用者怎麼走到、怎麼用 `AppSession` 驅動、什麼樣的結果才算證明它能用。這是本專案驗證用的正式來源：改到某個功能，證明要涵蓋它檔案裡列的每一個入口，只驅動其中一個方便的入口不算完整。

驅動方式、證據位置與清理見 [../SKILL.md](../SKILL.md) 第 4 層。下表「篩選」欄可直接給 `e2e.ps1 -Filter`。

| 功能 | 主視窗分頁 | 現有端對端測試 | 篩選 |
|---|---|---|---|
| [歌庫監視加歌](song-add-monitor.md) | 歌庫設定 › 監視目錄、加歌頁面 | 有 | `Category=EndToEnd&FullyQualifiedName~歌庫監視加歌` |
| [歌庫查詢與編輯](song-query-edit.md) | 歌庫查詢 | 有（只有改語系） | `Category=EndToEnd&FullyQualifiedName~歌庫監視加歌` |
| [歌曲編號維護](song-id-maintenance.md) | 歌庫維護 › 維護雜項 | 有 | `Category=EndToEnd&FullyQualifiedName~位數` |
| [套用錢櫃編號](cashbox-apply-id.md) | 錢櫃資料 | 有 | `Category=EndToEnd&FullyQualifiedName~套用錢櫃編號` |
| [資料庫結構更新](schema-upgrade.md) | 啟動時自動、歌庫維護 › 資料庫結構 | 有 | `Category=EndToEnd&FullyQualifiedName~資料庫結構更新` |

還沒有地圖的分頁：歌手管理、歌庫設定的其他子分頁（歌曲類別、歌手群組、歌庫結構、自訂結構）、歌庫維護的我的最愛／自訂語系／FFmpeg、程式設定、歌庫轉換、內建播放器。改到這些時先照同樣的形狀補一個檔案，再寫它的端對端測試。

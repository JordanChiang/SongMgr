# Glossary

Domain terms used in CrazyKTV SongMgr, with the code that owns each one.

## 歌庫 (song library)

- **歌庫** — the songs in `ktv_Song` (CrazySong.mdb, or CrazySong.db for a SQLite 歌庫) plus the folder that holds their files (`Global.SongMgrDestFolder`).
- **歌庫結構** — the folder and file-name layout songs are kept in: 資料夾結構 (`SongMgrFolderStructure`: 語系\歌手類別\歌手, 語系\歌手類別, 語系) and 檔案結構 (`SongMgrFileStructure`: 歌手_歌名, 歌名_歌手, 編號_歌手_歌名).
- **歌庫歸檔** (library placement) — where a song's file belongs under the 歌庫結構, and moving, copying, replacing or removing that file. `LibraryPlacement`. A file is never deleted or backed up onto itself.
  - **PlacementPurpose** — why a song is being placed: Add (加歌), Relocate (歌庫查詢 edit, 歌手管理 rename, 重複歌曲 replace), Rebuild (歌庫結構重建), UpdateStructure (歌庫結構更新).
- **加歌模式** — what happens to a file when it enters the 歌庫: 1 move, 2 copy, 3 leave in place, 4 歌庫監視. `SongAddMode`, read with `Global.AddMode`; files move only in modes 1 and 2 (`LibraryPlacement.MovesFiles`).
- **備份移除歌曲** — removed or replaced files go to `SongMgr\RemoveSong` instead of being deleted (`LibraryPlacement.RemoveSongFolder`).
- **歌庫監視** — 加歌模式 4: watched folders feed new files into 歌曲分析 and drop songs whose file is gone (`SongMonitor`).
- **歌庫檔案統計** — how many songs in `ktv_Song` have their file on disk, shown as 歌庫檔案 on the 歌庫查詢 statistics. `CommonFunc.GetSongFileCount` lists each folder once instead of checking every file. Runs when the 歌庫 loads only if 自動統計歌庫檔案 (`MainCfgAutoSongFileCount`) is on; otherwise the label reads 點此統計 and a click starts it.

## Songs and singers

- **歌曲記錄** (song record) — one song's fields by name, as text, travelling between the song grids, the edit panel, the 加歌 pipeline and `ktv_Song`. `SongRecord`. Replaces the old "|"-joined strings read by position.
- **歌曲分析** — reading 語系, 歌手, 歌手類別, 歌名, 歌曲類別 and 聲道 out of a file's name and folders. `SongAnalyzer.Analyze(file, SongAnalysisSettings, SingerLookup)`; holds no state, so files can be analysed in parallel. `SongAnalysis` runs a batch: it owns the 加歌 table, the duplicate check and the log lines.
  - **SongAnalysisSettings** — the 加歌 settings 歌曲分析 reads, captured once per run.
  - **SingerLookup** — the singers 歌曲分析 recognises: `ktv_AllSinger`, then any `ktv_Singer` entry not already there.
- **歌手資料表** — `ktv_Singer` (in the 歌庫 database, the 歌庫's own singers) or `ktv_AllSinger` (in the 參考資料庫, the 預設 list). `SingerTable`; 歌手管理 works on `SingerTable.Default`.
- **歌曲編號 (Song_Id)** — a song's number in `ktv_Song`: 5 or 6 digits (`SongMgrMaxDigitCode`), inside its 語系's 語系編號區段.
  - **語系編號區段** — the run of 歌曲編號 each 語系 owns: the ten start numbers in `SongMgrLangCode`; each runs up to the next one's start, the last up to 99999 or 999999.
  - **缺號** — an unused 歌曲編號 between a 語系's start number and its highest used number. Handed out before going past the highest. _Avoid_: 斷號.
  - **歌曲編號配置** — handing out 歌曲編號 within one operation: taking the next one, claiming a 自訂編號, counting what is left. `SongIdAllocator`: build one per operation from the 歌庫 (`FromLibrary`) and drop it afterwards.
- **套用錢櫃編號** — giving the 歌庫's songs the numbers 錢櫃 uses for them. `CashboxIdPlan`: every move is worked out before anything is written; a song sitting on a number another song gets moves to a new 歌曲編號 outside all 錢櫃 numbers; written in one transaction, the 我的最愛 moving with the songs.
- **語系**, **歌手類別**, **歌曲類別**, **聲道 (Song_Track)**, **點播次數 (Song_PlayCount)**, **拼音 / 筆劃 / 筆形** — columns of `ktv_Song`; see `SongColumns`. A singer name's 拼音 / 筆劃 / 筆形 travel together as `NameSpell` (`NameSpell.Of(name)` looks them up in the 參考資料庫's 注音 table).
- **最愛用戶** — one row of `ktv_User`: a `User_Id` and a name, owning a list of 我的最愛 in `ktv_Favorite` (key `User_Id` + `Song_Id`, so a song is on one user's list at most once). People get `0001`–`9998`; 錢櫃排行 and 金曲 lists are `^CB1`… and `^CG1`…, 錢櫃新歌 months `^CN` + yyyyMM, written whole by `SongLibraryStore.ReplaceFavoriteUser`.
- **我的最愛匯出 / 匯入** — saving every 最愛用戶 and 我的最愛 to `SongMgr\Backup\Favorite.txt` by 語系, 歌手 and 歌名 rather than 歌曲編號, and reading it back. `SongMaintenance.ImportFavorites`: replaces both tables in one transaction, so a failed import changes nothing.

## Storage and files

- **歌庫存取** (library store) — every write to `ktv_Song`, the 最愛用戶 and their 我的最愛 (`ktv_User`, `ktv_Favorite`) and the 歌手資料表. `SongLibraryStore`: open one per batch, choose columns with `SongColumns`; writes throw like `DbCommand.ExecuteNonQuery`. Removing a 最愛用戶 removes their 我的最愛; adding 我的最愛 skips songs already on the list. Only 資料庫結構更新, 匯出 SQLite 資料庫 and the DEBUG-build tools write these tables another way.
- **資料庫存取** — opening, querying and compacting either kind of database file: Access (.mdb) or SQLite (.db / .sqlite), chosen by file extension (`DbKind`). `Db`; the two providers sit in `DbHelper`. SQL is written for Access and `Db` rewrites it for SQLite.
- **資料庫結構更新** — bringing a 歌庫 database up to the tables, columns and data SongMgr needs (old tables dropped, text columns widened, `Song_ReplayGain` and `Langauage_KeyWord` added, GodLiu columns dropped, 歌手類別 合唱歌曲 renamed 合唱). `SchemaUpgrade`: `Check` reads what is missing from the file itself, so no version number is kept; `Run` backs the file up to `SongMgr\Backup`, does those steps and restores the backup if one fails. Runs by itself whenever a 歌庫 loads; a SQLite 歌庫 can only be checked. _Avoid_: 歌庫版本, 歌庫結構 (that is the folder layout).
- **參考資料庫** — CrazySongMgr (`Global.CrazyktvSongMgrDatabaseFile` in the SongMgr folder): 注音, 錢櫃, `ktv_AllSinger` and the other lists SongMgr ships with. Always the same kind as the 歌庫: CrazySongMgr.mdb with an Access 歌庫, CrazySongMgr.db with a SQLite one, never mixed (`Db.SongMgrDatabaseFor`). The title bar shows the kind and both file names.
- **匯出 SQLite 資料庫** — writing CrazySong.mdb and CrazySongMgr.mdb out as CrazySong.db and CrazySongMgr.db, each beside its .mdb, with the export version in `ktv_DBVersion`. `SqliteExport`; a .db already there is replaced only when its content differs and the user agrees.
- **YouTube 歌曲** — the `YoutubeSongList` rows UltimateKtv adds for each YouTube video it downloads; 歌庫查詢 lists, edits and deletes them. `YoutubeSongStore` (editable columns: `YoutubeColumn`).
- **設定檔** — the XML settings files (`CrazyKTV_SongMgr.cfg`, `CrazyKTV.cfg`, `UpdateDB.xml`). `SettingsFile`: read once, then by key name. `SongMgrSettings.Defaults()` is what a fresh `CrazyKTV_SongMgr.cfg` contains.
- **操作記錄** — the log list shown on the 操作記錄 tab; also the 加歌結果 failure and duplicate lists. `OperationLog` (`Global.SongLog`, `Global.FailureSongLog`, `Global.DuplicateSongLog`); safe to write from background tasks.

## Media

- **媒體探測** (media probe) — everything asked of ffprobe / ffmpeg: audio stream count, which 聲道 is the 伴唱 (`DetectTrack`), duration, 重播增益 per track, whether a file decodes cleanly. `MediaProbe`.
- **重播增益 (ReplayGain)** — measured loudness offset of the 伴唱, stored in `Song_ReplayGain`.
- **歌曲音量** — playback volume stored in `Song_Volume`; derived from 重播增益 by `MediaProbe.SongVolume` and kept inside a `VolumeRange`: 10..100 for 加歌 and 音量平衡 (`VolumeRange.Library`), 5..100 for 音量正規化 (`VolumeRange.Normalize`).

## Program updates

- **自動更新** — finding a newer SongMgr on GitHub Releases and replacing the program files once SongMgr exits. `AppUpdater`; turned on by 啟用程式自動更新 (`MainCfgEnableAutoUpdate`). It waits for any 背景工作 (the UI locked by `Common_SwitchSetUI(false)`) before asking or closing, backs up the files it will replace, and restores them if overwriting fails. User files listed in `AppUpdater.KeepExisting` are never overwritten.
- **元件** (component) — a part of the program updated on its own: App, Codec_x86, Tools. Defined only in `tools\pack-release.ps1`; SongMgr reads them from the 更新清單.
- **更新清單** — `update.json` attached to each release: every 元件's zip and the SHA256 of each of its files. Only 元件 whose files differ are downloaded. A release without one falls back to the 完整包 (`CrazyKtv_SongMgr<版本>.zip`, `AppUpdater.FullPackagePattern`).

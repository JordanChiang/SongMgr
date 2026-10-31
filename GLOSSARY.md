# Glossary

Domain terms used in CrazyKTV SongMgr, with the code that owns each one.

## 歌庫 (song library)

- **歌庫** — the songs in `ktv_Song` (CrazySong.mdb) plus the folder that holds their files (`Global.SongMgrDestFolder`).
- **歌庫結構** — the folder and file-name layout songs are kept in: 資料夾結構 (`SongMgrFolderStructure`: 語系\歌手類別\歌手, 語系\歌手類別, 語系) and 檔案結構 (`SongMgrFileStructure`: 歌手_歌名, 歌名_歌手, 編號_歌手_歌名).
- **歌庫歸檔** (library placement) — where a song's file belongs under the 歌庫結構, and moving, copying, replacing or removing that file. `LibraryPlacement`. A file is never deleted or backed up onto itself.
  - **PlacementPurpose** — why a song is being placed: Add (加歌), Relocate (歌庫查詢 edit, 歌手管理 rename, 重複歌曲 replace), Rebuild (歌庫結構重建), UpdateStructure (歌庫結構更新).
- **加歌模式** — what happens to a file when it enters the 歌庫: 1 move, 2 copy, 3 leave in place, 4 歌庫監視. `SongAddMode`, read with `Global.AddMode`; files move only in modes 1 and 2 (`LibraryPlacement.MovesFiles`).
- **備份移除歌曲** — removed or replaced files go to `SongMgr\RemoveSong` instead of being deleted (`LibraryPlacement.RemoveSongFolder`).
- **歌庫監視** — 加歌模式 4: watched folders feed new files into 歌曲分析 and drop songs whose file is gone (`SongMonitor`).

## Songs and singers

- **歌曲記錄** (song record) — one song's fields by name, as text, travelling between the song grids, the edit panel, the 加歌 pipeline and `ktv_Song`. `SongRecord`. Replaces the old "|"-joined strings read by position.
- **歌曲分析** — reading 語系, 歌手, 歌手類別, 歌名, 歌曲類別 and 聲道 out of a file's name and folders. `SongAnalyzer.Analyze(file, SongAnalysisSettings, SingerLookup)`; holds no state, so files can be analysed in parallel. `SongAnalysis` runs a batch: it owns the 加歌 table, the duplicate check and the log lines.
  - **SongAnalysisSettings** — the 加歌 settings 歌曲分析 reads, captured once per run.
  - **SingerLookup** — the singers 歌曲分析 recognises: `ktv_AllSinger`, then any `ktv_Singer` entry not already there.
- **歌手資料表** — `ktv_Singer` (in CrazySong.mdb, the 歌庫's own singers) or `ktv_AllSinger` (in CrazySongMgr.mdb, the 預設 list). `SingerTable`; 歌手管理 works on `SingerTable.Default`.
- **語系**, **歌手類別**, **歌曲類別**, **聲道 (Song_Track)**, **點播次數 (Song_PlayCount)**, **拼音 / 筆劃 / 筆形** — columns of `ktv_Song`; see `SongColumns`.

## Storage and files

- **歌庫存取** (library store) — every write to `ktv_Song`, `ktv_Favorite` and the 歌手資料表. `SongLibraryStore`: open one per batch, choose columns with `SongColumns`; writes throw like `OleDbCommand.ExecuteNonQuery`.
- **設定檔** — the XML settings files (`CrazyKTV_SongMgr.cfg`, `CrazyKTV.cfg`, `UpdateDB.xml`). `SettingsFile`: read once, then by key name. `SongMgrSettings.Defaults()` is what a fresh `CrazyKTV_SongMgr.cfg` contains.
- **操作記錄** — the log list shown on the 操作記錄 tab; also the 加歌結果 failure and duplicate lists. `OperationLog` (`Global.SongLog`, `Global.FailureSongLog`, `Global.DuplicateSongLog`); safe to write from background tasks.

## Media

- **媒體探測** (media probe) — everything asked of ffprobe / ffmpeg: audio stream count, which 聲道 is the 伴唱 (`DetectTrack`), duration, 重播增益 per track, whether a file decodes cleanly. `MediaProbe`.
- **重播增益 (ReplayGain)** — measured loudness offset of the 伴唱, stored in `Song_ReplayGain`.
- **歌曲音量** — playback volume stored in `Song_Volume`; derived from 重播增益 by `MediaProbe.SongVolume` and kept inside a `VolumeRange`: 10..100 for 加歌 and 音量平衡 (`VolumeRange.Library`), 5..100 for 音量正規化 (`VolumeRange.Normalize`).

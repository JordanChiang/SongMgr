<#
.SYNOPSIS
    產生 GitHub Release 附檔: update.json、各元件 zip 與完整安裝包.

.DESCRIPTION
    程式的自動更新會讀 release 的 update.json, 依每個元件的檔案 SHA256 比對使用者電腦上的檔案,
    只下載有變動的元件:
        App        加歌程式 (根目錄的 CrazyKTV_SongMgr.exe 及設定檔 + SongMgr\)
        Codec_x86  影音解碼器
        Tools      FFmpeg 工具

    與上一版 update.json 內容相同的元件不會重新打包, 直接沿用上一版的下載網址,
    所以 Codec_x86 / Tools 沒變時, 只需要上傳主程式那一包. 也因此舊 release 的附檔不能刪除.

    使用步驟:
        1. Visual Studio 以 FolderProfile 發佈 (輸出到 CrazyKTV_SongMgr\Publish)
        2. powershell -ExecutionPolicy Bypass -File tools\pack-release.ps1
        3. 在 GitHub 建立 tag 為版本號 (例如 2.12.0) 的 Release,
           上傳輸出資料夾裡的所有檔案 (update.json 與完整包一定要上傳)

.PARAMETER PublishDir
    發佈資料夾, 預設 CrazyKTV_SongMgr\Publish.
.PARAMETER OutDir
    輸出資料夾, 預設 <PublishDir>\ReleaseAssets\<版本>.
.PARAMETER Tag
    Release 的 tag, 預設為 exe 的檔案版本.
.PARAMETER PreviousManifest
    上一版 update.json 的路徑或網址, 預設取最新 release 的 update.json.
.PARAMETER NoReuse
    不沿用上一版的元件, 全部重新打包.
#>
param(
    [string]$PublishDir = (Join-Path $PSScriptRoot '..\CrazyKTV_SongMgr\Publish'),
    [string]$OutDir,
    [string]$Tag,
    [string]$PreviousManifest = 'https://github.com/JordanChiang/SongMgr/releases/latest/download/update.json',
    [switch]$NoReuse
)

$ErrorActionPreference = 'Stop'
$Repo = 'JordanChiang/SongMgr'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$PublishDir = (Resolve-Path $PublishDir).Path.TrimEnd('\')
$exe = Join-Path $PublishDir 'CrazyKTV_SongMgr.exe'
if (-not (Test-Path $exe)) { throw "找不到 $exe, 請先發佈程式。" }

$version = (Get-Item $exe).VersionInfo.FileVersion
if (-not $Tag) { $Tag = $version }
if (-not $OutDir) { $OutDir = Join-Path $PublishDir "ReleaseAssets\$version" }
$OutDir = [IO.Path]::GetFullPath($OutDir).TrimEnd('\')

# --- 元件定義 (rel 為相對於發佈資料夾的路徑, 以 \ 分隔) ---
# 元件清單只定義在這裡; 程式端 (AppUpdater) 依 update.json 處理, 不寫死元件名稱.
# App 收根目錄所有檔案 (目前只有 exe, 以後多出的檔案也會一起收進來, 不含 .pdb) 及 SongMgr\.
$components = @(
    @{ Name = 'App';       Desc = '加歌程式';               Zip = "SongMgr_App_$version.zip"
       Match = { param($rel) (($rel -notmatch '\\') -and ($rel -notmatch '\.pdb$')) -or $rel.StartsWith('SongMgr\', 'OrdinalIgnoreCase') } },
    @{ Name = 'Codec_x86'; Desc = '影音解碼器 (Codec_x86)'; Zip = "Codec_x86_$version.zip"
       Match = { param($rel) $rel.StartsWith('Codec_x86\', 'OrdinalIgnoreCase') } },
    @{ Name = 'Tools';     Desc = 'FFmpeg 工具 (Tools)';    Zip = "Tools_$version.zip"
       Match = { param($rel) $rel.StartsWith('Tools\', 'OrdinalIgnoreCase') } }
)

# --- 讀取上一版 update.json ---
function Get-PreviousManifest {
    if ($NoReuse -or -not $PreviousManifest) { return $null }
    try {
        if (Test-Path $PreviousManifest) {
            $text = [IO.File]::ReadAllText((Resolve-Path $PreviousManifest).Path, [Text.Encoding]::UTF8)
        } else {
            $r = Invoke-WebRequest -UseBasicParsing -Uri $PreviousManifest -Headers @{ 'User-Agent' = 'CrazyKTV_SongMgr-PackRelease' }
            $text = if ($r.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($r.Content) } else { [string]$r.Content }
        }
        $m = $text.TrimStart([char]0xFEFF) | ConvertFrom-Json
        Write-Host "上一版清單: $($m.tag)"
        return $m
    } catch {
        Write-Host "沒有取得上一版 update.json, 所有元件都重新打包。($($_.Exception.Message))"
        return $null
    }
}

function Test-SameFiles($prevFiles, [System.Collections.Specialized.OrderedDictionary]$files) {
    if ($null -eq $prevFiles) { return $false }
    $prev = @{}
    foreach ($p in $prevFiles.PSObject.Properties) { $prev[$p.Name] = ([string]$p.Value).ToLowerInvariant() }
    if ($prev.Count -ne $files.Count) { return $false }
    foreach ($k in $files.Keys) {
        if (-not $prev.ContainsKey($k) -or $prev[$k] -ne $files[$k]) { return $false }
    }
    return $true
}

function New-Zip([string]$path, [string[]]$rels) {
    if (Test-Path $path) { Remove-Item $path -Force }
    $fs = [IO.File]::Open($path, [IO.FileMode]::CreateNew)
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($rel in $rels) {
                [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $PublishDir $rel), $rel.Replace('\', '/'), [IO.Compression.CompressionLevel]::Optimal)
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

# --- 列出發佈檔案 (排除輸出資料夾) ---
$assetsRoot = Join-Path $PublishDir 'ReleaseAssets'
$allRels = Get-ChildItem -LiteralPath $PublishDir -Recurse -File |
    Where-Object { -not $_.FullName.StartsWith($assetsRoot + '\', 'OrdinalIgnoreCase') -and -not $_.FullName.StartsWith($OutDir + '\', 'OrdinalIgnoreCase') } |
    ForEach-Object { $_.FullName.Substring($PublishDir.Length + 1) } |
    Where-Object { $_ -notmatch '\.(cfg|log)$' } |   # 在發佈資料夾執行過程式留下的設定檔與記錄檔
    Sort-Object

if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Force $OutDir | Out-Null

$prevManifest = Get-PreviousManifest
$manifestComponents = @()
$packed = @()

foreach ($c in $components) {
    $rels = @($allRels | Where-Object { & $c.Match $_ })
    if ($rels.Count -eq 0) { throw "元件 $($c.Name) 沒有任何檔案。" }

    $files = [ordered]@{}
    foreach ($rel in $rels) {
        $files[$rel.Replace('\', '/')] = (Get-FileHash -LiteralPath (Join-Path $PublishDir $rel) -Algorithm SHA256).Hash.ToLowerInvariant()
    }

    $prev = $null
    if ($prevManifest) { $prev = $prevManifest.components | Where-Object { $_.name -eq $c.Name } | Select-Object -First 1 }

    if ($prev -and (Test-SameFiles $prev.files $files)) {
        $url = $prev.url; $size = [long]$prev.size
        Write-Host ("{0,-10} 沒有變動, 沿用 {1}" -f $c.Name, $url)
    } else {
        $zipPath = Join-Path $OutDir $c.Zip
        New-Zip $zipPath $rels
        $url = "https://github.com/$Repo/releases/download/$Tag/$($c.Zip)"
        $size = (Get-Item $zipPath).Length
        $packed += $c.Zip
        Write-Host ("{0,-10} 已打包 {1} ({2:N1} MB, {3} 個檔案)" -f $c.Name, $c.Zip, ($size / 1MB), $rels.Count)
    }

    $manifestComponents += [ordered]@{ name = $c.Name; desc = $c.Desc; url = $url; size = $size; files = $files }
}

# --- 完整安裝包 (給手動下載及第一次安裝; 沒有 update.json 時自動更新只認這個檔名, 見 AppUpdater.FullPackagePattern) ---
$fullZip = "CrazyKtv_SongMgr$version.zip"
$componentRels = @($allRels | Where-Object { $rel = $_; @($components | Where-Object { & $_.Match $rel }).Count -gt 0 })
New-Zip (Join-Path $OutDir $fullZip) $componentRels
Write-Host ("{0,-10} 已打包 {1} ({2:N1} MB)" -f '完整包', $fullZip, ((Get-Item (Join-Path $OutDir $fullZip)).Length / 1MB))

# --- update.json ---
$manifest = [ordered]@{ version = $version; tag = $Tag; components = $manifestComponents }
$json = $manifest | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText((Join-Path $OutDir 'update.json'), $json, (New-Object Text.UTF8Encoding($false)))

Write-Host ''
Write-Host "輸出資料夾: $OutDir"
Write-Host "請建立 tag 為 $Tag 的 Release, 上傳以下檔案:"
Get-ChildItem -LiteralPath $OutDir -File | ForEach-Object { Write-Host "  $($_.Name)" }

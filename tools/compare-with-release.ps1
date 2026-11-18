<#
.SYNOPSIS
    新舊版比對: 用真實歌庫比較目前的程式和已發佈的版本, 使用者看到的結果是否相同.

.DESCRIPTION
    把已發佈的版本 (預設 2.11.4) 取到暫存的 git worktree, 兩版各自在畫面外開啟真正的主視窗,
    以同一份歌庫副本與歌庫設定檔執行:
        歌庫查詢   一般查詢、異常查詢、我的最愛, 記下每次的清單與狀態列
        加歌       新歌 (add_new)、所有歌曲檔 (add_all)、新歌加上已在歌庫的歌 (add_mixed),
                   記下歌曲分析清單、加歌結果、操作記錄與加歌後的歌庫
    最後逐段比較兩版的輸出, 寫成報告.

    需要互動桌面 (會開出畫面外的視窗與自動回答的對話框), 全部約 15 到 20 分鐘.
    歌庫只用副本, 原檔不會被修改.

    使用方式 (在專案根目錄):
        powershell -ExecutionPolicy Bypass -File tools\compare-with-release.ps1

.PARAMETER Baseline
    要比較的已發佈版本 (git tag), 預設 2.11.4. 該版本要有 SongMgr.Tests 專案.
.PARAMETER Database
    歌庫 (Access .mdb), 預設專案根目錄的 CrazySong.mdb.
.PARAMETER Settings
    歌庫設定檔, 預設專案根目錄的 CrazyKTV_SongMgr.cfg.
.PARAMETER WorkDir
    暫存與輸出資料夾, 預設 %TEMP%\SongMgrCompare. 報告在 <WorkDir>\report.txt.
.PARAMETER KeepWorktree
    結束後保留已發佈版本的 worktree (預設會移除).
#>
param(
    [string]$Baseline = '2.11.4',
    [string]$Database = (Join-Path $PSScriptRoot '..\CrazySong.mdb'),
    [string]$Settings = (Join-Path $PSScriptRoot '..\CrazyKTV_SongMgr.cfg'),
    [string]$WorkDir = (Join-Path $env:TEMP 'SongMgrCompare'),
    [switch]$KeepWorktree
)

$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Database = (Resolve-Path $Database).Path
$Settings = (Resolve-Path $Settings).Path
$WorkDir = [IO.Path]::GetFullPath($WorkDir).TrimEnd('\')
$baseTree = Join-Path $WorkDir "baseline-$Baseline"
$data = Join-Path $WorkDir 'data'
$outBase = Join-Path $WorkDir "out-$Baseline"
$outNew = Join-Path $WorkDir 'out-current'
$report = Join-Path $WorkDir 'report.txt'
$compareVars = 'SONGMGR_COMPARE_SOURCE', 'SONGMGR_COMPARE_DATA', 'SONGMGR_COMPARE_CFG', 'SONGMGR_COMPARE_OUT',
               'SONGMGR_COMPARE_ADD', 'SONGMGR_COMPARE_BASE', 'SONGMGR_COMPARE_NEW', 'SONGMGR_COMPARE_REPORT'

function Remove-Baseline {
    if (Test-Path $baseTree) {
        $ErrorActionPreference = 'Continue'   # git 的 stderr 不要變成例外
        git -C $repo worktree remove --force $baseTree 2>&1 | Out-Null
        if (Test-Path $baseTree) { Remove-Item -Recurse -Force $baseTree }
        git -C $repo worktree prune
    }
}

function Build-Tests([string]$tree) {
    Write-Host "建置 $tree"
    dotnet build (Join-Path $tree 'SongMgr.Tests\SongMgr.Tests.csproj') -c Debug -v q -nologo
    if ($LASTEXITCODE -ne 0) { throw "建置失敗: $tree" }
}

# 以指定的環境變數執行一個比對測試; 測試沒通過就停止.
function Invoke-Compare([string]$tree, [string]$test, [hashtable]$vars) {
    foreach ($v in $compareVars) { Remove-Item "Env:$v" -ErrorAction SilentlyContinue }
    foreach ($k in $vars.Keys) { Set-Item "Env:$k" $vars[$k] }
    $start = Get-Date
    dotnet test (Join-Path $tree 'SongMgr.Tests\SongMgr.Tests.csproj') --no-build -v q --nologo `
        -s (Join-Path $tree 'SongMgr.Tests\x86.runsettings') --filter "FullyQualifiedName=SongMgr.Tests.Compare.$test" | Out-Host
    $ok = $LASTEXITCODE -eq 0
    foreach ($v in $compareVars) { Remove-Item "Env:$v" -ErrorAction SilentlyContinue }
    if (-not $ok) { throw "$test 失敗 ($tree)" }
    Write-Host ("  {0} 完成, {1:N0} 秒" -f $test, ((Get-Date) - $start).TotalSeconds)
}

New-Item -ItemType Directory -Force $WorkDir | Out-Null
try {
    Remove-Baseline
    Write-Host "取出 $Baseline 到 $baseTree"
    git -C $repo worktree add --detach $baseTree $Baseline
    if ($LASTEXITCODE -ne 0) { throw "無法取出 $Baseline" }
    if (-not (Test-Path (Join-Path $baseTree 'SongMgr.Tests\SongMgr.Tests.csproj'))) { throw "$Baseline 沒有 SongMgr.Tests 專案, 無法比對" }

    # 兩版用同一份驅動程式; 資料準備和報告只在目前的版本執行.
    $compareSrc = Join-Path $repo 'SongMgr.Tests\Compare\LibraryCompare.cs'
    New-Item -ItemType Directory -Force (Join-Path $baseTree 'SongMgr.Tests\Compare') | Out-Null
    Copy-Item $compareSrc (Join-Path $baseTree 'SongMgr.Tests\Compare\LibraryCompare.cs')

    Build-Tests $repo
    Build-Tests $baseTree

    foreach ($d in $outBase, $outNew) {
        if (Test-Path $d) { Remove-Item -Recurse -Force $d }
        New-Item -ItemType Directory -Force $d | Out-Null
    }

    Write-Host "準備資料: $Database"
    Invoke-Compare $repo 'CompareData.Prepare' @{ SONGMGR_COMPARE_SOURCE = $Database; SONGMGR_COMPARE_DATA = $data; SONGMGR_COMPARE_CFG = $Settings }
    Get-Content (Join-Path $data 'README.txt') -Encoding UTF8 | ForEach-Object { Write-Host "  $_" }

    $runs = @(
        @{ Test = 'LibraryCompare.Queries'; Add = $null },
        @{ Test = 'LibraryCompare.AddSongs'; Add = 'add_new' },
        @{ Test = 'LibraryCompare.AddSongs'; Add = 'add_all' },
        @{ Test = 'LibraryCompare.AddSongs'; Add = 'add_mixed' }
    )
    foreach ($run in $runs) {
        foreach ($side in @(@{ Tree = $baseTree; Out = $outBase; Name = $Baseline }, @{ Tree = $repo; Out = $outNew; Name = '目前版本' })) {
            Write-Host ("{0}: {1} {2}" -f $side.Name, $run.Test, $run.Add)
            $vars = @{ SONGMGR_COMPARE_DATA = $data; SONGMGR_COMPARE_CFG = $Settings; SONGMGR_COMPARE_OUT = $side.Out }
            if ($run.Add) { $vars.SONGMGR_COMPARE_ADD = $run.Add }
            Invoke-Compare $side.Tree $run.Test $vars
        }
    }

    Invoke-Compare $repo 'CompareReport.Write' @{ SONGMGR_COMPARE_BASE = $outBase; SONGMGR_COMPARE_NEW = $outNew; SONGMGR_COMPARE_REPORT = $report }
    Write-Host ''
    Write-Host (Get-Content $report -Encoding UTF8 -TotalCount 1)
    Write-Host "報告: $report"
}
finally {
    if (-not $KeepWorktree) { Remove-Baseline }
}

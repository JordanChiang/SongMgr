# Read-only check: is this machine and checkout fit to drive CrazyKTV_SongMgr end to end?
# Prints one OK / WARN / FAIL line per check. Exit code 0 = ready to drive, 1 = do not drive.
# Run from the repo (or worktree) root:  powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\doctor.ps1
param([string]$Repo)

$ErrorActionPreference = 'Stop'
if (-not $Repo) {
    $Repo = (& git rev-parse --show-toplevel 2>$null)
    if (-not $Repo) { $Repo = (Get-Location).Path }
}
$Repo = (Resolve-Path $Repo).Path
$bin = Join-Path $Repo 'SongMgr.Tests\bin\Debug\net8.0-windows7.0'
$sandboxRoot = Join-Path ([IO.Path]::GetTempPath()) 'SongMgrTests'
$fail = $false

function Report([string]$level, [string]$text) {
    Write-Output ("{0,-4} {1}" -f $level, $text)
    if ($level -eq 'FAIL') { $script:fail = $true }
}

Report 'INFO' "repo: $Repo"

# 32-bit runtime: the app and Jet OLEDB 4.0 are x86 only.
$dotnet86 = 'C:\Program Files (x86)\dotnet\dotnet.exe'
if (Test-Path $dotnet86) { Report 'OK' "x86 dotnet runtime: $dotnet86" } else { Report 'FAIL' "no x86 dotnet runtime at $dotnet86" }

if (Test-Path 'Registry::HKEY_CLASSES_ROOT\Microsoft.Jet.OLEDB.4.0') { Report 'OK' 'Microsoft.Jet.OLEDB.4.0 is registered' }
else { Report 'FAIL' 'Microsoft.Jet.OLEDB.4.0 is not registered: the app cannot open CrazySong.mdb' }

# UI Automation needs a real, unlocked desktop: it moves the mouse and sends keys.
$session = (Get-Process -Id $PID).SessionId
if ([Environment]::UserInteractive -and $session -ne 0) { Report 'OK' "interactive desktop (session $session)" }
else { Report 'FAIL' 'no interactive desktop: end-to-end tests skip themselves here' }
if (Get-Process LogonUI -ErrorAction SilentlyContinue) { Report 'FAIL' 'the screen looks locked (LogonUI is running): clicks will not land' }
else { Report 'OK' 'screen is not locked' }

# Build output the sandboxes are copied from.
$exe = Join-Path $bin 'CrazyKTV_SongMgr.exe'
$testDll = Join-Path $bin 'SongMgr.Tests.dll'
if ((Test-Path $exe) -and (Test-Path $testDll)) {
    $built = (Get-Item $testDll).LastWriteTime
    $newest = Get-ChildItem (Join-Path $Repo 'CrazyKTV_SongMgr'), (Join-Path $Repo 'SongMgr.Tests') -Recurse -Include *.cs, *.resx, *.csproj -File |
        Where-Object { $_.FullName -notmatch '\\(bin|obj|Publish)\\' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $version = (Get-Item $exe).VersionInfo.FileVersion
    if ($newest -and $newest.LastWriteTime -gt $built) {
        Report 'WARN' ("build is stale: {0} changed after the last build ({1:yyyy-MM-dd HH:mm}); build before driving" -f $newest.FullName.Substring($Repo.Length + 1), $built)
    } else {
        Report 'OK' ("build output is current: CrazyKTV_SongMgr.exe {0}, built {1:yyyy-MM-dd HH:mm}" -f $version, $built)
    }
} else {
    Report 'FAIL' "no build output in $bin; run: dotnet build SongMgr.Tests\SongMgr.Tests.csproj -c Debug -v q -nologo"
}
foreach ($f in 'Fixtures\CrazySongEmpty.mdb', 'Fixtures\CrazySongMgr.mdb', 'Tools\ffmpeg.exe', 'SongMgr') {
    if (-not (Test-Path (Join-Path $bin $f))) { Report 'FAIL' "missing $f in the build output (a sandbox needs it)" }
}

# One driver at a time: two runs would fight over the mouse and keyboard.
$ours = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -and $_.ExecutablePath -like "$sandboxRoot\e2e_*" })
if ($ours.Count -gt 0) {
    Report 'FAIL' ("a sandboxed SongMgr is already running ({0}): another run is driving the desktop, or one was stranded; wait, or run cleanup.ps1 if it is stranded" -f (($ours | ForEach-Object { "pid $($_.ProcessId)" }) -join ', '))
} else { Report 'OK' 'no sandboxed SongMgr running' }
$hosts = @(Get-CimInstance Win32_Process -Filter "Name='testhost.x86.exe'" | Where-Object { $_.CommandLine -like '*SongMgr.Tests*' })
if ($hosts.Count -gt 0) { Report 'FAIL' ("a SongMgr.Tests test run is in progress (testhost pid {0})" -f (($hosts | ForEach-Object { $_.ProcessId }) -join ', ')) }
else { Report 'OK' 'no SongMgr.Tests run in progress' }

$other = @(Get-CimInstance Win32_Process -Filter "Name='CrazyKTV_SongMgr.exe'" | Where-Object { -not ($_.ExecutablePath -like "$sandboxRoot\*") })
if ($other.Count -gt 0) { Report 'INFO' ("the user's own SongMgr is open ({0}); leave it alone, sessions only touch their sandbox" -f (($other | ForEach-Object { $_.ExecutablePath }) -join ', ')) }

if ($fail) { Write-Output 'NOT READY'; exit 1 } else { Write-Output 'READY'; exit 0 }

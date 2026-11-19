# Tears down what end-to-end runs left behind: processes started from a sandbox under %TEMP%\SongMgrTests\e2e_*
# (SongMgr itself, the ffmpeg it runs) and the sandbox folders. Matches by executable path, never by process name,
# so the user's own SongMgr is never touched. Never deletes evidence (%TEMP%\SongMgrEvidence, %TEMP%\SongMgrTests\e2e-*.png).
# Run from anywhere:  powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\cleanup.ps1
$ErrorActionPreference = 'Continue'
$sandboxRoot = Join-Path ([IO.Path]::GetTempPath()) 'SongMgrTests'

$running = @(Get-CimInstance Win32_Process -Filter "Name='testhost.x86.exe'" | Where-Object { $_.CommandLine -like '*SongMgr.Tests*' })
if ($running.Count -gt 0) {
    Write-Output ("SKIP a SongMgr.Tests run is in progress (testhost pid {0}); its sandboxes are in use. Run cleanup after it ends." -f (($running | ForEach-Object { $_.ProcessId }) -join ', '))
    exit 1
}

$procs = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -and $_.ExecutablePath -like "$sandboxRoot\e2e_*" })
foreach ($p in $procs) {
    Write-Output "STOP pid $($p.ProcessId) $($p.ExecutablePath)"
    Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
}
if ($procs.Count -gt 0) { Start-Sleep -Milliseconds 500 }

$dirs = @(Get-ChildItem $sandboxRoot -Directory -Filter 'e2e_*' -ErrorAction SilentlyContinue)
foreach ($d in $dirs) {
    Get-ChildItem $d.FullName -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Attributes = 'Normal' }
    Remove-Item $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $d.FullName) { Write-Output "LEFT $($d.FullName) (still in use)" } else { Write-Output "DEL  $($d.FullName)" }
}
Write-Output ("done: {0} process(es) stopped, {1} sandbox folder(s) removed" -f $procs.Count, $dirs.Count)

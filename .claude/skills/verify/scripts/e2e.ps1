# Drives the real CrazyKTV_SongMgr.exe through the end-to-end tests and keeps the proof.
#   doctor -> (build) -> dotnet test with SONGMGR_E2E_EVIDENCE set -> cleanup -> list the evidence
# Run from the repo (or worktree) root:
#   powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\e2e.ps1
#   powershell -ExecutionPolicy Bypass -File .claude\skills\verify\scripts\e2e.ps1 -Filter "Category=EndToEnd&DisplayName~Cashbox" -NoBuild
# Exit code: the test run's (0 = all passed), or 2 when doctor said not ready.
param(
    [string]$Filter = 'Category=EndToEnd',
    [switch]$NoBuild,
    [string]$Evidence
)

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = (& git rev-parse --show-toplevel 2>$null)
if (-not $repo) { $repo = (Get-Location).Path }
$repo = (Resolve-Path $repo).Path
Set-Location $repo
if (-not $Evidence) { $Evidence = Join-Path ([IO.Path]::GetTempPath()) ("SongMgrEvidence\" + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Force $Evidence | Out-Null

if (-not $NoBuild) {
    & dotnet build SongMgr.Tests\SongMgr.Tests.csproj -c Debug -v q -nologo
    if ($LASTEXITCODE -ne 0) { Write-Output 'build failed'; exit $LASTEXITCODE }
}

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'doctor.ps1') -Repo $repo
if ($LASTEXITCODE -ne 0) { Write-Output 'doctor: not ready, nothing driven'; exit 2 }

Write-Output "Driving SongMgr: windows will open on the desktop; do not use the mouse or keyboard until this ends."
$env:SONGMGR_E2E_EVIDENCE = $Evidence
Remove-Item Env:SONGMGR_SKIP_E2E -ErrorAction SilentlyContinue
$code = 1
try {
    & dotnet test SongMgr.Tests\SongMgr.Tests.csproj --no-build -v q --nologo -s SongMgr.Tests\x86.runsettings `
        --filter $Filter --logger 'trx;LogFileName=e2e.trx' --results-directory $Evidence |
        Tee-Object -FilePath (Join-Path $Evidence 'test-output.txt')
    $code = $LASTEXITCODE
}
finally {
    Remove-Item Env:SONGMGR_E2E_EVIDENCE -ErrorAction SilentlyContinue
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'cleanup.ps1')
}

Write-Output ''
Write-Output "EVIDENCE $Evidence"
Get-ChildItem $Evidence -Recurse -File | ForEach-Object { '  ' + $_.FullName.Substring($Evidence.Length + 1) }
$shots = @(Get-ChildItem (Join-Path ([IO.Path]::GetTempPath()) 'SongMgrTests') -Filter 'e2e-*.png' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -gt (Get-Item $Evidence).CreationTime })
foreach ($s in $shots) { Write-Output "FAILURE SCREENSHOT $($s.FullName)" }
exit $code

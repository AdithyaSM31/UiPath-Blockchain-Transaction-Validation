<#
    run_tests.ps1 - execute the UiPath test cases headlessly.

    Studio's Test Explorer is the normal way to run these (Test tab -> Run All Tests).
    This script does the same thing from the command line so the acceptance suite and any
    CI can check them.

    Each test asserts with VerifyExpression for the Test Explorer report, and then throws
    if any assertion failed - which is what makes a failure visible as a non-zero exit code
    here, since VerifyExpression on its own does not fail the process.

    Run:  powershell -File tools\run_tests.ps1
#>
$ErrorActionPreference = "Stop"

$Root       = Split-Path $PSScriptRoot -Parent
$ProjectDir = Join-Path $Root "BlockchainLogisticsValidator"
$UiRobot    = Join-Path $env:LOCALAPPDATA "Programs\UiPath\Studio\UiRobot.exe"
$PkgOut     = Join-Path $env:TEMP "blv-packages"

$ProjectJson = Join-Path $ProjectDir "project.json"
$tests = Get-ChildItem (Join-Path $ProjectDir "Tests") -Filter "TC*.xaml" | Sort-Object Name

# A *process* package deliberately excludes anything listed in designOptions
# .fileInfoCollection, because test cases are meant to ship in a separate test package.
# That registration is what makes Studio's Test Explorer list them, so it stays in the
# deliverable - but it also means the Robot CLI cannot find them.
#
# So: drop the registration just long enough to pack, run the tests as ordinary entry
# points, then put project.json back exactly as it was. Studio is unaffected either way.
$projectBackup = Get-Content $ProjectJson -Raw

try {
    $j = $projectBackup | ConvertFrom-Json
    $j.designOptions.fileInfoCollection = @()
    ($j | ConvertTo-Json -Depth 30) | Set-Content $ProjectJson -Encoding utf8

    & (Join-Path $Root "build.ps1") -SkipRun 6>&1 | Out-Null
    $pkg = Get-ChildItem $PkgOut -Filter "*.nupkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $pkg) { throw "No package produced." }
}
finally {
    # Restore everything EXCEPT projectVersion. UiRobot extracts packages into a cache
    # keyed by name+version, so rewinding the version makes the next pack reuse a number
    # that already has a cached extraction - and the robot then silently runs the OLD
    # package. Keeping the version monotonic avoids that entirely.
    $restored = $projectBackup | ConvertFrom-Json
    $packedVersion = (Get-Content $ProjectJson -Raw | ConvertFrom-Json).projectVersion
    $restored.projectVersion = $packedVersion
    ($restored | ConvertTo-Json -Depth 30) | Set-Content $ProjectJson -Encoding utf8
}

Write-Host "`n=== UiPath test cases ===" -ForegroundColor Cyan
Write-Host "  package: $($pkg.Name)`n" -ForegroundColor DarkGray

$failed = @()
foreach ($t in $tests) {
    $entry = "Tests\$($t.Name)"
    $startedAt = Get-Date
    # A failing test makes UiRobot write to stderr, which with ErrorActionPreference=Stop
    # would abort the whole loop at the first failure. Suppress the stream and read the
    # exit code instead, so every test runs and the report is complete.
    $ErrorActionPreference = "Continue"
    & $UiRobot execute --file $pkg.FullName --entry $entry 2>$null | Out-Null
    $code = $LASTEXITCODE
    $ErrorActionPreference = "Stop"

    # Pull this run's assertion detail out of the execution log.
    $logFile = Get-ChildItem (Join-Path $env:LOCALAPPDATA "UiPath\Logs") -Filter "*Execution*.log" |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $detail = $null
    Get-Content $logFile.FullName -Tail 80 | ForEach-Object {
        if ($_ -match '^(\d\d:\d\d:\d\d)\.\d+\s+\w+\s+(\{.*\})$') {
            try { $o = $Matches[2] | ConvertFrom-Json } catch { return }
            if ($o.timeStamp -and ([datetime]$o.timeStamp) -lt $startedAt.AddSeconds(-2)) { return }
            if ($o.level -eq 'Error' -and -not $detail) { $detail = ($o.message -split "`n")[0] }
            if ($o.message -match 'FAILED:') { $detail = ($o.message -split "`n")[0] }
        }
    }

    $name = $t.BaseName
    if ($code -eq 0) {
        Write-Host ("  [PASS] {0}" -f $name) -ForegroundColor Green
    } else {
        Write-Host ("  [FAIL] {0}" -f $name) -ForegroundColor Red
        if ($detail) { Write-Host ("         {0}" -f $detail) -ForegroundColor DarkGray }
        $failed += $name
    }
}

Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
$passed = $tests.Count - $failed.Count
if ($failed.Count -eq 0) {
    Write-Host "$passed/$($tests.Count) test cases passed." -ForegroundColor Green
    exit 0
} else {
    Write-Host "$passed/$($tests.Count) test cases passed. Failed: $($failed -join ', ')" -ForegroundColor Red
    exit 1
}

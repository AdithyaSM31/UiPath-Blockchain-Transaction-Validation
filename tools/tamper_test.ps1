<#
    tamper_test.ps1 - demonstrate that the audit trail is tamper-EVIDENT.

    Nothing stops someone opening ValidationAuditLog.csv and editing it. The point of
    the hash chain is that doing so cannot go unnoticed. This script proves that end to
    end and puts the file back exactly as it was:

        1. verify the chain            -> expect VALID
        2. flip one FAIL verdict to PASS, as a cover-up would
        3. verify again                -> expect INVALID, naming the exact line
        4. restore from backup and verify -> VALID again

    Run:  powershell -File tools\tamper_test.ps1
#>
$ErrorActionPreference = "Stop"

$Root      = Split-Path $PSScriptRoot -Parent
$ProjectDir = Join-Path $Root "BlockchainLogisticsValidator"
$Audit     = Join-Path $ProjectDir "Data\Output\AuditLogs\ValidationAuditLog.csv"
$UiRobot   = Join-Path $env:LOCALAPPDATA "Programs\UiPath\Studio\UiRobot.exe"

if (-not (Test-Path $Audit)) {
    throw "No audit log yet. Run .\build.ps1 first to produce one."
}

function Invoke-Verify([string]$Label) {
    $pkg = Get-ChildItem (Join-Path $env:TEMP "blv-packages") -Filter "*.nupkg" |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $pkg) { throw "No package found. Run .\build.ps1 first." }

    $startedAt = Get-Date
    # 12_Verify_AuditChain resolves the log location from Config.xlsx when given no
    # argument, which avoids passing a path containing spaces through the CLI.
    & $UiRobot execute --file $pkg.FullName --entry "Workflows\12_Verify_AuditChain.xaml" 2>&1 | Out-Null

    $logFile = Get-ChildItem (Join-Path $env:LOCALAPPDATA "UiPath\Logs") -Filter "*Execution*.log" |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $verdict = $null
    Get-Content $logFile.FullName -Tail 60 | ForEach-Object {
        if ($_ -match '^(\d\d:\d\d:\d\d)\.\d+\s+\w+\s+(\{.*\})$') {
            try { $o = $Matches[2] | ConvertFrom-Json } catch { return }
            if ($o.timeStamp -and ([datetime]$o.timeStamp) -lt $startedAt.AddSeconds(-2)) { return }
            if ($o.message -match '^\[VERIFY\] (VALID|INVALID)') { $verdict = $o.message }
        }
    }

    $colour = if ($verdict -match "INVALID") { "Red" } else { "Green" }
    Write-Host "`n$Label" -ForegroundColor Cyan
    Write-Host "  $verdict" -ForegroundColor $colour
    return $verdict
}

$backup = "$Audit.bak"
Copy-Item $Audit $backup -Force

try {
    $v1 = Invoke-Verify "STEP 1 - chain as written by the bot"
    if ($v1 -notmatch "VALID -") { throw "Expected a clean chain before tampering. Got: $v1" }

    # --- Tamper -----------------------------------------------------------
    $lines = Get-Content $Audit
    $target = -1
    for ($i = 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match ',FAIL,') { $target = $i; break }
    }
    if ($target -lt 0) { throw "No FAIL entry to tamper with." }

    $seq = ($lines[$target] -split ',')[0]
    $ship = ($lines[$target] -split ',')[6]
    $lines[$target] = $lines[$target] -replace ',FAIL,[A-Z]+,', ',PASS,NONE,'
    Set-Content $Audit -Value $lines -Encoding utf8

    Write-Host "`nSTEP 2 - tampering" -ForegroundColor Cyan
    Write-Host "  entry seq=$seq ($ship): verdict changed FAIL -> PASS, exactly as a cover-up would" -ForegroundColor Yellow
    Write-Host "  file line $($target + 1); no other byte touched" -ForegroundColor DarkGray

    $v2 = Invoke-Verify "STEP 3 - same chain, re-verified"
    if ($v2 -notmatch "INVALID") {
        throw "TAMPER TEST FAILED: the altered log still verified as valid."
    }
}
finally {
    Copy-Item $backup $Audit -Force
    Remove-Item $backup -Force
}

$v3 = Invoke-Verify "STEP 4 - restored from backup"

Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($v3 -match "VALID -") {
    Write-Host "Tamper test PASSED: the edit was detected and the file restored." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Restore did not verify - check $Audit" -ForegroundColor Red
    exit 1
}

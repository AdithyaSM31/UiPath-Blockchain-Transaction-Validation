<#
    run_all_checks.ps1 - full acceptance suite for the project.

    Runs every verifiable claim the project makes, end to end, and reports pass/fail.
    Use it before a review or a demo: if this is green, the whole pipeline works.

        1. Dataset oracle       - the seeded anomalies exist and the ABI decodes
        2. Workflow Analyzer    - no Error-severity findings
        3. Batch run            - all six rules fire with the expected counts
        4. Rule toggle          - a rule disabled in Excel is skipped (Novelty 1)
        5. HTML dashboard       - a self-contained dashboard is produced each run
        6. Escalation           - a CRITICAL anomaly is captured (Novelty 2)
        7. Chain swap           - identical results on a second network (Novelty 3)
        8. Tamper evidence      - an edited audit log is detected (Novelty 4)
        9. WEB mode             - explorer-page scraping reaches the same result
       10. UiPath test cases    - 7 unit/integration tests via VerifyExpression
       11. Dispatch mode        - the run splits into per-shipment queue payloads

    Run:  powershell -File tools\run_all_checks.ps1
#>
$ErrorActionPreference = "Stop"

$Root       = Split-Path $PSScriptRoot -Parent
$ProjectDir = Join-Path $Root "BlockchainLogisticsValidator"
$Build      = Join-Path $Root "build.ps1"
$Audit      = Join-Path $ProjectDir "Data\Output\AuditLogs\ValidationAuditLog.csv"
$ConfigXlsx = Join-Path $ProjectDir "Data\Config.xlsx"

$results = [System.Collections.ArrayList]::new()
function Add-Result([string]$Name, [bool]$Ok, [string]$Detail) {
    [void]$results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail })
    $tag = if ($Ok) { "PASS" } else { "FAIL" }
    $col = if ($Ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1}" -f $tag, $Name) -ForegroundColor $col
    if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Invoke-Build([hashtable]$Set = @{}) {
    # Write-Host output goes to the information stream; redirect it too, or the
    # build chatter leaks past this capture and buries the suite results.
    $out = & $Build -Set $Set -LogTail 80 6>&1 2>&1 | Out-String
    return $out
}

Write-Host "`n=== Acceptance suite ===`n" -ForegroundColor Cyan

# --- 1. Dataset oracle ----------------------------------------------------
$oracle = & python (Join-Path $PSScriptRoot "verify_seeded_data.py") 2>&1 | Out-String
Add-Result "Dataset oracle" ($oracle -match "All checks passed") `
    (($oracle -split "`n" | Where-Object { $_ -match "checks passed|CHECK\(S\) FAILED" }) -join "")

# --- 2. Workflow Analyzer -------------------------------------------------
$an = & $Build -Analyze -SkipRun 6>&1 2>&1 | Out-String
Add-Result "Workflow Analyzer" ($an -notmatch "\[ERROR\]") `
    $(if ($an -match "No Error-severity findings") { "no Error-severity findings" } else { "see output" })

# --- 3. Batch run: all six rules ------------------------------------------
Remove-Item $Audit -Force -ErrorAction SilentlyContinue
$run = Invoke-Build
$expected = "total=48 pass=40 warn=1 fail=7 escalations=1"
$gotLine = ($run -split "`n" | Where-Object { $_ -match "total=\d+ pass=" } | Select-Object -First 1)
Add-Result "Batch run - six rules fire" ($run -match [regex]::Escape($expected)) $gotLine.Trim()

$allRules = @("R1(TimestampCheck)=2","R2(QuantityMatch)=1","R3(AddressWhitelist)=1",
              "R4(DuplicateDetection)=2","R5(SequenceValidation)=1","R6(SmartContractEvent)=1")
$missing = $allRules | Where-Object { $run -notmatch [regex]::Escape($_) }
Add-Result "Every rule reports its seeded anomaly" ($missing.Count -eq 0) `
    $(if ($missing) { "missing: $($missing -join ', ')" } else { "R1-R6 all present" })

# --- 4. Rule toggle (Novelty 1) -------------------------------------------
$py = @'
import sys
from openpyxl import load_workbook
p, val = sys.argv[1], sys.argv[2]
wb = load_workbook(p); ws = wb["ValidationRules"]
for row in ws.iter_rows(min_row=2):
    if row[0].value == "R4":
        row[2].value = val
wb.save(p)
'@
$pyFile = Join-Path $env:TEMP "toggle_r4.py"
Set-Content -Path $pyFile -Value $py -Encoding utf8

& python $pyFile $ConfigXlsx "FALSE"
try {
    $off = Invoke-Build
    $skipped = $off -match "R4 is disabled in Config.xlsx - skipped"
    $droppedTo5 = $off -match "fail=5"
    Add-Result "Rule toggled off in Excel is skipped (Novelty 1)" ($skipped -and $droppedTo5) `
        "R4 disabled -> skipped at runtime, failures 7 -> 5, no workflow edited"
}
finally {
    & python $pyFile $ConfigXlsx "TRUE"
}

# --- 5. HTML dashboard -----------------------------------------------------
$dashPath = Join-Path $ProjectDir "Data\Output\Reports\dashboard.html"
$dashOk = $false
$dashDetail = "dashboard.html not produced"
if ($run -match "\[DASHBOARD\] dashboard\.html written" -and (Test-Path $dashPath)) {
    $html = Get-Content $dashPath -Raw
    # Self-contained is the point: no CDN, no external stylesheet, no <script src>.
    $selfContained = ($html -notmatch "src=|href=|@import")
    $dashOk = $selfContained -and ($html -match "Transactions validated")
    $dashDetail = "{0:N0} KB, self-contained: {1}" -f ((Get-Item $dashPath).Length / 1KB), $selfContained
}
Add-Result "HTML dashboard generated by the run" $dashOk $dashDetail

# --- 6. Escalation (Novelty 2) --------------------------------------------
Add-Result "CRITICAL anomaly escalated (Novelty 2)" `
    ($run -match "1 anomaly\(ies\) require human review: SHP-1007/Delivered") `
    "unapproved wallet on SHP-1007 routed to human review and dispositioned"

# --- 7. Chain swap (Novelty 3) --------------------------------------------
$swap = & (Join-Path $PSScriptRoot "chain_swap_demo.ps1") 6>&1 2>&1 | Out-String
Add-Result "Identical result on a second chain (Novelty 3)" `
    ($swap -match "Chain agnosticism demonstrated") "Ethereum and Polygon, config change only"

# --- 8. Tamper evidence (Novelty 4) ---------------------------------------
$tamper = & (Join-Path $PSScriptRoot "tamper_test.ps1") 6>&1 2>&1 | Out-String
Add-Result "Audit log tampering detected (Novelty 4)" `
    ($tamper -match "Tamper test PASSED") "edited verdict detected at the exact line, then restored"

# --- 9. WEB extraction mode ------------------------------------------------
$web = Invoke-Build @{ DataSourceMode = "WEB" }
$webLine = ($web -split "`n" | Where-Object { $_ -match "\[EXTRACT/WEB\]" } | Select-Object -Last 1)
Add-Result "WEB mode scrapes the explorer page to the same result" `
    ($web -match [regex]::Escape($expected)) `
    $(if ($webLine) { $webLine.Trim() } else { "no [EXTRACT/WEB] line in the run log" })

# --- 10. UiPath test cases --------------------------------------------------
$tests = & (Join-Path $PSScriptRoot "run_tests.ps1") 6>&1 2>&1 | Out-String
$testLine = ($tests -split "`n" | Where-Object { $_ -match "test cases passed" } | Select-Object -First 1)
Add-Result "UiPath test cases (7 rule/integration tests)" `
    ($tests -match "7/7 test cases passed") $testLine.Trim()

# --- 11. Dispatch mode -----------------------------------------------------
$disp = Invoke-Build @{ RunMode = "DISPATCH" }
$dispLine = ($disp -split "`n" | Where-Object { $_ -match "\[DISPATCH\]" } | Select-Object -First 1)
Add-Result "Dispatch splits the run per shipment" `
    ($disp -match "12 shipment payload\(s\) prepared from 48 transaction\(s\)") `
    $(if ($dispLine) { $dispLine.Trim() } else { "no [DISPATCH] line in the run log" })

# Leave the project back in its default state.
Invoke-Build | Out-Null

# --- Summary ---------------------------------------------------------------
$passed = ($results | Where-Object Ok).Count
$total  = $results.Count
Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($passed -eq $total) {
    Write-Host "$passed/$total checks passed." -ForegroundColor Green
    exit 0
} else {
    Write-Host "$passed/$total checks passed. Failures:" -ForegroundColor Red
    $results | Where-Object { -not $_.Ok } | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Red }
    exit 1
}

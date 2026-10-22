<#
    run_all_checks.ps1 - full acceptance suite for the project.

    Runs every verifiable claim the project makes, end to end, and reports pass/fail.
    Use it before a review or a demo: if this is green, the whole pipeline works.

    Static (no robot run):
         1. Dataset oracle        - the seeded anomalies exist, the ABI decodes, every feed agrees
         2. Invoke Code compile   - every VB code body compiles against the robot's references
         3. Workflow Analyzer     - no Error-severity findings; package builds

    End to end (robot runs, offline, against bundled data):
         4. Batch run             - all six rules fire with the expected counts
         5. Rule toggle           - a rule disabled in Excel is skipped (Novelty 1)
         6. HTML dashboard        - self-contained, with processing time and trends
         7. PDF audit report      - produced, structurally valid (Phase 4)
         8. Escalation            - a CRITICAL anomaly is routed to review (Novelty 2)
         9. Chain swap            - identical results on three feeds, two schemas (Novelty 3)
        10. Tamper evidence       - an edited audit log is detected (Novelty 4)
        11. WEB mode              - explorer-page scraping reaches the same result
        12. ERP from CSV          - LogisticsSourceMode = CSV
        13. ERP from SQL          - LogisticsSourceMode = SQL, Database Connect over ODBC
        14. LOOKUP mapping        - TxHash -> ShipmentID from the mapping file
        15. Alert delivery        - SMTP with attachments, Teams, Slack and dashboard push,
                                    captured by a local stand-in server
        16. Queue mode            - Dispatcher payloads processed by the Performer entry point
        17. Failure handling      - a broken run is logged, screenshotted and re-thrown
        18. UiPath test cases     - every TC*.xaml under Tests\

    Run:  powershell -File tools\run_all_checks.ps1
          powershell -File tools\run_all_checks.ps1 -StaticOnly
#>
param([switch]$StaticOnly)
$ErrorActionPreference = "Stop"

$Root       = Split-Path $PSScriptRoot -Parent
$ProjectDir = Join-Path $Root "BlockchainLogisticsValidator"
$Build      = Join-Path $Root "build.ps1"
$OutputRoot = Join-Path $ProjectDir "Data\Output"
$Audit      = Join-Path $OutputRoot "AuditLogs\ValidationAuditLog.csv"
$Reports    = Join-Path $OutputRoot "Reports"
$ConfigXlsx = Join-Path $ProjectDir "Data\Config.xlsx"
$Expected   = "total=48 pass=40 warn=1 fail=7 escalations=1"

$results = [System.Collections.ArrayList]::new()
function Add-Result([string]$Name, [bool]$Ok, [string]$Detail) {
    [void]$results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail })
    $tag = if ($Ok) { "PASS" } else { "FAIL" }
    $col = if ($Ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1}" -f $tag, $Name) -ForegroundColor $col
    if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Invoke-Build([hashtable]$Set = @{}, [string]$Entry = "") {
    # Write-Host output goes to the information stream; redirect it too, or the
    # build chatter leaks past this capture and buries the suite results.
    $ErrorActionPreference = "Continue"
    if ($Entry) { $out = & $Build -Set $Set -Entry $Entry -LogTail 120 6>&1 2>&1 | Out-String }
    else        { $out = & $Build -Set $Set -LogTail 120 6>&1 2>&1 | Out-String }
    return $out
}

function Get-Line([string]$Text, [string]$Pattern) {
    $l = ($Text -split "`n" | Where-Object { $_ -match $Pattern } | Select-Object -Last 1)
    if ($l) { return $l.Trim() } else { return "" }
}

function Test-PdfStructure([string]$Path) {
    # Mirrors TC09: header, trailer, xref offsets and page tree must agree.
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $text = [System.Text.Encoding]::ASCII.GetString($bytes)
    $problems = @()
    if (-not $text.StartsWith("%PDF-1.4")) { $problems += "header" }
    if (-not $text.EndsWith("%%EOF`n")) { $problems += "trailer" }
    if ($bytes | Where-Object { $_ -gt 127 } | Select-Object -First 1) { $problems += "non-ASCII byte" }
    $sx = $text.LastIndexOf("startxref")
    $xrefAt = [int](($text.Substring($sx + 9).Trim() -split "`n")[0])
    if (-not $text.Substring($xrefAt).StartsWith("xref")) { $problems += "startxref" }
    $lines = $text.Substring($xrefAt) -split "`n"
    $count = [int](($lines[1] -split ' ')[1])
    for ($i = 1; $i -lt $count; $i++) {
        $off = [int]$lines[1 + $i + 1].Substring(0, 10)
        if (-not $text.Substring($off).StartsWith("$i 0 obj")) { $problems += "xref entry $i"; break }
    }
    $pages = ([regex]::Matches($text, "/Type /Page /Parent")).Count
    $declared = [int]([regex]::Match($text, "/Count (\d+)").Groups[1].Value)
    if ($pages -eq 0 -or $pages -ne $declared) { $problems += "page tree ($pages vs $declared)" }
    return [pscustomobject]@{ Ok = ($problems.Count -eq 0); Pages = $pages; Problems = ($problems -join ", ") }
}

Write-Host "`n=== Acceptance suite ===`n" -ForegroundColor Cyan

# External tools report warnings on stderr. Their verdict is read from their output and
# exit code, so stderr must not abort the suite.
$ErrorActionPreference = "Continue"

# --- 1. Dataset oracle ----------------------------------------------------
$oracle = & python (Join-Path $PSScriptRoot "verify_seeded_data.py") 2>&1 | Out-String
Add-Result "Dataset oracle" ($oracle -match "All checks passed") `
    (($oracle -split "`n" | Where-Object { $_ -match "checks passed|CHECK\(S\) FAILED" }) -join "")

# --- 2. Invoke Code compile -------------------------------------------------
$cc = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "compile_check.ps1") 2>&1 | Out-String
Add-Result "Every Invoke Code body compiles" ($LASTEXITCODE -eq 0 -and $cc -match "(\d+)/\1 Invoke Code bodies compile") `
    (Get-Line $cc "Invoke Code bodies compile")

# --- 3. Workflow Analyzer + package -------------------------------------------
$an = & $Build -Analyze -SkipRun 6>&1 2>&1 | Out-String
Add-Result "Workflow Analyzer and package build" ($an -notmatch "\[ERROR\]" -and $an -match "\.nupkg") `
    $(if ($an -match "No Error-severity findings") { "no Error-severity findings; " + (Get-Line $an "\.nupkg") } else { "see output" })

if ($StaticOnly) {
    $passed = ($results | Where-Object Ok).Count
    Write-Host "`n$passed/$($results.Count) static checks passed." -ForegroundColor $(if ($passed -eq $results.Count) { "Green" } else { "Red" })
    exit $(if ($passed -eq $results.Count) { 0 } else { 1 })
}

# --- 4. Batch run: all six rules ------------------------------------------
Remove-Item $Audit -Force -ErrorAction SilentlyContinue
$run = Invoke-Build
Add-Result "Batch run - six rules fire" ($run -match [regex]::Escape($Expected)) (Get-Line $run "total=\d+ pass=")

$allRules = @("R1(TimestampCheck)=2","R2(QuantityMatch)=1","R3(AddressWhitelist)=1",
              "R4(DuplicateDetection)=2","R5(SequenceValidation)=1","R6(SmartContractEvent)=1")
$missing = $allRules | Where-Object { $run -notmatch [regex]::Escape($_) }
Add-Result "Every rule reports its seeded anomaly" ($missing.Count -eq 0) `
    $(if ($missing) { "missing: $($missing -join ', ')" } else { "R1-R6 all present" })

# --- 5. Rule toggle (Novelty 1) -------------------------------------------
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

# --- 6. HTML dashboard -----------------------------------------------------
$dashPath = Join-Path $Reports "dashboard.html"
$dashOk = $false
$dashDetail = "dashboard.html not produced"
if ($run -match "\[DASHBOARD\] dashboard\.html written" -and (Test-Path $dashPath)) {
    $html = Get-Content $dashPath -Raw
    # Self-contained is the point: no CDN, no external stylesheet, no <script src>.
    $selfContained = ($html -notmatch "src=|href=|@import")
    $dashOk = $selfContained -and ($html -match "Transactions validated") -and ($html -match "Processing time") -and ($html -match "Trends across")
    $dashDetail = "{0:N0} KB, self-contained: {1}, processing time and trend sections present: {2}" -f `
        ((Get-Item $dashPath).Length / 1KB), $selfContained, (($html -match "Processing time") -and ($html -match "Trends across"))
}
Add-Result "HTML dashboard with processing time and trends" $dashOk $dashDetail

# --- 7. PDF audit report ----------------------------------------------------
$pdf = Get-ChildItem $Reports -Filter "AuditReport_*.pdf" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($pdf) {
    $ps = Test-PdfStructure $pdf.FullName
    Add-Result "PDF audit report produced and valid" ($ps.Ok -and ($run -match "\[PDF\] PDF audit report written")) `
        $(if ($ps.Ok) { "$($pdf.Name), $($ps.Pages) page(s)" } else { "structure problems: $($ps.Problems)" })
} else {
    Add-Result "PDF audit report produced and valid" $false "no AuditReport_*.pdf in Reports"
}

# --- 8. Escalation (Novelty 2) --------------------------------------------
Add-Result "CRITICAL anomaly escalated (Novelty 2)" `
    ($run -match "1 anomaly\(ies\) require human review: SHP-1007/Delivered") `
    "unapproved wallet on SHP-1007 routed to human review and dispositioned"

# --- 9. Chain swap (Novelty 3) --------------------------------------------
$swap = & (Join-Path $PSScriptRoot "chain_swap_demo.ps1") 6>&1 2>&1 | Out-String
Add-Result "Identical result on three feeds, two schemas (Novelty 3)" `
    ($swap -match "Chain agnosticism demonstrated") "Ethereum, Polygon, and a Besu ethereum-etl export - config change only"

# --- 10. Tamper evidence (Novelty 4) ---------------------------------------
$tamper = & (Join-Path $PSScriptRoot "tamper_test.ps1") 6>&1 2>&1 | Out-String
Add-Result "Audit log tampering detected (Novelty 4)" `
    ($tamper -match "Tamper test PASSED") "edited verdict detected at the exact line, then restored"

# --- 11. WEB extraction mode ------------------------------------------------
$web = Invoke-Build @{ DataSourceMode = "WEB" }
Add-Result "WEB mode scrapes the explorer page to the same result" `
    ($web -match [regex]::Escape($Expected)) (Get-Line $web "\[EXTRACT/WEB\]")

# --- 12-14. ERP sources and mapping mode -------------------------------------
$csv = Invoke-Build @{ LogisticsSourceMode = "CSV" }
Add-Result "ERP read from CSV exports" `
    ($csv -match [regex]::Escape($Expected) -and $csv -match "source CSV") (Get-Line $csv "\[ERP\] \d+ shipments")

$sql = Invoke-Build @{ LogisticsSourceMode = "SQL" }
Add-Result "ERP read from SQL via Database Connect" `
    ($sql -match [regex]::Escape($Expected) -and $sql -match "source SQL") (Get-Line $sql "\[ERP\] \d+ shipments")

$lookup = Invoke-Build @{ MappingMode = "LOOKUP" }
Add-Result "LOOKUP mapping from TxShipmentMap.xlsx" `
    ($lookup -match [regex]::Escape($Expected) -and $lookup -match "LOOKUP: 48 resolved") (Get-Line $lookup "MappingMode LOOKUP:")

# --- 15. Alert delivery, captured locally -------------------------------------
$capture = Join-Path $env:TEMP ("blv-capture-" + [guid]::NewGuid().ToString("N"))
$server = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList `
    "-ExecutionPolicy Bypass -File `"$(Join-Path $PSScriptRoot 'capture_server.ps1')`" -CaptureDir `"$capture`" -TimeoutSeconds 600"
for ($i = 0; $i -lt 50 -and -not (Test-Path (Join-Path $capture "ready")); $i++) { Start-Sleep -Milliseconds 200 }
try {
    $alerts = Invoke-Build @{
        AlertMode = "SMTP"; SmtpHost = "127.0.0.1"; SmtpPort = "2525"; SmtpUseSsl = "False"
        SmtpUser = ""; SmtpPassword = ""
        TeamsWebhookUrl = "http://127.0.0.1:18080/teams"; SlackWebhookUrl = "http://127.0.0.1:18080/slack"
        DashboardPushUrl = "http://127.0.0.1:18080/powerbi"; DashboardPushFormat = "POWERBI"
    }
}
finally {
    New-Item (Join-Path $capture "stop") -ItemType File -Force | Out-Null
    $server.WaitForExit(15000) | Out-Null
}
$eml = Get-ChildItem $capture -Filter "smtp_*.eml" -ErrorAction SilentlyContinue | Select-Object -First 1
$teams = Get-ChildItem $capture -Filter "http_*_teams.txt" -ErrorAction SilentlyContinue | Select-Object -First 1
$slack = Get-ChildItem $capture -Filter "http_*_slack.txt" -ErrorAction SilentlyContinue | Select-Object -First 1
$pbi = Get-ChildItem $capture -Filter "http_*_powerbi.txt" -ErrorAction SilentlyContinue | Select-Object -First 1
$emlText = if ($eml) { Get-Content $eml.FullName -Raw } else { "" }
$smtpOk = $emlText -match "Subject: \[ACTION REQUIRED\]" -and $emlText -match "ValidationReport_.*\.xlsx" -and $emlText -match "AuditReport_.*\.pdf"
$teamsOk = $teams -and ((Get-Content $teams.FullName -Raw) -match "AdaptiveCard")
$slackOk = $slack -and ((Get-Content $slack.FullName -Raw) -match '"blocks"')
$pbiOk = $pbi -and ((Get-Content $pbi.FullName -Raw) -match '"passPercent"')
Add-Result "Alerts delivered: SMTP + attachments, Teams, Slack, dashboard push" `
    ($smtpOk -and $teamsOk -and $slackOk -and $pbiOk) `
    ("smtp={0} teams={1} slack={2} powerbi={3}" -f [bool]$smtpOk, [bool]$teamsOk, [bool]$slackOk, [bool]$pbiOk)

# --- 16. Queue mode: Dispatcher -> Performer -------------------------------------
$disp = Invoke-Build @{ RunMode = "DISPATCH" }
$dispOk = $disp -match "12 shipment payload\(s\) prepared from 48 transaction\(s\)"
$perf = Invoke-Build @{} "Workflows\Queue\Performer.xaml"
$perfOk = $perf -match [regex]::Escape("[PERFORM] $Expected") -and $perf -match "12 shipment\(s\) processed, 0 failed"
Add-Result "Queue mode: dispatch per shipment, Performer validates and reports" ($dispOk -and $perfOk) `
    ("dispatch={0}; performer: {1}" -f [bool]$dispOk, (Get-Line $perf "\[PERFORM\] total="))

# --- 17. Failure handling ----------------------------------------------------------
$shots = Join-Path $OutputRoot "Screenshots"
$before = @(Get-ChildItem $shots -Filter "Failure_*.png" -ErrorAction SilentlyContinue).Count
$broken = Invoke-Build @{ MockChainFile = "Data\Input\MockChain\does_not_exist.json" }
$after = @(Get-ChildItem $shots -Filter "Failure_*.png" -ErrorAction SilentlyContinue).Count
Add-Result "A failed run is logged, screenshotted and re-thrown" `
    ($broken -match "\[MAIN\] Run failed" -and $broken -match "Exit code [1-9]" -and $after -gt $before) `
    (Get-Line $broken "\[MAIN\] Run failed")

# --- 18. UiPath test cases --------------------------------------------------
$tests = & (Join-Path $PSScriptRoot "run_tests.ps1") 6>&1 2>&1 | Out-String
$testLine = Get-Line $tests "test cases passed"
Add-Result "UiPath test cases" ($tests -match "(\d+)/\1 test cases passed") $testLine

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

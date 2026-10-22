<#
    review_window_test.ps1 - end-to-end test of the human-in-the-loop review window.

    Runs the bot with EscalationMode = PROMPT, waits for the review window, and answers it
    through Windows UI Automation the way a person would: reads the facts, checks the
    submit button stays disabled until a decision is chosen, picks REJECTED, types a
    reviewer name and a note, and presses "Record decision and resume". Then checks the
    decision went where it must:

        1. the window appears for the CRITICAL anomaly and shows its facts
        2. submit is disabled until a decision is chosen
        3. the run finishes after the answer, rather than hanging
        4. the decision, note and reviewer are logged
        5. they are hashed into the audit chain, not just logged
        6. the audit chain still verifies with the decision in it

    A window will appear on screen for a few seconds while this runs. Do not click it.

    Run:  powershell -ExecutionPolicy Bypass -File tools\review_window_test.ps1
#>
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$Root     = Split-Path $PSScriptRoot -Parent
$Build    = Join-Path $Root "build.ps1"
$Audit    = Join-Path $Root "BlockchainLogisticsValidator\Data\Output\AuditLogs\ValidationAuditLog.csv"
$RunLog   = Join-Path $env:TEMP "blv-review-window-run.log"
$Choice   = "REJECTED - hold the shipment and open an investigation"
$Reviewer = "UIA Review Tester"
$Note     = "Wallet not on the partner list; verified by automated UI test"

$AE = [Windows.Automation.AutomationElement]
$TS = [Windows.Automation.TreeScope]
$CT = [Windows.Automation.ControlType]

$results = [System.Collections.ArrayList]::new()
function Add-Result([string]$Name, [bool]$Ok, [string]$Detail) {
    [void]$results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok })
    $tag = if ($Ok) { "PASS" } else { "FAIL" }
    $col = if ($Ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1}" -f $tag, $Name) -ForegroundColor $col
    if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Find-All($parent, $controlType) {
    $parent.FindAll($TS::Descendants,
        (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $controlType)))
}

function Find-Window([string]$prefix) {
    foreach ($w in $AE::RootElement.FindAll($TS::Children, [Windows.Automation.Condition]::TrueCondition)) {
        if ($w.Current.Name -like "$prefix*") { return $w }
    }
    return $null
}

Write-Host "`n=== Review window test (a window will appear briefly - do not click it) ===`n" -ForegroundColor Cyan

$auditRowsBefore = if (Test-Path $Audit) { @(Import-Csv $Audit).Count } else { 0 }
$startedAt = Get-Date

# --- Start a PROMPT-mode run in the background ------------------------------------
$run = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList @(
    "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
    "& '$Build' -Set @{EscalationMode='PROMPT'; AttendedMode='True'; ReviewUi='WINDOW'; ReviewTimeoutMinutes='10'} -LogTail 120 *> '$RunLog'")

try {
    # --- 1. Wait for the window ----------------------------------------------------
    $win = $null
    $deadline = (Get-Date).AddSeconds(300)
    while (-not $win -and (Get-Date) -lt $deadline -and -not $run.HasExited) {
        Start-Sleep -Milliseconds 600
        $win = Find-Window "Human review required"
    }
    if (-not $win) { throw "The review window never appeared." }
    Start-Sleep -Milliseconds 800

    $values = @(Find-All $win $CT::Edit | ForEach-Object {
        try { $_.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { "" }
    })
    $shown = $values -join " | "
    Add-Result "Window appears for the CRITICAL anomaly and shows its facts" `
        ($shown -match "SHP-1007" -and $shown -match "Delivered" -and $shown -match "CRITICAL") `
        (($values | Where-Object { $_ -match "SHP-|Delivered|CRITICAL" }) -join " | ")

    # --- 2. Submit stays disabled until a decision is chosen ---------------------------
    $submit = $win.FindFirst($TS::Descendants,
        (New-Object Windows.Automation.PropertyCondition($AE::NameProperty, "Record decision and resume")))
    if (-not $submit) { throw "No 'Record decision and resume' button." }
    $disabledFirst = -not $submit.Current.IsEnabled

    $radio = $win.FindFirst($TS::Descendants,
        (New-Object Windows.Automation.PropertyCondition($AE::NameProperty, $Choice)))
    if (-not $radio) { throw "Option '$Choice' is not on the form." }
    $radio.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
    Start-Sleep -Milliseconds 300

    # The two editable boxes after the read-only facts: reviewer name, then the note.
    $editable = @(Find-All $win $CT::Edit | Where-Object {
        -not $_.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.IsReadOnly })
    if ($editable.Count -lt 2) { throw "Expected the reviewer and note boxes, found $($editable.Count) editable box(es)." }
    $editable[0].GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Reviewer)
    $editable[1].GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Note)
    Start-Sleep -Milliseconds 300

    $enabledAfter = $submit.Current.IsEnabled
    Add-Result "Submit is disabled until a decision is chosen" ($disabledFirst -and $enabledAfter) `
        ("before choosing: {0}; after: {1}" -f $(if ($disabledFirst) { "disabled" } else { "ENABLED" }),
                                              $(if ($enabledAfter) { "enabled" } else { "DISABLED" }))

    $submit.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()

    # --- 3. The run finishes ----------------------------------------------------------
    $finished = $run.WaitForExit(300000)
    Add-Result "Run continues and finishes after the answer" $finished `
        $(if ($finished) { "the bot resumed and completed" } else { "still running after 5 minutes" })
}
finally {
    if (-not $run.HasExited) { Stop-Process -Id $run.Id -Force -ErrorAction SilentlyContinue }
}

$log = if (Test-Path $RunLog) { Get-Content $RunLog -Raw } else { "" }

# --- 4. Logged, with the reviewer and the note ------------------------------------
# Read from the robot's own execution log: console output can cut long lines short.
$execLog = Get-ChildItem (Join-Path $env:LOCALAPPDATA "UiPath\Logs") -Filter "*Execution.log" |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
$line = Get-Content $execLog.FullName | ForEach-Object {
    if ($_ -match '^(\d\d:\d\d:\d\d)\.\d+\s+\w+\s+(\{.*\})$') {
        try { $o = $Matches[2] | ConvertFrom-Json } catch { return }
        if ($o.timeStamp -and ([datetime]$o.timeStamp) -ge $startedAt.AddSeconds(-2) -and
            $o.message -match "^\[ESCALATE\] SHP-1007 ->") { $o.message }
    }
} | Select-Object -Last 1
Add-Result "Decision, note and reviewer are logged" `
    ($line -match "REJECTED" -and $line -match [regex]::Escape($Note) -and $line -match [regex]::Escape("(by $Reviewer)")) `
    $(if ($line) { $line } else { "no [ESCALATE] SHP-1007 line in the execution log" })

# --- 5. Hashed into the audit chain -------------------------------------------------
$entry = if (Test-Path $Audit) {
    @(Import-Csv $Audit) | Select-Object -Skip $auditRowsBefore |
        Where-Object { $_.ShipmentID -eq "SHP-1007" -and $_.EventName -eq "Delivered" } | Select-Object -Last 1
}
Add-Result "Decision is in the audit chain, attributed to the reviewer" `
    ($entry -and $entry.HumanDecision -match "REJECTED" -and $entry.HumanDecidedBy -eq $Reviewer) `
    $(if ($entry) { "HumanDecision='$($entry.HumanDecision)' HumanDecidedBy='$($entry.HumanDecidedBy)'" } else { "no new audit entry for SHP-1007/Delivered" })

# --- 6. The chain still verifies ----------------------------------------------------
Add-Result "Audit chain verifies with the human decision in it" ($log -match "\[VERIFY\] VALID") `
    (($log -split "`n" | Where-Object { $_ -match "\[VERIFY\] (VALID|INVALID)" } | Select-Object -Last 1) -replace '^\s*\S+\s+\S+\s+', '').Trim()

# Restore the default local config for the next ordinary build.
& $Build -SkipRun 6>&1 2>&1 | Out-Null

$passed = ($results | Where-Object Ok).Count
Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($passed -eq $results.Count) {
    Write-Host "$passed/$($results.Count) review window checks passed." -ForegroundColor Green
    exit 0
}
Write-Host "$passed/$($results.Count) review window checks passed." -ForegroundColor Red
exit 1

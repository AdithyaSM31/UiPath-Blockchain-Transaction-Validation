<#
    dialog_test.ps1 - end-to-end test of the human-in-the-loop reviewer dialog.

    Runs the bot with EscalationMode=PROMPT, waits for the reviewer dialog, and answers
    it through Windows UI Automation exactly as a person would: reads it, picks an
    option in the drop-down, presses Ok. Then checks the decision went where it must:

        1. the dialog appears for the CRITICAL anomaly, and names it
        2. the run finishes after the answer, rather than hanging
        3. the decision is logged against the real Windows user
        4. the decision is hashed into the audit chain, not just logged
        5. the audit chain still verifies with the decision in it

    A dialog will appear on screen for a few seconds while this runs.

    Run:  powershell -File tools\dialog_test.ps1
#>
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$Root     = Split-Path $PSScriptRoot -Parent
$Build    = Join-Path $Root "build.ps1"
$Audit    = Join-Path $Root "BlockchainLogisticsValidator\Data\Output\AuditLogs\ValidationAuditLog.csv"
$RunLog   = Join-Path $env:TEMP "blv-dialog-run.log"
$Title    = "Blockchain anomaly requires review"
$Choice   = "REJECTED - hold the shipment and open an investigation"
$Me       = $env:USERNAME

$AE = [Windows.Automation.AutomationElement]
$TS = [Windows.Automation.TreeScope]

$results = [System.Collections.ArrayList]::new()
function Add-Result([string]$Name, [bool]$Ok, [string]$Detail) {
    [void]$results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok })
    $tag = if ($Ok) { "PASS" } else { "FAIL" }
    $col = if ($Ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1}" -f $tag, $Name) -ForegroundColor $col
    if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Find-One($parent, $scope, $prop, $value) {
    $parent.FindFirst($scope, (New-Object Windows.Automation.PropertyCondition($prop, $value)))
}

Write-Host "`n=== Reviewer dialog test (a dialog will appear briefly) ===`n" -ForegroundColor Cyan

$auditLinesBefore = if (Test-Path $Audit) { (Get-Content $Audit).Count } else { 0 }

# --- Start a PROMPT-mode run in the background ------------------------------------
$run = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList @(
    "-NoProfile", "-Command",
    "& '$Build' -Set @{EscalationMode='PROMPT'; AttendedMode='True'} -LogTail 80 *> '$RunLog'")

try {
    # --- 1. Wait for the dialog --------------------------------------------------
    $win = $null
    $deadline = (Get-Date).AddSeconds(240)
    while (-not $win -and (Get-Date) -lt $deadline -and -not $run.HasExited) {
        Start-Sleep -Milliseconds 600
        $win = Find-One $AE::RootElement $TS::Children $AE::NameProperty $Title
    }
    if (-not $win) { throw "The reviewer dialog never appeared." }

    $texts = $win.FindAll($TS::Descendants,
        (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Text)))
    $label = ($texts | ForEach-Object { $_.Current.Name } | Sort-Object Length -Descending | Select-Object -First 1)
    Add-Result "Dialog appears for the CRITICAL anomaly and names it" `
        ($label -match "SHP-1007" -and $label -match "Delivered" -and $label -match "not an approved partner wallet") `
        (($label -split "`n" | Where-Object { $_ -match "Shipment:|Signed by:" }) -join " | ").Trim()

    # --- Answer it the way a person would ----------------------------------------
    $combo = Find-One $win $TS::Descendants $AE::AutomationIdProperty "OptionsCombobox"
    $expand = $combo.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern)
    $expand.Expand()
    Start-Sleep -Milliseconds 400
    $item = Find-One $combo $TS::Descendants $AE::NameProperty $Choice
    if (-not $item) { throw "Option '$Choice' is not in the drop-down." }
    $item.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
    $expand.Collapse()
    Start-Sleep -Milliseconds 300

    $ok = Find-One $win $TS::Descendants $AE::NameProperty "Ok"
    $ok.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()

    # --- 2. The run finishes ------------------------------------------------------
    $finished = $run.WaitForExit(240000)
    Add-Result "Run continues and finishes after the answer" $finished `
        $(if ($finished) { "exit code $($run.ExitCode)" } else { "still running after 4 minutes" })
}
finally {
    if (-not $run.HasExited) { Stop-Process -Id $run.Id -Force -ErrorAction SilentlyContinue }
}

$log = if (Test-Path $RunLog) { Get-Content $RunLog -Raw } else { "" }

# --- 3. Logged against the real user ---------------------------------------------
$line = ($log -split "`n" | Where-Object { $_ -match "\[ESCALATE\] SHP-1007 ->" } | Select-Object -First 1)
Add-Result "Decision logged against the real Windows user" `
    ($line -match [regex]::Escape($Choice) -and $line -match "\(by $([regex]::Escape($Me))\)") `
    ($line -replace '^\s*\S+\s+\S+\s+', '').Trim()

# --- 4. Hashed into the audit chain ----------------------------------------------
$newLines = if (Test-Path $Audit) { Get-Content $Audit | Select-Object -Skip ([Math]::Max($auditLinesBefore, 1)) } else { @() }
$entry = $newLines | Where-Object { $_ -match ",SHP-1007,Delivered," } | Select-Object -Last 1
$fields = if ($entry) { $entry -split "," } else { @() }
Add-Result "Decision is in the audit chain, attributed to the reviewer" `
    ($entry -and $fields[12] -eq $Choice.Replace(",", ";") -and $fields[13] -eq $Me) `
    $(if ($entry) { "HumanDecision='$($fields[12])' HumanDecidedBy='$($fields[13])'" } else { "no new audit entry for SHP-1007/Delivered" })

# --- 5. The chain still verifies -------------------------------------------------
Add-Result "Audit chain verifies with the human decision in it" ($log -match "\[VERIFY\] VALID") `
    (($log -split "`n" | Where-Object { $_ -match "\[VERIFY\] (VALID|INVALID)" } | Select-Object -Last 1) -replace '^\s*\S+\s+\S+\s+', '').Trim()

# Restore the default local config for the next ordinary build.
& $Build -SkipRun 6>&1 2>&1 | Out-Null

$passed = ($results | Where-Object Ok).Count
Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($passed -eq $results.Count) {
    Write-Host "$passed/$($results.Count) dialog checks passed." -ForegroundColor Green
    exit 0
}
Write-Host "$passed/$($results.Count) dialog checks passed." -ForegroundColor Red
exit 1

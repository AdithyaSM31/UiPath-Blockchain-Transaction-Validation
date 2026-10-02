<#
    integration_tests.ps1 - exercise every path that reaches outside the machine.

    Starts tools\mock_services.py (a local Etherscan V2 API, explorer page, webhook
    receiver and SMTP server), runs the bot against it, and asserts on what the
    services actually RECEIVED - not only on what the bot logged about sending it.

        1. API mode            Etherscan request is well formed; verdict matches MOCK
                               + SMTP and Teams webhook delivered in the same run
        2. Bad API key         fails cleanly with Etherscan's own error, not a crash
        3. Second chain        chainid 137 via the API, Slack webhook format
        4. Empty contract      "No transactions found" is an empty run, not an error
        5. WEB over HTTP       explorer page fetched from a server, not the disk

    Run:  powershell -File tools\integration_tests.ps1
#>
$ErrorActionPreference = "Stop"

$Root     = Split-Path $PSScriptRoot -Parent
$Build    = Join-Path $Root "build.ps1"
$Capture  = Join-Path $env:TEMP "blv-mock-capture"
$Contract = "0x0421f2bb09e32F206A76bFAef9950ECAf0ec30A7"
$Base     = "http://127.0.0.1:8765"
$Expected = "total=48 pass=40 warn=1 fail=7 escalations=1"

$results = [System.Collections.ArrayList]::new()
function Add-Result([string]$Name, [bool]$Ok, [string]$Detail) {
    [void]$results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok })
    $tag = if ($Ok) { "PASS" } else { "FAIL" }
    $col = if ($Ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1}" -f $tag, $Name) -ForegroundColor $col
    if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Invoke-Run([hashtable]$Set) {
    if (Test-Path $Capture) { Get-ChildItem $Capture -File | ForEach-Object { [IO.File]::Delete($_.FullName) } }
    $ErrorActionPreference = "Continue"
    $out = & $Build -Set $Set -LogTail 120 6>&1 2>&1 | Out-String
    $code = $LASTEXITCODE
    $ErrorActionPreference = "Stop"
    $received = @()
    if (Test-Path $Capture) {
        $received = Get-ChildItem $Capture -Filter *.json | Sort-Object Name |
                    ForEach-Object { Get-Content $_.FullName -Raw | ConvertFrom-Json }
    }
    return [pscustomobject]@{ Log = $out; Exit = $code; Received = $received }
}

function Get-Line($log, $pattern) {
    ($log -split "`n" | Where-Object { $_ -match $pattern } | Select-Object -Last 1) -replace '^\s*\S+\s+\S+\s+', ''
}

# --- Start the test doubles ----------------------------------------------------
New-Item -ItemType Directory -Force $Capture | Out-Null
$mock = Start-Process python -ArgumentList @("`"$(Join-Path $PSScriptRoot 'mock_services.py')`"", "--capture", "`"$Capture`"") `
                             -PassThru -WindowStyle Hidden
try {
    $up = $false
    for ($i = 0; $i -lt 40 -and -not $up; $i++) {
        try { $null = Invoke-RestMethod "$Base/health" -TimeoutSec 1; $up = $true } catch { Start-Sleep -Milliseconds 250 }
    }
    if (-not $up) { throw "mock_services.py did not come up on $Base" }

    Write-Host "`n=== Integration tests (against local test doubles) ===`n" -ForegroundColor Cyan

    # --- 1. API mode, with SMTP and a Teams webhook in the same run ------------
    $r = Invoke-Run @{
        DataSourceMode = "API"; EtherscanBaseUrl = "$Base/v2/api"; EtherscanApiKey = "TESTKEY"
        AlertMode = "SMTP,WEBHOOK"; SmtpHost = "127.0.0.1"; SmtpPort = "8025"
        AlertWebhookUrl = "$Base/webhook/teams"; AlertWebhookFormat = "TEAMS"
    }
    $req = $r.Received | Where-Object kind -eq "etherscan_request" | Select-Object -First 1
    $reqOk = $req -and $req.query.chainid -eq "1" -and $req.query.apikey -eq "TESTKEY" -and
             $req.query.address -eq $Contract -and $req.query.module -eq "account" -and $req.query.action -eq "txlist"
    Add-Result "API: Etherscan V2 request is well formed" $reqOk `
        $(if ($req) { "chainid=$($req.query.chainid) module=$($req.query.module) action=$($req.query.action) apikey sent" } else { "no request received" })
    Add-Result "API: same verdict as the bundled feed" ($r.Log -match [regex]::Escape($Expected)) (Get-Line $r.Log "total=\d+")
    Add-Result "API: the key never appears in the log" ($r.Log -notmatch "apikey=TESTKEY" -and $r.Log -match "REDACTED") `
        "logged URL is redacted"

    $mail = $r.Received | Where-Object kind -eq "smtp" | Select-Object -First 1
    Add-Result "SMTP: alert delivered to the mail server" `
        ($mail -and $mail.message -match "ACTION REQUIRED" -and $mail.message -match "SHP-1007") `
        $(if ($mail) { "{0} bytes to {1}" -f $mail.bytes, ($mail.to -join ", ") } else { "nothing received" })

    $hook = $r.Received | Where-Object kind -eq "webhook" | Select-Object -First 1
    $card = if ($hook) { $hook.json.attachments[0].content } else { $null }
    Add-Result "Webhook: Teams Adaptive Card delivered" `
        ($hook -and $hook.json.type -eq "message" -and $card.type -eq "AdaptiveCard" -and ($card.body | Where-Object type -eq "FactSet")) `
        $(if ($card) { "{0} card elements, path {1}" -f $card.body.Count, $hook.path } else { "nothing received" })

    # --- 2. A bad API key fails cleanly ---------------------------------------
    $r = Invoke-Run @{ DataSourceMode = "API"; EtherscanBaseUrl = "$Base/v2/api"; EtherscanApiKey = "INVALID" }
    Add-Result "Bad API key: fails with Etherscan's own message" `
        ($r.Exit -ne 0 -and $r.Log -match "Invalid API Key" -and $r.Log -match "status=0") `
        (Get-Line $r.Log "Invalid API Key")

    # --- 3. Second chain through the API, Slack format ------------------------
    $r = Invoke-Run @{
        DataSourceMode = "API"; EtherscanBaseUrl = "$Base/v2/api"; EtherscanApiKey = "TESTKEY"; ChainId = "137"
        AlertMode = "WEBHOOK"; AlertWebhookUrl = "$Base/webhook/slack"; AlertWebhookFormat = "SLACK"
    }
    $req = $r.Received | Where-Object kind -eq "etherscan_request" | Select-Object -First 1
    Add-Result "API: chain 137 selected by config alone, same verdict" `
        ($req.query.chainid -eq "137" -and $r.Log -match [regex]::Escape($Expected)) (Get-Line $r.Log "total=\d+")
    $hook = $r.Received | Where-Object kind -eq "webhook" | Select-Object -First 1
    Add-Result "Webhook: Slack Block Kit delivered" `
        ($hook -and $hook.json.text -match "ACTION REQUIRED" -and ($hook.json.blocks | Where-Object type -eq "header")) `
        $(if ($hook) { "{0} blocks" -f $hook.json.blocks.Count } else { "nothing received" })

    # --- 4. A contract with no transactions -----------------------------------
    $r = Invoke-Run @{
        DataSourceMode = "API"; EtherscanBaseUrl = "$Base/v2/api"; EtherscanApiKey = "TESTKEY"
        ContractAddress = "0x000000000000000000000000000000000000dEaD"
    }
    Add-Result "Empty contract: an empty run, not an error" `
        ($r.Exit -eq 0 -and $r.Log -match "total=0 pass=0" -and $r.Log -match "\[VERIFY\] VALID") `
        (Get-Line $r.Log "total=\d+")

    # --- 5. WEB mode over real HTTP -------------------------------------------
    $r = Invoke-Run @{ DataSourceMode = "WEB"; ExplorerPageUrl = "$Base/explorer/address/$Contract" }
    $pageReq = $r.Received | Where-Object kind -eq "explorer_request" | Select-Object -First 1
    Add-Result "WEB: explorer page fetched over HTTP, same verdict" `
        ($pageReq -and $r.Log -match [regex]::Escape($Expected)) (Get-Line $r.Log "\[EXTRACT/WEB\].*decoded")
}
finally {
    if ($mock -and -not $mock.HasExited) { Stop-Process -Id $mock.Id -Force -ErrorAction SilentlyContinue }
    # Put the default local config back for the next ordinary build.
    & $Build -SkipRun 6>&1 2>&1 | Out-Null
}

$passed = ($results | Where-Object Ok).Count
Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($passed -eq $results.Count) {
    Write-Host "$passed/$($results.Count) integration checks passed." -ForegroundColor Green
    exit 0
}
Write-Host "$passed/$($results.Count) integration checks passed." -ForegroundColor Red
$results | Where-Object { -not $_.Ok } | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Red }
exit 1

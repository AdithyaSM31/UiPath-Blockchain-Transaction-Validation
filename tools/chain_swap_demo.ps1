<#
    chain_swap_demo.ps1 - demonstrate Novelty 3: blockchain agnosticism.

    Runs the bot twice against two different networks. Nothing changes between the runs
    except configuration values - no workflow is edited, and the project is not rebuilt
    from different sources. The extraction layer normalises every explorer response to
    one internal schema, so the validation engine never learns which chain it came from.

        run 1: Ethereum mainnet  (chainid 1)
        run 2: Polygon           (chainid 137)

    Run:  powershell -File tools\chain_swap_demo.ps1
#>
$ErrorActionPreference = "Stop"

$Root       = Split-Path $PSScriptRoot -Parent
$ProjectDir = Join-Path $Root "BlockchainLogisticsValidator"
$LocalCfg   = Join-Path $ProjectDir "Data\Config.local.json"
$OutputRoot = Join-Path $ProjectDir "Data\Output"
$UiRobot    = Join-Path $env:LOCALAPPDATA "Programs\UiPath\Studio\UiRobot.exe"
$PkgOut     = Join-Path $env:TEMP "blv-packages"

function Invoke-Run([hashtable]$Overrides, [string]$Label) {
    $cfg = @{ OutputRoot = $OutputRoot } + $Overrides
    $cfg | ConvertTo-Json | Set-Content -Path $LocalCfg -Encoding utf8

    Get-ChildItem $PkgOut -Filter "*.nupkg" -ErrorAction SilentlyContinue | Remove-Item -Force
    & $UiRobot pack (Join-Path $ProjectDir "project.json") --output $PkgOut 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Pack failed for $Label" }
    $pkg = Get-ChildItem $PkgOut -Filter "*.nupkg" | Select-Object -First 1

    $startedAt = Get-Date
    & $UiRobot execute --file $pkg.FullName 2>&1 | Out-Null

    $logFile = Get-ChildItem (Join-Path $env:LOCALAPPDATA "UiPath\Logs") -Filter "*Execution*.log" |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $lines = @()
    Get-Content $logFile.FullName -Tail 120 | ForEach-Object {
        if ($_ -match '^(\d\d:\d\d:\d\d)\.\d+\s+\w+\s+(\{.*\})$') {
            try { $o = $Matches[2] | ConvertFrom-Json } catch { return }
            if ($o.timeStamp -and ([datetime]$o.timeStamp) -lt $startedAt.AddSeconds(-2)) { return }
            if ($o.message -match '^\[(INIT|EXTRACT/|VALIDATE\] total|AUDIT)') { $lines += $o.message }
        }
    }

    Write-Host "`n=== $Label ===" -ForegroundColor Cyan
    foreach ($l in $lines) { Write-Host "  $l" -ForegroundColor Gray }
    return $lines
}

Write-Host "No workflow is modified between these two runs - only Config values." -ForegroundColor Yellow

$eth = Invoke-Run @{
    ChainId       = "1"
    MockChainFile = "Data\Input\MockChain\etherscan_txlist_response.json"
} "RUN 1 - Ethereum mainnet (chainid 1)"

$poly = Invoke-Run @{
    ChainId       = "137"
    MockChainFile = "Data\Input\MockChain\etherscan_txlist_polygon.json"
} "RUN 2 - Polygon (chainid 137)"

# Restore the plain local config so a later build.ps1 run is unaffected.
@{ OutputRoot = $OutputRoot } | ConvertTo-Json | Set-Content -Path $LocalCfg -Encoding utf8

function Get-Totals($lines) { ($lines | Where-Object { $_ -match 'total=' }) -join "" }

$a = Get-Totals $eth
$b = Get-Totals $poly

Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($a -and $a -eq $b) {
    Write-Host "Identical validation outcome across two networks:" -ForegroundColor Green
    Write-Host "  $a" -ForegroundColor Green
    Write-Host "Chain agnosticism demonstrated - config change only." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Outcomes differed between chains:" -ForegroundColor Yellow
    Write-Host "  ethereum: $a"
    Write-Host "  polygon : $b"
    exit 1
}

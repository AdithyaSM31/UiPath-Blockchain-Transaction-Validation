<#
    chain_swap_demo.ps1 - demonstrate Novelty 3: blockchain agnosticism.

    Runs the bot against three feeds. Nothing changes between the runs except
    configuration values - no workflow is edited, and the project is not rebuilt from
    different sources. The FieldMapping sheet selected in Config.xlsx translates each
    feed into one canonical schema, so the validation engine never learns which chain,
    explorer or exporter the data came from.

        run 1: Ethereum mainnet  (chainid 1)     Etherscan txlist   FieldMapping
        run 2: Polygon           (chainid 137)   Etherscan txlist   FieldMapping
        run 3: Hyperledger Besu  (chainid 1337)  ethereum-etl export, different field
               names and timestamp format        FieldMapping_EvmEtl

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
            if ($o.message -match '^\[(INIT|EXTRACT/|MAP-FIELDS|VALIDATE\] total|AUDIT)') { $lines += $o.message }
        }
    }

    Write-Host "`n=== $Label ===" -ForegroundColor Cyan
    foreach ($l in $lines) { Write-Host "  $l" -ForegroundColor Gray }
    return $lines
}

Write-Host "No workflow is modified between these runs - only Config values." -ForegroundColor Yellow

$eth = Invoke-Run @{
    ChainId       = "1"
    MockChainFile = "Data\Input\MockChain\etherscan_txlist_response.json"
} "RUN 1 - Ethereum mainnet (chainid 1)"

$poly = Invoke-Run @{
    ChainId       = "137"
    MockChainFile = "Data\Input\MockChain\etherscan_txlist_polygon.json"
} "RUN 2 - Polygon (chainid 137)"

$besu = Invoke-Run @{
    ChainId           = "1337"
    MockChainFile     = "Data\Input\MockChain\evm_etl_besu_export.json"
    ChainRecordsPath  = "transactions"
    FieldMappingSheet = "FieldMapping_EvmEtl"
} "RUN 3 - Hyperledger Besu (chainid 1337), ethereum-etl export"

# Restore the plain local config so a later build.ps1 run is unaffected.
@{ OutputRoot = $OutputRoot } | ConvertTo-Json | Set-Content -Path $LocalCfg -Encoding utf8

function Get-Totals($lines) { ($lines | Where-Object { $_ -match 'total=' }) -join "" }

$a = Get-Totals $eth
$b = Get-Totals $poly
$c = Get-Totals $besu

Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($a -and $a -eq $b -and $a -eq $c) {
    Write-Host "Identical validation outcome across three feeds and two schemas:" -ForegroundColor Green
    Write-Host "  $a" -ForegroundColor Green
    Write-Host "Chain agnosticism demonstrated - config change only." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Outcomes differed between chains:" -ForegroundColor Yellow
    Write-Host "  ethereum: $a"
    Write-Host "  polygon : $b"
    Write-Host "  besu    : $c"
    exit 1
}

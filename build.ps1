<#
    build.ps1 - pack, analyze and run the BlockchainLogisticsValidator project headlessly.

    Usage:
        .\build.ps1                 # pack + execute Main.xaml
        .\build.ps1 -Analyze        # also run the Workflow Analyzer
        .\build.ps1 -Entry "Workflows\12_Verify_AuditChain.xaml"
        .\build.ps1 -LogTail 40     # show more of the execution log

    Windows (.NET 8) UiPath projects cannot be run from a loose .xaml by the Robot CLI,
    so this packs to a .nupkg first and executes that. Same path Orchestrator uses.
#>
param(
    [string]$Entry    = "",
    [switch]$Analyze,
    [switch]$SkipRun,
    [int]$LogTail     = 25,
    # Extra Config.local.json overrides for this run, e.g. -Set @{ RunMode = "DISPATCH" }
    [hashtable]$Set   = @{}
)

$ErrorActionPreference = "Stop"

$StudioDir   = Join-Path $env:LOCALAPPDATA "Programs\UiPath\Studio"
$UiRobot     = Join-Path $StudioDir "UiRobot.exe"
$StudioCli   = Join-Path $StudioDir "UiPath.Studio.CommandLine.exe"
$ProjectDir  = Join-Path $PSScriptRoot "BlockchainLogisticsValidator"
$ProjectJson = Join-Path $ProjectDir "project.json"
$OutDir      = Join-Path $env:TEMP "blv-packages"
$outputRoot  = Join-Path $ProjectDir "Data\Output"

if (-not (Test-Path $UiRobot))     { throw "UiRobot.exe not found at $UiRobot" }
if (-not (Test-Path $ProjectJson)) { throw "project.json not found at $ProjectJson" }

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# A previous run killed mid-execution leaves an orphaned executor holding an Excel COM
# server, and the next Excel Application Scope then blocks forever waiting on it. Clear
# any strays first. Only UiPath processes are touched - never the user's own Excel.
Get-Process -Name "UiPath.Executor", "UiPath.Executor.NetCore" -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue

# --- Analyze -------------------------------------------------------------
if ($Analyze) {
    Write-Host "`n=== Workflow Analyzer ===" -ForegroundColor Cyan
    $raw = & $StudioCli analyze --project-path $ProjectJson 2>&1 | Out-String
    # Surface only Error/Warning severities; Info entries are counts and stats.
    $findings = [regex]::Matches($raw, '"[0-9a-f-]{36}-ErrorCode":\s*"([^"]+)"[\s\S]*?"[0-9a-f-]{36}-ErrorSeverity":\s*"([^"]+)"[\s\S]*?"[0-9a-f-]{36}-Description":\s*"([^"]*)"')
    $bad = 0
    foreach ($m in $findings) {
        $code = $m.Groups[1].Value; $sev = $m.Groups[2].Value; $desc = $m.Groups[3].Value
        if ($sev -eq "Error")   { Write-Host "  [ERROR]   $code  $desc" -ForegroundColor Red;    $bad++ }
        elseif ($sev -eq "Warning") { Write-Host "  [warn]    $code  $desc" -ForegroundColor Yellow }
    }
    if ($bad -eq 0) { Write-Host "  No Error-severity findings." -ForegroundColor Green }
}

# --- Pin the output location ---------------------------------------------
# `UiRobot execute` hands the job to the Robot service, which does NOT inherit this
# shell's environment - so an env var set here never reaches the workflow. Config.local.json
# travels inside the package instead, and 00_Init_ReadConfig reads it. It is uncommitted
# and machine-specific by design; Config.xlsx stays portable.
Write-Host "`n=== Configure ===" -ForegroundColor Cyan
$localCfgPath = Join-Path $ProjectDir "Data\Config.local.json"
$localCfg = @{ OutputRoot = $outputRoot }
foreach ($k in $Set.Keys) { $localCfg[$k] = $Set[$k]; Write-Host "  override: $k = $($Set[$k])" -ForegroundColor DarkGray }
$localCfg | ConvertTo-Json | Set-Content -Path $localCfgPath -Encoding utf8
Write-Host "  output root: $outputRoot" -ForegroundColor DarkGray

# --- Pack ----------------------------------------------------------------
Write-Host "`n=== Pack ===" -ForegroundColor Cyan
Get-ChildItem $OutDir -Filter "*.nupkg" -ErrorAction SilentlyContinue | Remove-Item -Force
$packOut = & $UiRobot pack $ProjectJson --output $OutDir 2>&1 | Out-String
if ($LASTEXITCODE -ne 0) { Write-Host $packOut -ForegroundColor Red; throw "Pack failed." }

$pkg = Get-ChildItem $OutDir -Filter "*.nupkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $pkg) { throw "Pack reported success but produced no .nupkg" }
Write-Host ("  {0}  ({1:N0} bytes)" -f $pkg.Name, $pkg.Length) -ForegroundColor Green

if ($SkipRun) { return }

# --- Run -----------------------------------------------------------------
Write-Host "`n=== Execute ===" -ForegroundColor Cyan
$startedAt = Get-Date

$args = @("execute", "--file", $pkg.FullName)
if ($Entry) { $args += @("--entry", $Entry); Write-Host "  entry point: $Entry" }

$runOut = & $UiRobot @args 2>&1 | Out-String
$exit = $LASTEXITCODE
if ($runOut.Trim()) { Write-Host $runOut.Trim() }
if ($exit -eq 0) { Write-Host "  Exit code 0" -ForegroundColor Green }
else             { Write-Host "  Exit code $exit" -ForegroundColor Red }

# --- Execution log -------------------------------------------------------
Write-Host "`n=== Execution log ===" -ForegroundColor Cyan
$logFile = Get-ChildItem (Join-Path $env:LOCALAPPDATA "UiPath\Logs") -Filter "*Execution*.log" -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($logFile) {
    Get-Content $logFile.FullName -Tail 400 | ForEach-Object {
        # Log lines are "HH:mm:ss.ffff <Level> {json}". Pull out message + level.
        if ($_ -match '^(\d\d:\d\d:\d\d)\.\d+\s+\w+\s+(\{.*\})$') {
            $ts = $Matches[1]
            try { $o = $Matches[2] | ConvertFrom-Json } catch { return }
            if ($o.timeStamp -and ([datetime]$o.timeStamp) -lt $startedAt.AddSeconds(-2)) { return }
            $lvl = $o.level
            $colour = switch ($lvl) { "Error" {"Red"} "Warn" {"Yellow"} "Warning" {"Yellow"} default {"Gray"} }
            Write-Host ("  {0}  {1,-11} {2}" -f $ts, $lvl, $o.message) -ForegroundColor $colour
        }
    } | Select-Object -Last $LogTail
} else {
    Write-Host "  (no execution log found)" -ForegroundColor DarkGray
}

exit $exit

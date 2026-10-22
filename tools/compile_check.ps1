<#
    compile_check.ps1 - compile every Invoke Code body in the project, without running it.

    UiPath compiles an Invoke Code activity's VB source only when the activity executes,
    so a typo in a branch that a test run never reaches (SMTP delivery, the reviewer
    window, Orchestrator queue mode) would otherwise stay hidden until the day it matters.
    `UiRobot pack` already compiles every XAML expression; this script closes the other
    half by compiling each Invoke Code body the way the robot does:

      * the workflow's own Imports list (TextExpression.NamespacesForImplementation)
      * one Sub whose parameters are the activity's arguments - In by value, Out and
        InOut by reference - typed from their x:TypeArguments
      * Option Strict Off / Option Infer On, as Invoke Code uses
      * the .NET 8 + Windows Desktop reference assemblies that ship with Studio, plus
        Newtonsoft.Json and ClosedXML from the NuGet cache

    It uses the Roslyn VB compiler bundled with Studio, needs no network and loads nothing
    from the project into memory - it only compiles. That also makes it usable on a machine
    where Smart App Control blocks the robot from loading freshly built workflow assemblies.

    Run:  powershell -File tools\compile_check.ps1
    Exit code 0 = every Invoke Code body compiles.
#>
param(
    [string]$ProjectDir = (Join-Path (Split-Path $PSScriptRoot -Parent) "BlockchainLogisticsValidator")
)
$ErrorActionPreference = "Stop"

$Studio = Join-Path $env:LOCALAPPDATA "Programs\UiPath\Studio"
$RoslynDir = Join-Path $Studio "net472"
$RefsDir = Join-Path $Studio "refs"
$NuGet = Join-Path $env:USERPROFILE ".nuget\packages"

if (-not (Test-Path (Join-Path $RoslynDir "Microsoft.CodeAnalysis.VisualBasic.dll"))) {
    throw "Roslyn VB compiler not found under $RoslynDir - is UiPath Studio installed?"
}

# ---- Load Roslyn (net472 build) with its dependencies resolved from Studio ------------
# The handler uses plain .NET calls only. Cmdlets such as Join-Path or Test-Path can
# themselves raise AssemblyResolve (for satellite .resources assemblies), and a handler
# that calls them recurses until the process dies with a StackOverflowException.
$script:resolving = New-Object 'System.Collections.Generic.Dictionary[string,System.Reflection.Assembly]'
$resolver = [System.ResolveEventHandler] {
    param($sender, $e)
    $name = $e.Name.Split(',')[0]
    if ($name.EndsWith(".resources")) { return $null }
    if ($script:resolving.ContainsKey($name)) { return $script:resolving[$name] }
    $candidate = [System.IO.Path]::Combine($RoslynDir, $name + ".dll")
    if (-not [System.IO.File]::Exists($candidate)) { return $null }
    $script:resolving[$name] = $null
    $asm = [System.Reflection.Assembly]::LoadFrom($candidate)
    $script:resolving[$name] = $asm
    return $asm
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $RoslynDir "System.Collections.Immutable.dll"))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $RoslynDir "System.Reflection.Metadata.dll"))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $RoslynDir "Microsoft.CodeAnalysis.dll"))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $RoslynDir "Microsoft.CodeAnalysis.VisualBasic.dll"))

# ---- Reference set ----------------------------------------------------------------------
function Get-PackageDll([string]$id, [string]$dll) {
    $root = Join-Path $NuGet $id
    if (-not (Test-Path $root)) { return $null }
    $version = Get-ChildItem $root -Directory | Sort-Object { [version]($_.Name -replace '[^0-9.].*$', '') } -Descending | Select-Object -First 1
    foreach ($tfm in @("net8.0", "net6.0", "netstandard2.1", "netstandard2.0")) {
        $p = Join-Path $version.FullName "lib\$tfm\$dll"
        if (Test-Path $p) { return $p }
    }
    return $null
}

$refPaths = @(Get-ChildItem $RefsDir -Filter *.dll | ForEach-Object FullName)
foreach ($pkg in @(@("newtonsoft.json", "Newtonsoft.Json.dll"),
                   @("closedxml", "ClosedXML.dll"),
                   @("documentformat.openxml", "DocumentFormat.OpenXml.dll"),
                   @("documentformat.openxml.framework", "DocumentFormat.OpenXml.Framework.dll"))) {
    $p = Get-PackageDll $pkg[0] $pkg[1]
    if ($p) { $refPaths += $p }
}
$references = New-Object 'System.Collections.Generic.List[Microsoft.CodeAnalysis.MetadataReference]'
foreach ($p in $refPaths) {
    try { $references.Add([Microsoft.CodeAnalysis.MetadataReference]::CreateFromFile($p)) } catch { }
}

# ---- XAML argument types -> VB types -----------------------------------------------------
$simple = @{
    "x:String" = "String"; "x:Object" = "Object"; "x:Int32" = "Integer"; "x:Int64" = "Long"
    "x:Boolean" = "Boolean"; "x:Double" = "Double"; "x:Decimal" = "Decimal"
    "s:DateTime" = "System.DateTime"; "s:TimeSpan" = "System.TimeSpan"; "s:String[]" = "String()"
    "sd:DataTable" = "System.Data.DataTable"; "sd:DataRow" = "System.Data.DataRow"
    "njl:JObject" = "Newtonsoft.Json.Linq.JObject"; "njl:JArray" = "Newtonsoft.Json.Linq.JArray"
    "njl:JToken" = "Newtonsoft.Json.Linq.JToken"
}
function ConvertTo-VbType([string]$t) {
    $t = $t.Trim()
    if ($simple.ContainsKey($t)) { return $simple[$t] }
    if ($t -match '^scg:(\w+)\((.*)\)$') {
        $args = @()
        $depth = 0; $cur = ""
        foreach ($ch in $Matches[2].ToCharArray()) {
            if ($ch -eq '(') { $depth++ } elseif ($ch -eq ')') { $depth-- }
            if ($ch -eq ',' -and $depth -eq 0) { $args += $cur; $cur = "" } else { $cur += $ch }
        }
        $args += $cur
        return "System.Collections.Generic.$($Matches[1])(Of " + (($args | ForEach-Object { ConvertTo-VbType $_ }) -join ", ") + ")"
    }
    throw "compile_check: no VB mapping for argument type '$t'"
}

$vbOptions = New-Object Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilationOptions(
    [Microsoft.CodeAnalysis.OutputKind]::DynamicallyLinkedLibrary)
$vbOptions = $vbOptions.WithOptionStrict([Microsoft.CodeAnalysis.VisualBasic.OptionStrict]::Off).WithOptionInfer($true).WithOptionExplicit($true)
$parseOptions = New-Object Microsoft.CodeAnalysis.VisualBasic.VisualBasicParseOptions([Microsoft.CodeAnalysis.VisualBasic.LanguageVersion]::Latest)

$nsX = "http://schemas.microsoft.com/winfx/2006/xaml"
$files = Get-ChildItem $ProjectDir -Recurse -Filter *.xaml | Where-Object { $_.FullName -notmatch '\\\.(local|settings|objects|tmh)\\' } | Sort-Object FullName

$total = 0; $failed = 0
Write-Host "`n=== Invoke Code compile check ===`n" -ForegroundColor Cyan
foreach ($f in $files) {
    [xml]$doc = Get-Content $f.FullName -Raw
    $imports = @($doc.GetElementsByTagName("x:String") | Where-Object { $_.ParentNode.ParentNode.LocalName -eq "TextExpression.NamespacesForImplementation" } | ForEach-Object { $_.InnerText.Trim() })
    $codeNodes = @($doc.GetElementsByTagName("*") | Where-Object { $_.LocalName -eq "InvokeCode" })
    foreach ($node in $codeNodes) {
        $total++
        $display = $node.GetAttribute("DisplayName")
        $code = $node.GetAttribute("Code")
        $params = @()
        foreach ($arg in @($node.GetElementsByTagName("*") | Where-Object { $_.LocalName -match '^(In|Out|InOut)Argument$' })) {
            $key = $arg.GetAttribute("Key", $nsX)
            $type = ConvertTo-VbType ($arg.GetAttribute("TypeArguments", $nsX))
            # Invoke Code passes In arguments by value and Out/InOut by reference - the
            # robot's stack traces show Run(String in_X) for an In argument.
            $passing = if ($arg.LocalName -eq "InArgument") { "ByVal" } else { "ByRef" }
            $params += "$passing $key As $type"
        }
        $relName = $f.FullName.Substring($ProjectDir.Length + 1)
        $header = "Option Strict Off`r`nOption Infer On`r`nOption Explicit On`r`n" +
                  (($imports | ForEach-Object { "Imports $_" }) -join "`r`n") + "`r`n" +
                  "Public Module CodeRunner`r`n    Public Sub Run(" + ($params -join ", ") + ")`r`n"
        $source = $header + $code + "`r`n    End Sub`r`nEnd Module`r`n"
        $offset = ($header -split "`r`n").Count - 1

        $tree = [Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxTree]::ParseText($source, $parseOptions, "", $null, [System.Threading.CancellationToken]::None)
        $trees = New-Object 'System.Collections.Generic.List[Microsoft.CodeAnalysis.SyntaxTree]'
        $trees.Add($tree)
        $comp = [Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilation]::Create("InvokeCodeCheck", $trees, $references, $vbOptions)
        $ms = New-Object System.IO.MemoryStream
        $result = $comp.Emit($ms)
        $errors = @($result.Diagnostics | Where-Object { $_.Severity -eq [Microsoft.CodeAnalysis.DiagnosticSeverity]::Error })
        if ($errors.Count -eq 0) {
            Write-Host ("  [ OK ] {0} :: {1}" -f $relName, $display) -ForegroundColor Green
        } else {
            $failed++
            Write-Host ("  [FAIL] {0} :: {1}" -f $relName, $display) -ForegroundColor Red
            foreach ($d in ($errors | Select-Object -First 8)) {
                $line = $d.Location.GetLineSpan().StartLinePosition.Line - $offset + 1
                Write-Host ("         code line {0}: {1} {2}" -f $line, $d.Id, $d.GetMessage()) -ForegroundColor DarkGray
            }
        }
    }
}

Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkGray
if ($failed -eq 0) {
    Write-Host "$total/$total Invoke Code bodies compile." -ForegroundColor Green
    exit 0
} else {
    Write-Host "$($total - $failed)/$total Invoke Code bodies compile. $failed failed." -ForegroundColor Red
    exit 1
}

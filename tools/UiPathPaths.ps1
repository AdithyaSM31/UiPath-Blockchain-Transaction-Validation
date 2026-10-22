<#
    UiPathPaths.ps1 - locate the installed UiPath Studio and Robot, whichever layout
    this machine uses. Dot-source it; it sets three script-scope variables:

        $UiPathStudioDir   the Studio folder (holds net472\, refs\, Packages\)
        $UiPathRobotExe    UiRobot.exe
        $UiPathStudioCli   UiPath.Studio.CommandLine.exe

    Two layouts exist in the wild and we support both:

      legacy    %LOCALAPPDATA%\Programs\UiPath\Studio\            Studio and Robot together,
                                                                  unversioned (Studio <= 25.10)
      platform  %LOCALAPPDATA%\Programs\UiPathPlatform\Studio\<version>\
                %LOCALAPPDATA%\Programs\UiPathPlatform\Robot\<version>\
                                                                  split and versioned, which is
                                                                  what the Connected installer
                                                                  produces (Studio >= 26.0)

    Hardcoding the legacy path meant every script here broke the moment the machine was
    upgraded to 26.x, with a "UiRobot.exe not found" that looks like UiPath is missing
    when it is merely somewhere else.

    IMPORTANT - the resolution below uses plain .NET calls only, no cmdlets.

    compile_check.ps1 dot-sources this file and then loads the net472 Roslyn out of the
    Studio folder, installing an AssemblyResolve handler to satisfy Roslyn's dependencies.
    Cmdlets used *before* that load (Get-ChildItem, Sort-Object, Select-Object) pull
    Microsoft.PowerShell.Commands.Utility and its dependencies into the AppDomain first;
    Roslyn then asks for its own System.Collections.Immutable, the already-loaded copy
    does not satisfy the request, and the handler re-enters until the process dies with a
    StackOverflowException. Sticking to System.IO and System.Environment keeps the
    AppDomain in the state compile_check expects. Do not "tidy" these into cmdlets.
#>

function Get-UiPathNewestVersionDir([string]$uipPathParent) {
    # Newest version subfolder of $uipPathParent, by the leading numeric part of its name
    # ("26.0.204-cloud.25540" -> 26.0.204). Returns $null if there are none.
    if (-not [System.IO.Directory]::Exists($uipPathParent)) { return $null }

    $uipPathBestDir = $null
    $uipPathBestVer = $null
    foreach ($uipPathDir in [System.IO.Directory]::GetDirectories($uipPathParent)) {
        $uipPathName = [System.IO.Path]::GetFileName($uipPathDir)

        # Take the leading digits-and-dots run without Regex, so no extra assembly loads.
        $uipPathEnd = 0
        while ($uipPathEnd -lt $uipPathName.Length -and
               ([char]::IsDigit($uipPathName[$uipPathEnd]) -or $uipPathName[$uipPathEnd] -eq '.')) {
            $uipPathEnd++
        }
        $uipPathHead = $uipPathName.Substring(0, $uipPathEnd).TrimEnd('.')
        if ($uipPathHead.Length -eq 0) { continue }

        $uipPathVer = $null
        try { $uipPathVer = [version]$uipPathHead } catch { continue }

        if ($null -eq $uipPathBestVer -or $uipPathVer -gt $uipPathBestVer) {
            $uipPathBestVer = $uipPathVer
            $uipPathBestDir = $uipPathDir
        }
    }
    return $uipPathBestDir
}

# Every internal variable here is prefixed "uipPath" on purpose. This file is
# dot-sourced, so it shares scope with its caller - a plain $root or $studio would
# silently overwrite the caller's own variable of that name (PowerShell variable names
# are case-insensitive, so $root clobbers $Root too).
$uipPathRoots = @(
    $env:LOCALAPPDATA,
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)}
)

$UiPathStudioDir = $null
$UiPathRobotExe  = $null

foreach ($uipPathRoot in $uipPathRoots) {
    if (-not $uipPathRoot) { continue }

    # --- platform layout (preferred: it is the newer installer) ---
    $uipPathPlatform = [System.IO.Path]::Combine($uipPathRoot, "Programs\UiPathPlatform")
    if (-not [System.IO.Directory]::Exists($uipPathPlatform)) {
        $uipPathPlatform = [System.IO.Path]::Combine($uipPathRoot, "UiPathPlatform")
    }
    if ([System.IO.Directory]::Exists($uipPathPlatform)) {
        $uipPathStudioCandidate = Get-UiPathNewestVersionDir ([System.IO.Path]::Combine($uipPathPlatform, "Studio"))
        $uipPathRobotCandidate  = Get-UiPathNewestVersionDir ([System.IO.Path]::Combine($uipPathPlatform, "Robot"))
        if ($uipPathStudioCandidate -and
            [System.IO.File]::Exists([System.IO.Path]::Combine($uipPathStudioCandidate, "UiPath.Studio.CommandLine.exe"))) {
            $UiPathStudioDir = $uipPathStudioCandidate
            if ($uipPathRobotCandidate -and
                [System.IO.File]::Exists([System.IO.Path]::Combine($uipPathRobotCandidate, "UiRobot.exe"))) {
                $UiPathRobotExe = [System.IO.Path]::Combine($uipPathRobotCandidate, "UiRobot.exe")
            }
            break
        }
    }

    # --- legacy layout ---
    $uipPathLegacy = [System.IO.Path]::Combine($uipPathRoot, "Programs\UiPath\Studio")
    if (-not [System.IO.Directory]::Exists($uipPathLegacy)) {
        $uipPathLegacy = [System.IO.Path]::Combine($uipPathRoot, "UiPath\Studio")
    }
    if ([System.IO.File]::Exists([System.IO.Path]::Combine($uipPathLegacy, "UiRobot.exe"))) {
        $UiPathStudioDir = $uipPathLegacy
        $UiPathRobotExe  = [System.IO.Path]::Combine($uipPathLegacy, "UiRobot.exe")
        break
    }
}

$UiPathStudioCli = $null
if ($UiPathStudioDir) {
    $UiPathStudioCli = [System.IO.Path]::Combine($UiPathStudioDir, "UiPath.Studio.CommandLine.exe")
}

function Assert-UiPathRobot {
    if (-not $UiPathRobotExe -or -not [System.IO.File]::Exists($UiPathRobotExe)) {
        throw @"
UiRobot.exe not found. Looked for both installation layouts under:
  %LOCALAPPDATA%\Programs\UiPathPlatform\Robot\<version>\UiRobot.exe   (Studio 26+)
  %LOCALAPPDATA%\Programs\UiPath\Studio\UiRobot.exe                    (Studio 25.10 and earlier)
Install UiPath Studio, or set `$UiPathRobotExe before dot-sourcing this file.
"@
    }
}

function Assert-UiPathStudio {
    if (-not $UiPathStudioDir -or -not [System.IO.Directory]::Exists($UiPathStudioDir)) {
        throw "UiPath Studio installation not found (looked for the UiPathPlatform and legacy layouts under %LOCALAPPDATA%\Programs and Program Files)."
    }
}

function Invoke-UiPathPack {
<#
    Pack a project to a .nupkg in $OutputDir, returning the packing tool's output.

    `UiRobot pack` is the original path and still the first choice. In the split
    "UiPathPlatform" layout it fails with "Could not find Studio Command-Line
    application!", because UiRobot delegates packing to the Studio CLI and only looks
    for it beside itself - and Studio now lives in a sibling versioned folder. The
    Studio CLI publishes to a local folder itself, which produces the same package, so
    fall back to that rather than reaching into the UiPath installation.
#>
    param(
        [Parameter(Mandatory)] [string] $ProjectJson,
        [Parameter(Mandatory)] [string] $OutputDir
    )
    Assert-UiPathRobot
    [void][System.IO.Directory]::CreateDirectory($OutputDir)

    $uipPathPrevEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $uipPathOut = & $UiPathRobotExe pack $ProjectJson --output $OutputDir 2>&1 | Out-String
    $uipPathCode = $LASTEXITCODE
    $ErrorActionPreference = $uipPathPrevEap

    if ($uipPathCode -eq 0) { return $uipPathOut }

    if ($uipPathOut -notmatch "Could not find Studio Command-Line application") {
        $global:LASTEXITCODE = $uipPathCode
        return $uipPathOut
    }

    Assert-UiPathStudio
    $uipPathPrevEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $uipPathOut2 = & $UiPathStudioCli publish -p $ProjectJson -g Custom -f $OutputDir -o Process -t 600 2>&1 | Out-String
    $uipPathCode2 = $LASTEXITCODE
    $ErrorActionPreference = $uipPathPrevEap

    $global:LASTEXITCODE = $uipPathCode2
    return ($uipPathOut + $uipPathOut2)
}

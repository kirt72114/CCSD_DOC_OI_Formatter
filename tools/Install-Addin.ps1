<#
.SYNOPSIS
  Installs USAF_OI_Formatter.dotm into the user's Word STARTUP folder.

.DESCRIPTION
  Copying the template into %APPDATA%\Microsoft\Word\STARTUP makes Word
  auto-load it at every launch, which is what wires up the "USAF OI
  Formatter" toolbar buttons defined in ThisDocument.cls.

  If the template hasn't been built yet, this script will invoke
  build\build-dotm.ps1 first.

.PARAMETER TemplatePath
  Optional explicit path to USAF_OI_Formatter.dotm.

.PARAMETER Force
  Overwrite any existing copy in STARTUP without prompting.
#>
[CmdletBinding()]
param(
    [string]$TemplatePath = "",
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $PSCommandPath
$repoRoot = Resolve-Path (Join-Path $here '..')

if (-not $TemplatePath) {
    $TemplatePath = Join-Path $repoRoot 'build\USAF_OI_Formatter.dotm'
}

if (-not (Test-Path $TemplatePath)) {
    Write-Host "Template not found, building..."
    & (Join-Path $repoRoot 'build\build-dotm.ps1')
}

if (-not (Test-Path $TemplatePath)) {
    throw "Build did not produce $TemplatePath"
}

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
if (-not (Test-Path $startup)) {
    New-Item -ItemType Directory -Path $startup -Force | Out-Null
}

$dest = Join-Path $startup 'USAF_OI_Formatter.dotm'
if ((Test-Path $dest) -and -not $Force) {
    $ans = Read-Host "$dest exists. Overwrite? (y/N)"
    if ($ans -notmatch '^[Yy]') {
        Write-Host "Cancelled."
        exit 0
    }
}

Copy-Item -Path $TemplatePath -Destination $dest -Force
Write-Host "Installed: $dest"
Write-Host "Restart Word. A 'USAF OI Formatter' toolbar will appear."

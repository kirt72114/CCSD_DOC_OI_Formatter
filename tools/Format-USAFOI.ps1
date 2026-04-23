<#
.SYNOPSIS
  Headless driver for the USAF OI Formatter Word add-in.

.DESCRIPTION
  Starts Word via COM, loads USAF_OI_Formatter.dotm as a global template,
  and runs the formatter on one .docx or all .docx under a folder. Emits
  per-file status and an exit code (0 all-OK, 1 any-failure).

.PARAMETER Path
  A single .docx file OR a directory of .docx files.

.PARAMETER Recurse
  When Path is a directory, recurse into subdirectories.

.PARAMETER OutputDir
  Optional override for the output location. Default: next to source
  with "_formatted" suffix.

.PARAMETER OPR, OIName, Date, Subject, Unit, Category, Supersedes, CertifiedBy, Pages, Accessibility, Releasability
  Metadata for the DAFMAN 90-161 title block.

.EXAMPLE
  Format-USAFOI.ps1 -Path C:\incoming\MyOI.docx `
    -OPR 'CCSD/CCC' -OIName 'CCSD OI 36-1' `
    -Date '23 April 2026' -Subject 'Personnel Actions'

.EXAMPLE
  Format-USAFOI.ps1 -Path C:\incoming -Recurse -OutputDir C:\out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [switch]$Recurse,

    [string]$OutputDir = "",

    [string]$Unit = "",
    [string]$UnitShort = "",
    [string]$OIName = "",
    [string]$Date = "",
    [string]$Category = "",
    [string]$Subject = "",
    [string]$OPR = "",
    [string]$Supersedes = "",
    [string]$CertifiedBy = "",
    [string]$Pages = "",
    [string]$Accessibility = "",
    [string]$Releasability = "",

    [string]$TemplatePath = ""
)

$ErrorActionPreference = 'Stop'

function Resolve-Template {
    param([string]$Explicit)

    if ($Explicit -and (Test-Path $Explicit)) { return (Resolve-Path $Explicit).Path }

    $here = Split-Path -Parent $PSCommandPath
    $candidates = @(
        (Join-Path $here '..\build\USAF_OI_Formatter.dotm'),
        (Join-Path $here '..\USAF_OI_Formatter.dotm'),
        (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\USAF_OI_Formatter.dotm')
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return (Resolve-Path $c).Path }
    }
    throw "USAF_OI_Formatter.dotm not found. Pass -TemplatePath or run build\build-dotm.ps1."
}

function Build-MetaKvp {
    $pairs = @(
        "unit=$Unit",
        "unitshort=$UnitShort",
        "oinumber=$OIName",
        "date=$Date",
        "category=$Category",
        "subject=$Subject",
        "opr=$OPR",
        "supersedes=$Supersedes",
        "certifiedby=$CertifiedBy",
        "pages=$Pages",
        "accessibility=$Accessibility",
        "releasability=$Releasability"
    )
    return ($pairs -join ';;')
}

function Invoke-Formatter {
    param(
        [object]$Word,
        [string]$Target,
        [string]$MetaKvp,
        [bool]$IsFolder
    )

    if ($IsFolder) {
        $recurse01 = if ($Recurse) { '1' } else { '0' }
        $Word.Run('modCLI.FormatFolder', [ref]$Target, [ref]$recurse01,
                  [ref]$OutputDir, [ref]$MetaKvp) | Out-Null
    } else {
        $Word.Run('modCLI.FormatOne', [ref]$Target, [ref]$MetaKvp) | Out-Null
    }
}

$templatePath = Resolve-Template -Explicit $TemplatePath
Write-Host "Using template: $templatePath"

$resolved = Resolve-Path $Path
$isFolder = (Get-Item $resolved).PSIsContainer
$metaKvp = Build-MetaKvp

$word = $null
$exitCode = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $word.AddIns.Add($templatePath, $true) | Out-Null

    Invoke-Formatter -Word $word -Target $resolved.Path -MetaKvp $metaKvp `
                     -IsFolder:$isFolder
    Write-Host "Done: $resolved"
}
catch {
    Write-Error $_
    $exitCode = 1
}
finally {
    if ($word) {
        try { $word.Quit([ref]0) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

exit $exitCode

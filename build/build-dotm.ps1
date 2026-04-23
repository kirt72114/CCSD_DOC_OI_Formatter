<#
.SYNOPSIS
  Assembles USAF_OI_Formatter.dotm from the plain-text VBA sources in src\vba.

.DESCRIPTION
  Requires Word installed locally. Creates a fresh .dotm, enables trust
  access to the VBA project (one-time manual step if not already set),
  then imports every .bas and .frm from src\vba. The result is saved to
  build\USAF_OI_Formatter.dotm.

  Trust access requirement:
    File -> Options -> Trust Center -> Trust Center Settings ->
    Macro Settings -> "Trust access to the VBA project object model"
    must be CHECKED. This can't be flipped programmatically; it is a
    one-time per-user setting.

.PARAMETER OutputPath
  Override the output .dotm location.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = ""
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $PSCommandPath
$repoRoot = Resolve-Path (Join-Path $here '..')
$srcDir = Join-Path $repoRoot 'src\vba'

if (-not $OutputPath) {
    $OutputPath = Join-Path $repoRoot 'build\USAF_OI_Formatter.dotm'
}

if (-not (Test-Path $srcDir)) {
    throw "VBA source directory not found: $srcDir"
}

$bas = Get-ChildItem $srcDir -Filter '*.bas'
$frm = Get-ChildItem $srcDir -Filter '*.frm'
$cls = Get-ChildItem $srcDir -Filter '*.cls'

if (-not $bas) { throw "No .bas files in $srcDir" }

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()

    # Force .dotm by saving first with the macro-enabled template format.
    $wdFormatXMLTemplateMacroEnabled = 15
    $doc.SaveAs([ref]$OutputPath, [ref]$wdFormatXMLTemplateMacroEnabled)

    $vbproj = $doc.VBProject
    if (-not $vbproj) {
        throw "Could not access VBA project. Enable 'Trust access to the VBA project object model' in Word Trust Center."
    }

    # Import standard modules.
    foreach ($f in $bas) {
        Write-Host "Import module: $($f.Name)"
        $vbproj.VBComponents.Import($f.FullName) | Out-Null
    }

    # Import UserForms.
    foreach ($f in $frm) {
        Write-Host "Import form:   $($f.Name)"
        $vbproj.VBComponents.Import($f.FullName) | Out-Null
    }

    # ThisDocument.cls has to replace the existing ThisDocument module's
    # *code*, not be imported as a new class. Read and paste its body.
    foreach ($f in $cls) {
        if ($f.BaseName -eq 'ThisDocument') {
            Write-Host "Patch ThisDocument code from $($f.Name)"
            $code = Get-Content -Raw -Path $f.FullName
            # Strip the class header so only Option/Subs are pasted.
            $code = ($code -split "Option Explicit", 2)[-1]
            $code = "Option Explicit`r`n" + $code
            $thisDoc = $vbproj.VBComponents.Item('ThisDocument')
            $cm = $thisDoc.CodeModule
            if ($cm.CountOfLines -gt 0) {
                $cm.DeleteLines(1, $cm.CountOfLines)
            }
            $cm.AddFromString($code)
        } else {
            Write-Host "Import class:  $($f.Name)"
            $vbproj.VBComponents.Import($f.FullName) | Out-Null
        }
    }

    $doc.Save()
    Write-Host "Built: $OutputPath"
}
finally {
    if ($doc) { try { $doc.Close([ref]$false) } catch {} }
    if ($word) {
        try { $word.Quit([ref]0) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

<#
.SYNOPSIS
  Package the USAF OI Formatter as standalone Windows .exe files.

.DESCRIPTION
  Uses PyInstaller to build two binaries in dist\:
    usaf-oi-formatter.exe      (CLI; console app)
    usaf-oi-formatter-gui.exe  (Tkinter GUI; windowed app)

  Neither binary requires Python to be installed on the target machine;
  neither requires Word. Word is only needed to *view* the formatted .docx.

.PARAMETER OutputDir
  Override the PyInstaller dist/ output directory.

.PARAMETER Clean
  Wipe build/ and dist/ before packaging.
#>
[CmdletBinding()]
param(
    [string]$OutputDir = "",
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath
$repo = Resolve-Path (Join-Path $here '..')

Push-Location $repo
try {
    if ($Clean) {
        foreach ($d in 'build', 'dist') {
            if (Test-Path $d) { Remove-Item -Recurse -Force $d }
        }
        Get-ChildItem -Filter '*.spec' | Remove-Item -Force
    }

    # Fresh virtual env so the .exe is minimal.
    $venv = Join-Path $repo '.venv-build'
    if (-not (Test-Path $venv)) {
        Write-Host "Creating build venv in $venv"
        python -m venv $venv
    }
    $pyExe = Join-Path $venv 'Scripts\python.exe'
    $pipExe = Join-Path $venv 'Scripts\pip.exe'

    & $pipExe install --upgrade pip | Out-Null
    & $pipExe install '.[dev]' | Out-Null

    $distArgs = @()
    if ($OutputDir) {
        $distArgs = @('--distpath', $OutputDir)
    }

    Write-Host "Building CLI..."
    & $pyExe -m PyInstaller --clean --onefile --name usaf-oi-formatter `
        @distArgs `
        --collect-all docx `
        --collect-all lxml `
        'src/usaf_oi_formatter/cli.py'

    Write-Host "Building GUI..."
    & $pyExe -m PyInstaller --clean --onefile --windowed --name usaf-oi-formatter-gui `
        @distArgs `
        --collect-all docx `
        --collect-all lxml `
        'src/usaf_oi_formatter/gui.py'

    Write-Host ""
    Write-Host "Done. Binaries are in $(Join-Path $repo 'dist')."
}
finally {
    Pop-Location
}

# Test samples

Drop non-compliant OI `.docx` files here (not committed) and run:

```powershell
pwsh ..\..\tools\Format-USAFOI.ps1 -Path . -OutputDir .\out `
     -OPR 'CCSD/CCC' -OIName 'CCSD OI 36-1' `
     -Date '23 April 2026' -Subject 'Personnel Actions' `
     -Unit '442d Maintenance Squadron' -Category 'Personnel'
```

The formatter writes `<file>_formatted.docx` and a sidecar
`<file>_changes.txt` describing what was rewritten. Diff the output
visually against DAFMAN 90-161 Figure A2.2 to confirm the title block,
numbering, and attachment conventions are correct.

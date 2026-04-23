# CCSD_DOC_OI_Formatter

Automate formatting of **USAF Operating Instructions (OIs)** into compliance
with **AFH 33-337** (*The Tongue and Quill*) and **DAFMAN 90-161**
(*Publishing Processes and Procedures*).

The tool is a macro-enabled Word template (`USAF_OI_Formatter.dotm`) with an
embedded VBA engine. It exposes the same engine three ways:

- **Inside Word (GUI)** — tabbed UserForm for single-file or batch mode.
- **Inside Word (current document)** — a toolbar button runs the formatter on
  whatever document is open.
- **Headless CLI** — a PowerShell driver opens Word via COM and runs the
  formatter over a file or a folder (for scheduled / scripted batches).

All rules are hard-coded from AFH 33-337 and DAFMAN 90-161 — see
[`docs/rules.md`](docs/rules.md) for the citation table.

## What it does (automatically)

Given an arbitrary `.docx`, the formatter:

1. Sets 1" margins, 8.5 × 11 portrait, Arabic page numbers bottom-center
   (title page suppressed).
2. Installs canonical styles (`OI Body`, `OI Heading 1..5`, `OI Title`,
   `OI TitleBlock`, `OI Attachment Title`, `OI Bullet 1..4`).
3. Rebuilds the DAFMAN 90-161 Figure A2.2 title block from user-supplied
   metadata (Unit, OI number, date, OPR, Supersedes, Certified by, Pages,
   Accessibility, Releasability).
4. Walks every paragraph, heuristically classifies it (numbered heading,
   ALL-CAPS heading, body, bullet, attachment title) and reassigns the
   matching canonical style.
5. Rebuilds the multi-level numbering template (`1.`, `1.1.`, `1.1.1.`,
   `1.1.1.1.`, `1.1.1.1.1.`) and binds it to the five heading styles.
6. Normalizes bullets to the T&Q sequence (`-`, `•`, `–`, `»`) based on
   indent depth.
7. Scans for ALL-CAPS acronyms, collects them, and seeds a Glossary
   (Attachment 1) if the document doesn't already have one.
8. Rebuilds `Attachment N—TITLE` headings (em dash, ALL CAPS).
9. Cleans up whitespace (collapses double-spaces, converts smart quotes to
   straight, etc.).
10. Writes a side-car `<name>_changes.txt` describing every change.

Output is saved as `<name>_formatted.docx` next to the source (or in
`-OutputDir` if specified).

## What it does NOT do

Items listed in [`docs/rules.md` § Known gaps](docs/rules.md#known-gaps)
require human review — most notably tone/voice, semantic acronym accuracy,
reference citations, classification markings, and signature correctness.

## Repo layout

```
src/vba/                  VBA source (text, git-friendly)
  modRules.bas              all formatting constants
  modFormatter.bas          FormatDocument() pipeline
  modPageSetup.bas          margins / paper / page numbers
  modStyles.bas             OI style install/refresh
  modNumbering.bas          multi-level list template
  modBulletsLists.bas       bullet normalization
  modHeaderBlock.bas        DAFMAN 90-161 title block builder
  modAcronyms.bas           first-use expansion + glossary seed
  modAttachments.bas        Attachment N—TITLE rebuild
  modReport.bas             change log
  modBatch.bas              folder iteration
  modCLI.bas                Application.Run entry points
  frmMain.frm               tabbed UserForm (GUI)
  frmReport.frm             change report viewer
  ThisDocument.cls          AutoExec toolbar wiring
tools/
  Format-USAFOI.ps1         PowerShell headless CLI
  Install-Addin.ps1         copies .dotm into Word STARTUP
build/
  build-dotm.ps1            assembles .dotm from src\vba sources
docs/rules.md               rule citation table (AFH 33-337 / DAFMAN 90-161)
tests/samples/              drop test .docx files here
```

## Prerequisites

- Microsoft Word (2016 or newer).
- Windows PowerShell 5.1 or PowerShell 7+ (for the CLI and build scripts).
- One-time Word setting: **File → Options → Trust Center → Trust Center
  Settings → Macro Settings → "Trust access to the VBA project object model"**
  must be checked, otherwise `build\build-dotm.ps1` cannot inject the VBA
  sources.

## Install

```powershell
# From the repo root:
powershell -ExecutionPolicy Bypass -File build\build-dotm.ps1
powershell -ExecutionPolicy Bypass -File tools\Install-Addin.ps1
```

After install, restart Word. A **USAF OI Formatter** toolbar appears with
two buttons:

- **Format OI...** — runs the formatter on the current document with
  default metadata. Fast path for routine fixes.
- **Open Formatter GUI** — full tabbed dialog (Single / Batch / Metadata /
  Options) for richer control.

## CLI usage

Single file:

```powershell
tools\Format-USAFOI.ps1 -Path C:\incoming\MyOI.docx `
    -OPR 'CCSD/CCC' -OIName 'CCSD OI 36-1' `
    -Date '23 April 2026' -Subject 'Personnel Actions' `
    -Unit '442d Maintenance Squadron' -Category 'Personnel' `
    -CertifiedBy 'Col Jane Doe, Commander' -Pages '12'
```

Recursive batch with custom output directory:

```powershell
tools\Format-USAFOI.ps1 -Path C:\incoming -Recurse -OutputDir C:\out
```

Exit code is `0` on success, `1` if any file failed. See the emitted
`*_changes.txt` and the master `batch_<timestamp>.log` for details.

## Changing rules

1. Edit `src/vba/modRules.bas` (single source of truth).
2. Update the relevant row in `docs/rules.md` with the new AFH/DAFMAN
   reference.
3. Rebuild: `powershell -File build\build-dotm.ps1`.
4. Re-install: `powershell -File tools\Install-Addin.ps1 -Force`.
5. Run the tool against `tests/samples/` and verify the diff is confined to
   the changed rule.

## License

See [LICENSE](LICENSE).

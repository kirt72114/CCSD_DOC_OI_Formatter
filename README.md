# CCSD_DOC_OI_Formatter

Automate formatting of **USAF Operating Instructions (OIs)** into compliance
with **AFH 33-337** (*The Tongue and Quill*) and **DAFMAN 90-161** (*Publishing
Processes and Procedures*).

**No Word macros are needed.** The tool is a pure Python package that edits the
`.docx` XML directly (via `python-docx` + `lxml`). It runs:

- As a **CLI** — `usaf-oi-formatter path\to\file.docx [flags]`.
- As a **Tkinter GUI** — `usaf-oi-formatter-gui` — tabbed dialog with file
  picker, batch mode, and sticky OI-metadata fields.
- As a **standalone Windows `.exe`** bundled by PyInstaller for locked-down
  environments where Python isn't available.

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
   matching canonical style — skipping the title block.
5. Rebuilds the multi-level numbering template (`1.`, `1.1.`, `1.1.1.`,
   `1.1.1.1.`, `1.1.1.1.1.`) and binds it to the five heading styles.
6. Normalizes bullets to the T&Q sequence (`-`, `•`, `–`, `»`) based on
   indent depth.
7. Scans body text for ALL-CAPS acronyms, collects them, and seeds a
   Glossary (Attachment 1) if the document doesn't already have one.
8. Rebuilds `Attachment N—TITLE` headings (em dash, ALL CAPS).
9. Cleans up whitespace (collapses double-spaces, converts smart quotes to
   straight).
10. Writes a side-car `<name>_changes.txt` describing every change.

Output is saved as `<name>_formatted.docx` next to the source (or in
`--output-dir`).

## What it does NOT do

Items listed in [`docs/rules.md` § Known gaps](docs/rules.md#known-gaps)
require human review — most notably tone/voice, semantic acronym accuracy,
reference citations, classification markings, and signature correctness.

## Repo layout

```
src/usaf_oi_formatter/
  rules.py              all formatting constants (tune rules here)
  meta.py               OIMeta dataclass for title-block inputs
  formatter.py          orchestrator: format_file(path, meta) -> (out, report)
  pagesetup.py          margins / paper / page numbers
  styles.py             installs the OI * paragraph styles
  numbering.py          multi-level list rebuild via numbering.xml
  bullets.py            bullet normalization based on indent depth
  headerblock.py        DAFMAN 90-161 Fig A2.2 title block builder
  acronyms.py           collect acronyms for the Glossary attachment
  attachments.py        Attachment N—TITLE rebuild and Glossary seeding
  hygiene.py            whitespace, smart-quote, dash cleanup
  report.py             before/after change log (sidecar .txt)
  batch.py              folder iteration
  cli.py                argparse entry point
  gui.py                Tkinter GUI
  __main__.py           `python -m usaf_oi_formatter`
tools/
  build-exe.ps1         PyInstaller wrapper -> dist\*.exe
tests/
  test_rules.py
  test_formatter.py     end-to-end smoke test
  samples/              drop your own .docx fixtures here
docs/rules.md           rule citation table
pyproject.toml
```

## Prerequisites

- **Python 3.10+** (any modern cpython; uses only `python-docx`, `lxml`, and
  stdlib `tkinter`).
- **Word 2016+** only if you want to *view* the formatted `.docx`. The
  formatter itself never launches Word.

To build the standalone `.exe`:
- PowerShell 5.1+ on Windows.
- The `[dev]` extra (PyInstaller); `tools\build-exe.ps1` installs it into a
  throwaway `.venv-build`.

## Install (developer / editable)

```bash
python -m venv .venv
.venv\Scripts\activate       # Windows
# source .venv/bin/activate    # macOS/Linux
pip install -e .[dev]
```

## CLI usage

Single file:

```
usaf-oi-formatter C:\incoming\MyOI.docx ^
    --opr "CCSD/CCC" --oi-number "CCSD OI 36-1" ^
    --date "23 April 2026" --subject "Personnel Actions" ^
    --unit "442d Maintenance Squadron" --category "Personnel" ^
    --certified-by "Col Jane Doe, Commander" --pages 12
```

Recursive batch with custom output directory:

```
usaf-oi-formatter C:\incoming --recurse --output-dir C:\out
```

Exit code is `0` on success, `1` if any file failed. See the emitted
`*_changes.txt` and the master `batch_<timestamp>.log` for details.

Run `usaf-oi-formatter --help` for the full flag list.

## GUI usage

```
usaf-oi-formatter-gui
```

Tabs: **Input** (single file or batch folder + output folder) and
**OI Metadata** (Unit, OI number, date, OPR, Supersedes, ...). Metadata
values persist to `~/.usaf_oi_formatter.json` so recurring fields don't need
retyping.

## Build a standalone Windows `.exe`

For use on machines without Python installed:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build-exe.ps1 -Clean
```

Produces `dist\usaf-oi-formatter.exe` (CLI) and
`dist\usaf-oi-formatter-gui.exe` (windowed Tk app). Both are self-contained
and can be copied to any Windows machine.

## Tests

```bash
pytest
```

Includes an end-to-end test that builds a non-compliant `.docx`, runs the
full pipeline, and asserts the output has the expected styles, title block,
and glossary attachment.

## Changing rules

1. Edit `src/usaf_oi_formatter/rules.py`.
2. Update the corresponding row in `docs/rules.md` with the AFH/DAFMAN
   citation.
3. `pytest` to confirm nothing regressed.
4. Run the tool against a sample `.docx` and open it in Word to eyeball.
5. Rebuild the `.exe` if you distribute it: `tools\build-exe.ps1 -Clean`.

## License

See [LICENSE](LICENSE).

# CCSD_DOC_OI_Formatter

Automate formatting of **USAF Operating Instructions (OIs)** into compliance
with **AFH 33-337** (*The Tongue and Quill*) and **DAFMAN 90-161** (*Publishing
Processes and Procedures*).

**No Word macros are needed.** The tool is a pure Python package that edits the
`.docx` XML directly (via `python-docx` + `lxml`). It runs:

- As a **CLI** — `usaf-oi-formatter path\to\file.docx [flags]`.
- As a **modern Tkinter GUI** — `usaf-oi-formatter-gui` — with drag-and-drop,
  template picker, simple/advanced settings, sticky OI metadata, and a
  Tongue-and-Quill prose lint panel.
- As a **standalone Windows `.exe`** bundled by PyInstaller for locked-down
  environments where Python isn't available.

All structural rules are pinned to AFH 33-337 and DAFMAN 90-161 — see
[`docs/rules.md`](docs/rules.md) for the citation table. Tunable values
(margins, fonts, bullet glyphs, spacing) are bundled into a
**FormattingProfile** so units can deviate from the defaults without forking
the code.

## What it does (automatically)

Given an arbitrary `.docx` and a profile, the formatter:

1. Sets the page setup (margins, paper, page numbers) per the profile.
2. Installs canonical styles (`OI Body`, `OI Heading 1..5`, `OI Title`,
   `OI TitleBlock`, `OI Attachment Title`, `OI Bullet 1..4`).
3. Rebuilds the DAFMAN 90-161 Figure A2.2 title block from user-supplied
   metadata.
4. Walks every paragraph, heuristically classifies it, and reassigns the
   matching canonical style.
5. Rebuilds the multi-level numbering template (`1.`, `1.1.`, …) and binds it
   to the heading styles.
6. Normalizes bullets to the profile's bullet sequence based on indent depth.
7. Scans body text for ALL-CAPS acronyms and seeds a Glossary (Attachment 1)
   if the document doesn't already have one.
8. Rebuilds `Attachment N—TITLE` headings (em dash, ALL CAPS).
9. Cleans up whitespace (collapses double-spaces, converts smart quotes).
10. Writes a side-car `<name>_changes.txt` describing every change.

Output is saved as `<name>_formatted.docx` next to the source (or in
`--output-dir`).

## What it suggests (lint pass)

Enable `--lint` (CLI) or "Run T&Q lint" (GUI) to also write
`<name>_formatted_lint.txt` with prose-level findings the structural pass
cannot fix:

- Passive voice (T&Q Ch. 6).
- Sentences over 30 words / paragraphs over 120 words.
- Gendered pronouns ("he"/"his"/"she"/"her") and `Airman` (singular).
- Wordy phrases (`utilize → use`, `in order to → to`, …) per T&Q Ch. 4.
- Acronyms used without an inline expansion (T&Q Ch. 8).

The lint pass never modifies the document — every finding is a suggestion.

## What it does NOT do

Items listed in [`docs/rules.md` § Known gaps](docs/rules.md#known-gaps)
require human review — most notably semantic acronym accuracy, reference
citations, classification markings, and signature correctness.

## Repo layout

```
src/usaf_oi_formatter/
  rules.py              DAFMAN-mandated structural constants (style names, labels)
  profile.py            FormattingProfile dataclass — every user-tunable knob
  templates.py          Built-in templates: T&Q, Approved OI, Compact
  meta.py               OIMeta dataclass for title-block inputs
  formatter.py          orchestrator: format_file(path, meta, profile, run_lint)
  pagesetup.py          margins / paper / page numbers
  styles.py             installs the OI * paragraph styles
  numbering.py          multi-level list rebuild via numbering.xml
  bullets.py            bullet normalization based on indent depth
  headerblock.py        DAFMAN 90-161 Fig A2.2 title block builder
  acronyms.py           collect acronyms for the Glossary attachment
  attachments.py        Attachment N—TITLE rebuild and Glossary seeding
  hygiene.py            whitespace, smart-quote, dash cleanup
  lint.py               prose-level Tongue & Quill suggestions (read-only)
  report.py             before/after change log (sidecar .txt)
  batch.py              folder iteration
  cli.py                argparse entry point
  gui.py                Tkinter GUI (drag-drop, templates, simple/advanced settings)
  __main__.py           `python -m usaf_oi_formatter`
tools/
  build-exe.ps1         PyInstaller wrapper -> dist\*.exe
tests/
  test_rules.py
  test_profile.py
  test_lint.py
  test_formatter.py     end-to-end smoke test
  samples/              drop your own .docx fixtures here
docs/rules.md           rule citation table
pyproject.toml
```

## Prerequisites

- **Python 3.10+** (uses only `python-docx`, `lxml`, and stdlib `tkinter`).
- **Word 2016+** only if you want to *view* the formatted `.docx`.
- *Optional:* `tkinterdnd2` for OS-level drag-and-drop in the GUI. Without it
  the GUI still works — the drop zone becomes a click-to-browse target.

To build the standalone `.exe`:
- PowerShell 5.1+ on Windows.
- The `[dev]` extra (PyInstaller); `tools\build-exe.ps1` installs it into a
  throwaway `.venv-build`.

## Install (developer / editable)

```bash
python -m venv .venv
.venv\Scripts\activate       # Windows
# source .venv/bin/activate    # macOS/Linux
pip install -e .[dev,gui]
```

## CLI usage

Single file with the default Tongue-and-Quill template plus lint:

```
usaf-oi-formatter C:\incoming\MyOI.docx --lint ^
    --opr "CCSD/CCC" --oi-number "CCSD OI 36-1" ^
    --date "23 April 2026" --subject "Personnel Actions" ^
    --unit "442d Maintenance Squadron" --category "Personnel" ^
    --certified-by "Col Jane Doe, Commander" --pages 12
```

Pick a different built-in template or load a custom profile:

```
usaf-oi-formatter C:\incoming\MyOI.docx --template "Compact"
usaf-oi-formatter C:\incoming\MyOI.docx --profile-file C:\profiles\local.json
```

Recursive batch with custom output directory:

```
usaf-oi-formatter C:\incoming --recurse --output-dir C:\out
```

Exit code is `0` on success, `1` if any file failed. See the emitted
`*_changes.txt`, `*_lint.txt`, and the master `batch_<timestamp>.log` for
details.

Run `usaf-oi-formatter --help` for the full flag list.

## GUI usage

```
usaf-oi-formatter-gui
```

Tabs:

- **Files** — drop `.docx` files into the drop zone (or click to browse / add
  folders). Pick an output folder and toggle the lint pass.
- **Settings** — choose a template (Tongue and Quill / Approved OI / Compact /
  custom JSON), then override any field. Common settings are shown by
  default; flip "Show advanced settings" for the full FormattingProfile.
  Save the current configuration as a custom profile JSON for reuse.
- **OI Metadata** — DAFMAN 90-161 Fig A2.2 fields. Persisted to
  `~/.usaf_oi_formatter.json` so recurring values don't need retyping.
- **Lint Results** — populated after each run when "Run T&Q lint" is on.

The Settings tab live-validates and warns if values stray from
DAFMAN/T&Q expectations (unusual margins, non-letter paper, non-standard
fonts, etc.).

## Templates and profiles

Built-in templates live in [`src/usaf_oi_formatter/templates.py`](src/usaf_oi_formatter/templates.py):

- **Tongue and Quill** — strict AFH 33-337 / DAFMAN 90-161 defaults.
- **Approved OI** — same defaults, with title-block rebuild and glossary
  seeding both forced on.
- **Compact** — tighter spacing for short OIs.

A profile is a `FormattingProfile` instance; saving from the GUI (or calling
`profile.save(path)` in code) writes it to a JSON file you can ship with
`--profile-file`.

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

Includes:
- `test_rules.py` — pure-helper sanity (level → style mapping).
- `test_profile.py` — profile validation, copy/save/load round-trip,
  template registry.
- `test_lint.py` — prose-level findings for passive voice, long sentences,
  wordy phrases, undefined acronyms (with the parenthetical-definition
  heuristic).
- `test_formatter.py` — end-to-end pipeline plus custom-profile margins,
  glossary toggle, and lint sidecar.

## Changing rules

1. If the change is structural (DAFMAN-mandated): edit
   `src/usaf_oi_formatter/rules.py`.
2. If it's a default value the user could override: edit the field's default
   in `src/usaf_oi_formatter/profile.py`.
3. Update the corresponding row in `docs/rules.md` with the AFH/DAFMAN
   citation.
4. `pytest` to confirm nothing regressed.
5. Run the tool against a sample `.docx` and open it in Word to eyeball.
6. Rebuild the `.exe` if you distribute it: `tools\build-exe.ps1 -Clean`.

## License

See [LICENSE](LICENSE).

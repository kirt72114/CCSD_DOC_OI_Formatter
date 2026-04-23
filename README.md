# CCSD_DOC_OI_Formatter

Automate formatting of USAF Operating Instructions (OIs) by cloning the
formatting of an **approved** OI onto a **draft**. No macros required —
it's a pure Python tool that edits the `.docx` XML directly.

## How it works

You point the tool at:

- an **approved OI** (`--template approved.docx`) — a known-good file whose
  fonts, styles, margins, running headers, footers, and numbering scheme
  represent your unit's house style;
- one or more **drafts** — the files whose formatting you want to normalize.

The tool produces a `<draft>_formatted.docx` that has the **draft's content**
in the **template's clothing**. Specifically, these parts of the `.docx`
are copied from template → draft:

| What | From | Why |
|---|---|---|
| `word/styles.xml` | template | All paragraph + character style definitions |
| `word/numbering.xml` | template | Multi-level list schemes |
| `word/theme/theme1.xml` | template | Color/font theme |
| `word/header*.xml` + `word/footer*.xml` | template | Running headers, page numbers |
| Referenced images (e.g., USAF seal) | template | So the header block renders |
| Page size, margins, first-page flag | template | Section properties |
| Draft body content | *draft* | Paragraphs, tables, images stay put |

The draft's embedded images, tables, and body paragraphs are preserved.

## Repo layout

```
src/usaf_oi_formatter/
  template.py           zip-surgery: styles/numbering/headers/footers + page setup
  formatter.py          orchestrator: format_file(src, output_dir, template)
  hygiene.py            whitespace + smart-quote cleanup (runs in both modes)
  report.py             sidecar change log
  batch.py              folder iteration
  cli.py                argparse entry point
  gui.py                Tkinter GUI
  rules.py              output naming constants
tools/build-exe.ps1     PyInstaller wrapper -> dist\*.exe
tests/
  test_formatter.py     covers template and hygiene-only paths
wheels/                 prebuilt Windows wheels for offline install
docs/rules.md           notes on which rules the tool does and doesn't enforce
pyproject.toml
```

## Prerequisites

- **Python 3.10+** (the tool itself + dependencies: `python-docx`, `lxml`).
- **Word 2016+** only to *view* the formatted `.docx` — the formatter never
  launches Word.

## Install

### Online (developer workstation)

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

### Offline (locked-down USAF workstation)

PyPI is often unreachable through corporate TLS inspection. The repo ships
with prebuilt Windows wheels in `wheels/`:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --no-index --find-links .\wheels `
    setuptools wheel
.\.venv\Scripts\python.exe -m pip install --no-index --find-links .\wheels `
    python-docx lxml pywin32
.\.venv\Scripts\python.exe -m pip install --no-deps --no-build-isolation -e .
```

`pywin32` is only needed if you pass legacy `.doc` files as the draft or the
template; it drives Word via COM to upconvert them to `.docx`. Skip it if you
only work with `.docx`.

## CLI usage

Single draft, using an approved OI as the template:

```powershell
.\.venv\Scripts\usaf-oi-formatter.exe "Draft OI 17-1203.docx" `
    --template "Approved OI 17-1203.docx"
```

Legacy `.doc` templates (or drafts) are accepted too — the tool runs a
headless Word instance via `pywin32` to convert them to `.docx` in a temp
folder before the template-clone step. The output is always `.docx`:

```powershell
.\.venv\Scripts\usaf-oi-formatter.exe "Draft OI 17-1203.doc" `
    --template "Approved OI 17-1203 28 Jul 17.doc"
# -> Draft OI 17-1203_formatted.docx
```

Whole folder of drafts against one template:

```powershell
.\.venv\Scripts\usaf-oi-formatter.exe "C:\incoming" --recurse `
    --template "Approved OI 17-1203.docx" --output-dir "C:\out"
```

No template — just do hygiene (whitespace/quotes), don't touch structure:

```powershell
.\.venv\Scripts\usaf-oi-formatter.exe "Draft.docx"
```

Each run writes `<name>_formatted.docx` plus `<name>_formatted_changes.txt`.

## GUI usage

```powershell
.\.venv\Scripts\usaf-oi-formatter-gui.exe
```

Tkinter dialog with fields for the draft (or folder), the approved
template, and an optional output folder. Values persist to
`~/.usaf_oi_formatter.json` between runs.

## Build a standalone Windows `.exe`

For machines without Python:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build-exe.ps1 -Clean
```

Produces `dist\usaf-oi-formatter.exe` and `dist\usaf-oi-formatter-gui.exe`
— self-contained, no Python install required on the target machine.

## Tests

```powershell
.\.venv\Scripts\pytest.exe
```

Covers both the template-copy and hygiene-only paths.

## What it does not do

The tool is intentionally **structural** — it does not try to enforce
Tongue and Quill *prose* guidelines (voice, acronym first-use, sentence
length, etc.). Those still need a human reviewer. If you want automated
linting for those, file an issue describing which rules matter most and
we can add a `--lint` mode that reports (but doesn't auto-fix) violations.

## License

See [LICENSE](LICENSE).

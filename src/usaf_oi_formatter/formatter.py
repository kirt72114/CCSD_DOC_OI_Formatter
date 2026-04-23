"""Orchestrator.

Two modes:

- Template mode (recommended): given an approved OI as a reference
  template, copy its styles, numbering, headers, footers, theme, and
  page setup onto the draft. See `template.py` for the mechanics.

- Hygiene-only mode (no template): just collapse double spaces, convert
  smart quotes to straight, etc. No structural rewrites.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import docx

from . import hygiene
from . import report as report_mod
from . import rules
from . import template as template_mod


def format_file(src: Path,
                output_dir: Path | None = None,
                template: Path | None = None) -> tuple[Path, Path]:
    """Format `src`. Writes the result next to it (or into `output_dir`)
    and returns `(formatted_path, report_path)`.

    If `template` is given, it is treated as an approved OI and its
    formatting parts are cloned onto the draft. Otherwise the draft gets
    only text hygiene.
    """
    src = Path(src)
    out_path = _output_path(src, output_dir)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    report = report_mod.ChangeReport(src)

    if template is not None:
        report.note("template", f"Cloning formatting from {Path(template).name}")
        template_mod.apply_template(src, Path(template), out_path)
    else:
        report.note("copy", "No template supplied; performing hygiene only")
        shutil.copyfile(src, out_path)

    doc = docx.Document(str(out_path))
    report.snapshot_pre(_snapshot(doc))

    report.note("hygiene", "Whitespace / quotes / dashes")
    hygiene.apply(doc)

    doc.save(str(out_path))

    post_doc = docx.Document(str(out_path))
    report.diff_post(_snapshot(post_doc))

    report_path = report.write_sidecar(out_path)
    return out_path, report_path


def _output_path(src: Path, output_dir: Path | None) -> Path:
    new_name = src.stem + rules.OUTPUT_SUFFIX + src.suffix
    if output_dir is None:
        return src.with_name(new_name)
    return Path(output_dir) / new_name


def _snapshot(doc) -> dict[str, str]:
    section = doc.sections[0]
    return {
        "margin.top_emu": str(section.top_margin),
        "margin.bottom_emu": str(section.bottom_margin),
        "margin.left_emu": str(section.left_margin),
        "margin.right_emu": str(section.right_margin),
        "page.width_emu": str(section.page_width),
        "page.height_emu": str(section.page_height),
        "paragraph.count": str(len(doc.paragraphs)),
    }

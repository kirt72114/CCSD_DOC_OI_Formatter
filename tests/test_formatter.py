"""End-to-end smoke test: build a throwaway .docx, format it, reopen, check."""

from __future__ import annotations

from pathlib import Path

import pytest

docx = pytest.importorskip("docx")

from usaf_oi_formatter import formatter, rules, templates
from usaf_oi_formatter.meta import OIMeta
from usaf_oi_formatter.profile import FormattingProfile


def _make_sample(path: Path) -> None:
    doc = docx.Document()
    doc.add_paragraph("SOME OLD TITLE")
    doc.add_paragraph("1. Purpose")
    doc.add_paragraph("This Operating Instruction defines local procedures.")
    doc.add_paragraph("1.1. Scope")
    doc.add_paragraph("Applies to all CCSD personnel.")
    doc.add_paragraph("- First bullet item")
    doc.add_paragraph("- Second bullet item")
    doc.add_paragraph("2. References")
    doc.add_paragraph("See AFI 36-2618 and DAFMAN 90-161.")
    doc.save(str(path))


def test_end_to_end(tmp_path: Path) -> None:
    src = tmp_path / "raw.docx"
    _make_sample(src)

    meta = OIMeta(
        unit="442D MAINTENANCE SQUADRON",
        oi_number="CCSD OI 36-1",
        date_str="23 April 2026",
        category="Personnel",
        subject="Personnel Actions",
        opr="CCSD/CCC",
        certified_by="Col Jane Doe, Commander",
        pages="12",
    )

    out_path, report_path = formatter.format_file(src, meta)

    assert out_path.exists()
    assert report_path.exists()
    assert out_path.name == f"raw{rules.OUTPUT_SUFFIX}.docx"

    result = docx.Document(str(out_path))
    style_names = {s.name for s in result.styles}
    for name in (rules.STY_BODY, rules.STY_H1, rules.STY_TITLEBLOCK,
                 rules.STY_ATTACH_TITLE):
        assert name in style_names, name

    # Title block text should appear somewhere in the first few paragraphs.
    text_blob = "\n".join(p.text for p in result.paragraphs[:20]
                          for _ in [0])
    assert rules.LBL_COMPLIANCE in text_blob or any(
        rules.LBL_COMPLIANCE in cell.text
        for t in result.tables for row in t.rows for cell in row.cells
    )

    # Attachment 1 Glossary was seeded.
    all_text = "\n".join(p.text for p in result.paragraphs)
    assert f"{rules.ATTACH_PREFIX}1{rules.ATTACH_SEP}" in all_text

    # Change report is non-empty.
    assert report_path.read_text(encoding="utf-8").strip()


def test_custom_profile_changes_margins(tmp_path: Path) -> None:
    src = tmp_path / "raw.docx"
    _make_sample(src)

    profile = FormattingProfile.tongue_and_quill().copy(
        name="Half-inch margins",
        margin_top_in=0.5, margin_bottom_in=0.5,
        margin_left_in=0.5, margin_right_in=0.5,
    )
    out_path, _ = formatter.format_file(src, OIMeta(), profile=profile)

    result = docx.Document(str(out_path))
    section = result.sections[0]
    # 0.5 in == 457200 EMU
    assert section.top_margin.emu == 457200
    assert section.left_margin.emu == 457200


def test_lint_sidecar_written_when_requested(tmp_path: Path) -> None:
    src = tmp_path / "raw.docx"
    _make_sample(src)

    result = formatter.format_file(
        src, OIMeta(), profile=templates.default(), run_lint=True)

    assert len(result) == 3
    out_path, report_path, lint_path = result
    assert lint_path.exists()
    assert "lint" in lint_path.name


def test_disable_glossary_skips_attachment_1(tmp_path: Path) -> None:
    src = tmp_path / "raw.docx"
    _make_sample(src)

    profile = FormattingProfile.tongue_and_quill().copy(seed_glossary=False)
    out_path, _ = formatter.format_file(src, OIMeta(), profile=profile)

    result = docx.Document(str(out_path))
    text = "\n".join(p.text for p in result.paragraphs)
    assert rules.GLOSSARY_TITLE not in text

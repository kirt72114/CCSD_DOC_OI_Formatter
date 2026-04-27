"""Normalize bullet characters to the T&Q sequence based on indent depth."""

from __future__ import annotations

from docx.document import Document
from docx.shared import Inches

from . import rules
from .profile import FormattingProfile, default as _default_profile


_NON_BULLET_STYLES = set(rules.HEADING_STYLES) | {
    rules.STY_TITLE, rules.STY_TITLEBLOCK, rules.STY_ATTACH_TITLE
}

# Glyphs we treat as "this paragraph starts with a bullet". Static set
# (covers any reasonable input doc, regardless of profile choices).
BULLET_CHARS = set("-*o•–—»▪")


def apply(doc: Document, profile: FormattingProfile | None = None) -> None:
    p = profile or _default_profile()
    indent_step = max(0.05, p.bullet_indent_step_in)

    for para in doc.paragraphs:
        if para.style.name in _NON_BULLET_STYLES:
            continue

        text = para.text
        if len(text) < 2 or text[0] not in BULLET_CHARS:
            continue

        level = _level_from_indent(para, indent_step)
        _replace_leading_bullet(para, p.bullet_for_level(level))
        para.style = doc.styles[p.bullet_style_for_level(level)]


def _level_from_indent(paragraph, indent_step_in: float) -> int:
    pf = paragraph.paragraph_format
    left = pf.left_indent
    if left is None:
        return 1
    step = Inches(indent_step_in)
    level = int(left / step) + 1
    return max(1, min(level, 4))


def _replace_leading_bullet(paragraph, new_glyph: str) -> None:
    if not paragraph.runs:
        return
    first = paragraph.runs[0]
    rest = first.text[1:].lstrip()
    first.text = f"{new_glyph} {rest}"

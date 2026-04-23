"""Normalize bullet characters to the T&Q sequence based on indent depth."""

from __future__ import annotations

from docx.document import Document
from docx.shared import Inches

from . import rules

# Glyphs we treat as "this paragraph starts with a bullet".
BULLET_CHARS = set("-*o" + "".join([rules.BULLET_L2, rules.BULLET_L3, rules.BULLET_L4,
                                    "–", "—", "▪"]))

_NON_BULLET_STYLES = set(rules.HEADING_STYLES) | {
    rules.STY_TITLE, rules.STY_TITLEBLOCK, rules.STY_ATTACH_TITLE
}


def apply(doc: Document) -> None:
    for p in doc.paragraphs:
        if p.style.name in _NON_BULLET_STYLES:
            continue

        text = p.text
        if len(text) < 2 or text[0] not in BULLET_CHARS:
            continue

        level = _level_from_indent(p)
        _replace_leading_bullet(p, rules.bullet_for_level(level))
        p.style = doc.styles[rules.bullet_style_for_level(level)]


def _level_from_indent(paragraph) -> int:
    pf = paragraph.paragraph_format
    left = pf.left_indent
    if left is None:
        return 1
    # 0.25" per level; clamp 1..4
    quarter_inch = Inches(0.25)
    level = int(left / quarter_inch) + 1
    return max(1, min(level, 4))


def _replace_leading_bullet(paragraph, new_glyph: str) -> None:
    if not paragraph.runs:
        return
    first = paragraph.runs[0]
    rest = first.text[1:].lstrip()
    first.text = f"{new_glyph} {rest}"

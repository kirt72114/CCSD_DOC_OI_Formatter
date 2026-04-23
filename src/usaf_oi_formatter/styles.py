"""Install or refresh the OI named paragraph styles."""

from __future__ import annotations

from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Inches, Pt

from . import rules


def install_or_refresh(doc: Document) -> None:
    _ensure_body(doc)
    for name, level in zip(rules.HEADING_STYLES, range(1, 6)):
        _ensure_heading(doc, name, level)
    _ensure_title(doc)
    _ensure_titleblock(doc)
    _ensure_attachment_title(doc)
    for name, level in zip(rules.BULLET_STYLES, range(1, 5)):
        _ensure_bullet(doc, name, level)


def _ensure(doc: Document, name: str) -> object:
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def _ensure_body(doc: Document) -> None:
    s = _ensure(doc, rules.STY_BODY)
    s.next_paragraph_style = s
    s.font.name = rules.BODY_FONT
    s.font.size = Pt(rules.BODY_SIZE_PT)
    s.font.bold = False
    s.font.italic = False
    pf = s.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(rules.SPACE_AFTER_PT)
    pf.first_line_indent = Inches(0)
    pf.left_indent = Inches(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.widow_control = True


def _ensure_heading(doc: Document, name: str, level: int) -> None:
    s = _ensure(doc, name)
    s.base_style = doc.styles[rules.STY_BODY]
    s.next_paragraph_style = doc.styles[rules.STY_BODY]
    s.font.name = rules.HEADING_FONT
    s.font.size = Pt(rules.HEADING_SIZE_PT)
    s.font.bold = True
    s.font.italic = False
    pf = s.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(12 if level == 1 else 6)
    pf.space_after = Pt(rules.SPACE_AFTER_PT)
    pf.keep_with_next = True
    pf.widow_control = True
    pf.left_indent = Inches(0)


def _ensure_title(doc: Document) -> None:
    s = _ensure(doc, rules.STY_TITLE)
    s.next_paragraph_style = doc.styles[rules.STY_BODY]
    s.font.name = rules.HEADING_FONT
    s.font.size = Pt(14)
    s.font.bold = True
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(12)
    pf.space_after = Pt(12)


def _ensure_titleblock(doc: Document) -> None:
    s = _ensure(doc, rules.STY_TITLEBLOCK)
    s.next_paragraph_style = s
    s.font.name = rules.TITLEBLOCK_FONT
    s.font.size = Pt(rules.TITLEBLOCK_SIZE_PT)
    s.font.bold = False
    pf = s.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _ensure_attachment_title(doc: Document) -> None:
    s = _ensure(doc, rules.STY_ATTACH_TITLE)
    s.base_style = doc.styles[rules.STY_BODY]
    s.next_paragraph_style = doc.styles[rules.STY_BODY]
    s.font.bold = True
    s.font.all_caps = True
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.page_break_before = True
    pf.space_after = Pt(12)
    pf.keep_with_next = True


def _ensure_bullet(doc: Document, name: str, level: int) -> None:
    s = _ensure(doc, name)
    s.base_style = doc.styles[rules.STY_BODY]
    s.next_paragraph_style = s
    pf = s.paragraph_format
    pf.left_indent = Inches(0.25 * level)
    pf.first_line_indent = Inches(-0.25)
    pf.space_after = Pt(3)

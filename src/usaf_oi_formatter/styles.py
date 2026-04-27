"""Install or refresh the OI named paragraph styles."""

from __future__ import annotations

from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Inches, Pt

from . import rules
from .profile import FormattingProfile, default as _default_profile


def install_or_refresh(doc: Document, profile: FormattingProfile | None = None) -> None:
    p = profile or _default_profile()
    _ensure_body(doc, p)
    for name, level in zip(rules.HEADING_STYLES, range(1, 6)):
        _ensure_heading(doc, p, name, level)
    _ensure_title(doc, p)
    _ensure_titleblock(doc, p)
    _ensure_attachment_title(doc, p)
    for name, level in zip(rules.BULLET_STYLES, range(1, 5)):
        _ensure_bullet(doc, p, name, level)


def _ensure(doc: Document, name: str) -> object:
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def _ensure_body(doc: Document, p: FormattingProfile) -> None:
    s = _ensure(doc, rules.STY_BODY)
    s.next_paragraph_style = s
    s.font.name = p.body_font
    s.font.size = Pt(p.body_size_pt)
    s.font.bold = False
    s.font.italic = False
    pf = s.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(p.space_after_pt)
    pf.first_line_indent = Inches(0)
    pf.left_indent = Inches(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.widow_control = True


def _ensure_heading(doc: Document, p: FormattingProfile, name: str, level: int) -> None:
    s = _ensure(doc, name)
    s.base_style = doc.styles[rules.STY_BODY]
    s.next_paragraph_style = doc.styles[rules.STY_BODY]
    s.font.name = p.heading_font
    s.font.size = Pt(p.heading_size_pt)
    s.font.bold = True
    s.font.italic = False
    pf = s.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(p.heading_space_before_pt if level == 1
                         else p.sub_heading_space_before_pt)
    pf.space_after = Pt(p.space_after_pt)
    pf.keep_with_next = True
    pf.widow_control = True
    pf.left_indent = Inches(0)


def _ensure_title(doc: Document, p: FormattingProfile) -> None:
    s = _ensure(doc, rules.STY_TITLE)
    s.next_paragraph_style = doc.styles[rules.STY_BODY]
    s.font.name = p.heading_font
    s.font.size = Pt(p.title_size_pt)
    s.font.bold = True
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(p.heading_space_before_pt)
    pf.space_after = Pt(p.heading_space_before_pt)


def _ensure_titleblock(doc: Document, p: FormattingProfile) -> None:
    s = _ensure(doc, rules.STY_TITLEBLOCK)
    s.next_paragraph_style = s
    s.font.name = p.titleblock_font
    s.font.size = Pt(p.titleblock_size_pt)
    s.font.bold = False
    pf = s.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _ensure_attachment_title(doc: Document, p: FormattingProfile) -> None:
    s = _ensure(doc, rules.STY_ATTACH_TITLE)
    s.base_style = doc.styles[rules.STY_BODY]
    s.next_paragraph_style = doc.styles[rules.STY_BODY]
    s.font.bold = True
    s.font.all_caps = True
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.page_break_before = True
    pf.space_after = Pt(p.heading_space_before_pt)
    pf.keep_with_next = True


def _ensure_bullet(doc: Document, p: FormattingProfile, name: str, level: int) -> None:
    s = _ensure(doc, name)
    s.base_style = doc.styles[rules.STY_BODY]
    s.next_paragraph_style = s
    pf = s.paragraph_format
    pf.left_indent = Inches(p.bullet_indent_step_in * level)
    pf.first_line_indent = Inches(-p.bullet_indent_step_in)
    pf.space_after = Pt(p.bullet_space_after_pt)

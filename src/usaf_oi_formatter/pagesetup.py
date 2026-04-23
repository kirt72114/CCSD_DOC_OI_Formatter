"""Page setup: margins, paper size, orientation, page numbers."""

from __future__ import annotations

from docx.document import Document
from docx.enum.section import WD_ORIENTATION
from docx.oxml.ns import qn
from docx.shared import Inches

from . import rules


def apply(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Inches(rules.MARGIN_IN)
        section.bottom_margin = Inches(rules.MARGIN_IN)
        section.left_margin = Inches(rules.MARGIN_IN)
        section.right_margin = Inches(rules.MARGIN_IN)
        section.page_width = Inches(rules.PAGE_WIDTH_IN)
        section.page_height = Inches(rules.PAGE_HEIGHT_IN)
        section.orientation = WD_ORIENTATION.PORTRAIT
        section.different_first_page_header_footer = True
        _install_page_number(section)


def _install_page_number(section) -> None:
    """Bottom-center Arabic numeral via PAGE field in the primary footer."""
    footer = section.footer
    if footer.paragraphs:
        p = footer.paragraphs[0]
    else:
        p = footer.add_paragraph()

    # Clear existing runs.
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.alignment = 1  # WD_PARAGRAPH_ALIGNMENT.CENTER

    run = p.add_run()
    _append_field(run, "PAGE \\* Arabic")


def _append_field(run, instruction: str) -> None:
    """Insert a Word field into an existing run."""
    from lxml import etree

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    nsmap = {"w": W}

    def el(tag: str) -> etree._Element:
        return etree.SubElement(run._element, qn(f"w:{tag}"))

    fld_begin = el("fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = etree.SubElement(run._element, qn("w:instrText"))
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "

    fld_sep = el("fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    # Placeholder number; Word overwrites it on field update.
    t = etree.SubElement(run._element, qn("w:t"))
    t.text = "1"

    fld_end = el("fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    _ = nsmap  # silence linters; namespace is registered via qn()

"""Multi-level list (1. / 1.1. / 1.1.1. / ... / 1.1.1.1.1.) for OI headings.

python-docx has no high-level API for numbering.xml, so we build the
abstractNum + num elements directly and attach a pPr/numPr to each heading
paragraph that already maps to an OI heading style.
"""

from __future__ import annotations

from docx.document import Document
from docx.oxml.ns import qn
from lxml import etree

from . import rules

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def apply(doc: Document) -> None:
    numbering_part = _ensure_numbering_part(doc)
    abs_id, num_id = _ensure_oi_numbering(numbering_part)

    for p in doc.paragraphs:
        if p.style.name in rules.HEADING_STYLES:
            level = rules.HEADING_STYLES.index(p.style.name)
            _attach_numbering(p, num_id, level)


# --- numbering.xml plumbing -----------------------------------------

def _ensure_numbering_part(doc: Document):
    """Return the document's numbering part, creating it if absent."""
    part = doc.part.numbering_part
    if part is None:
        # python-docx lazily creates this only when add_num is called.
        # Force creation via a throwaway abstractNum.
        part = _create_numbering_part(doc)
    return part


def _create_numbering_part(doc: Document):
    from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
    from docx.opc.part import PartFactory
    from docx.opc.packuri import PackURI
    from docx.parts.numbering import NumberingPart

    partname = PackURI("/word/numbering.xml")
    numbering_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:numbering xmlns:w="{W}"/>'
    )
    part = NumberingPart(
        partname,
        CONTENT_TYPE.WML_NUMBERING,
        numbering_xml.encode("utf-8"),
        doc.part.package,
    )
    doc.part.relate_to(part, RELATIONSHIP_TYPE.NUMBERING)
    return part


def _ensure_oi_numbering(numbering_part) -> tuple[int, int]:
    """Insert (idempotent) an abstractNum for the OI scheme and a num that
    references it. Returns (abstractNumId, numId)."""
    root = numbering_part.element

    # Reuse if we've already installed the OI list.
    for abs_num in root.findall(qn("w:abstractNum")):
        nsid = abs_num.find(qn("w:nsid"))
        if nsid is not None and nsid.get(qn("w:val")) == "0FC0FFEE":
            abs_id = int(abs_num.get(qn("w:abstractNumId")))
            for num in root.findall(qn("w:num")):
                ref = num.find(qn("w:abstractNumId"))
                if ref is not None and int(ref.get(qn("w:val"))) == abs_id:
                    return abs_id, int(num.get(qn("w:numId")))

    abs_id = _next_id(root, qn("w:abstractNum"), qn("w:abstractNumId"))
    num_id = _next_id(root, qn("w:num"), qn("w:numId"))

    abs_num = _build_abstract_num(abs_id)
    # abstractNum must come before num in numbering.xml.
    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abs_num)
    else:
        root.append(abs_num)

    num = etree.SubElement(root, qn("w:num"))
    num.set(qn("w:numId"), str(num_id))
    ref = etree.SubElement(num, qn("w:abstractNumId"))
    ref.set(qn("w:val"), str(abs_id))

    return abs_id, num_id


def _next_id(root, tag: str, id_attr: str) -> int:
    ids = [int(el.get(id_attr) or "0") for el in root.findall(tag)]
    return (max(ids) + 1) if ids else 1


def _build_abstract_num(abs_id: int) -> etree._Element:
    abs_num = etree.Element(qn("w:abstractNum"))
    abs_num.set(qn("w:abstractNumId"), str(abs_id))

    nsid = etree.SubElement(abs_num, qn("w:nsid"))
    nsid.set(qn("w:val"), "0FC0FFEE")

    multi = etree.SubElement(abs_num, qn("w:multiLevelType"))
    multi.set(qn("w:val"), "multilevel")

    for i in range(rules.MAX_NUMBER_DEPTH):
        lvl = etree.SubElement(abs_num, qn("w:lvl"))
        lvl.set(qn("w:ilvl"), str(i))

        start = etree.SubElement(lvl, qn("w:start"))
        start.set(qn("w:val"), "1")

        fmt = etree.SubElement(lvl, qn("w:numFmt"))
        fmt.set(qn("w:val"), "decimal")

        lvl_text = etree.SubElement(lvl, qn("w:lvlText"))
        # "%1." / "%1.%2." / ...
        lvl_text.set(qn("w:val"), "".join(f"%{n + 1}." for n in range(i + 1)))

        jc = etree.SubElement(lvl, qn("w:lvlJc"))
        jc.set(qn("w:val"), "left")

        ppr = etree.SubElement(lvl, qn("w:pPr"))
        ind = etree.SubElement(ppr, qn("w:ind"))
        # 0.25" per level, 20ths of a point (1 inch = 1440 twips)
        ind.set(qn("w:left"), str(int(360 * (i + 1))))
        ind.set(qn("w:hanging"), "360")

    return abs_num


def _attach_numbering(paragraph, num_id: int, level: int) -> None:
    p = paragraph._p
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = etree.SubElement(p, qn("w:pPr"))
        p.insert(0, pPr)

    # Replace any existing numPr.
    existing = pPr.find(qn("w:numPr"))
    if existing is not None:
        pPr.remove(existing)

    numPr = etree.SubElement(pPr, qn("w:numPr"))
    ilvl = etree.SubElement(numPr, qn("w:ilvl"))
    ilvl.set(qn("w:val"), str(level))
    nId = etree.SubElement(numPr, qn("w:numId"))
    nId.set(qn("w:val"), str(num_id))

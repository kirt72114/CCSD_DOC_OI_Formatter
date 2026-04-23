"""Template-based formatting.

Given an approved OI (`template.docx`) and a draft (`draft.docx`), this
module copies the formatting parts of the approved doc onto the draft
without touching the draft's content. What moves across:

  - word/styles.xml           (paragraph + character styles)
  - word/numbering.xml        (multi-level list definitions)
  - word/theme/theme1.xml     (color + font theme)
  - word/header*.xml          (running page headers)
  - word/footer*.xml          (footers incl. page number formatting)
  - word/media/*              (any image referenced by the headers/footers,
                                e.g. the USAF seal)
  - section properties        (margins, paper size, orientation,
                                different-first-page flag)

The draft's body (text, tables, images in the body) is preserved as-is.
The result is the draft's content in the template's clothing.

Implementation is deliberately zip-level rather than python-docx-level
because python-docx does not expose easy APIs for cross-document header
or numbering part copying.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

# ---- Namespaces -----------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NSMAP = {"w": W_NS, "r": R_NS}
CT_NSMAP = {"ct": CT_NS}
REL_NSMAP = {"rel": REL_NS}

# ---- Content types & relationship types -----------------------------

CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CT_IMAGE_PNG = "image/png"
CT_IMAGE_JPEG = "image/jpeg"
CT_IMAGE_GIF = "image/gif"
CT_IMAGE_EMF = "image/x-emf"
CT_IMAGE_WMF = "image/x-wmf"

REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


# =====================================================================
# Public API
# =====================================================================


def apply_template(draft_path: Path, template_path: Path,
                   output_path: Path) -> None:
    """Copy `template_path`'s formatting parts onto `draft_path` and save
    the result to `output_path`. Draft's body content is preserved."""
    draft_path = Path(draft_path)
    template_path = Path(template_path)
    output_path = Path(output_path)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_root = Path(tmp)
        draft_dir = tmp_root / "draft"
        template_dir = tmp_root / "template"
        _unzip(draft_path, draft_dir)
        _unzip(template_path, template_dir)

        _copy_styles(draft_dir, template_dir)
        _copy_theme(draft_dir, template_dir)
        _copy_numbering(draft_dir, template_dir)
        _copy_headers_footers(draft_dir, template_dir)
        _apply_page_setup(draft_dir, template_dir)

        _rezip(draft_dir, output_path)


# =====================================================================
# Steps
# =====================================================================


def _copy_styles(draft_dir: Path, template_dir: Path) -> None:
    src = template_dir / "word" / "styles.xml"
    if not src.exists():
        return
    dst = draft_dir / "word" / "styles.xml"
    shutil.copyfile(src, dst)

    _ensure_content_type_override(draft_dir, "/word/styles.xml", CT_STYLES)
    _ensure_document_relationship(
        draft_dir, rel_type=REL_STYLES, target="styles.xml")


def _copy_theme(draft_dir: Path, template_dir: Path) -> None:
    src = template_dir / "word" / "theme" / "theme1.xml"
    if not src.exists():
        return
    dst_dir = draft_dir / "word" / "theme"
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / "theme1.xml"
    shutil.copyfile(src, dst)

    _ensure_content_type_override(draft_dir, "/word/theme/theme1.xml", CT_THEME)
    _ensure_document_relationship(
        draft_dir, rel_type=REL_THEME, target="theme/theme1.xml")


def _copy_numbering(draft_dir: Path, template_dir: Path) -> None:
    src = template_dir / "word" / "numbering.xml"
    if not src.exists():
        return
    dst = draft_dir / "word" / "numbering.xml"
    shutil.copyfile(src, dst)

    _ensure_content_type_override(draft_dir, "/word/numbering.xml", CT_NUMBERING)
    _ensure_document_relationship(
        draft_dir, rel_type=REL_NUMBERING, target="numbering.xml")


def _copy_headers_footers(draft_dir: Path, template_dir: Path) -> None:
    """Replace the draft's header/footer references with copies of the
    template's. Must also move any images the headers/footers reference.
    """
    # 1. Strip existing header/footer parts + references from the draft.
    _strip_draft_headers_footers(draft_dir)

    # 2. For each template header/footer, copy the part, rewrite its rels
    #    to point to newly-copied images, register content types, and add
    #    a relationship + sectPr reference in the draft.
    template_word = template_dir / "word"
    draft_word = draft_dir / "word"

    header_footer_parts = sorted(
        p for p in template_word.iterdir()
        if p.is_file() and re.match(r"^(header|footer)\d+\.xml$", p.name)
    )

    doc_rels = _read_relationships(draft_word / "_rels" / "document.xml.rels")
    sect_prs = _collect_section_prs(draft_word / "document.xml")

    # Order: the draft's sections should point to the SAME header/footer
    # set in the SAME order. For simplicity we apply the same refs to every
    # section in the draft.
    new_header_refs: list[tuple[str, str]] = []   # (rel_id, type)
    new_footer_refs: list[tuple[str, str]] = []

    for part in header_footer_parts:
        kind = "header" if part.name.startswith("header") else "footer"
        ref_type = _sect_ref_type_from_part(part.name, template_word)

        # Copy the part XML itself.
        dst_part = draft_word / part.name
        shutil.copyfile(part, dst_part)

        # Copy its rels file + referenced images.
        src_rels = template_word / "_rels" / f"{part.name}.rels"
        copied_images: list[tuple[str, str]] = []  # (template_target, new_target_relative_to_word)
        if src_rels.exists():
            _copy_part_rels_and_media(
                src_rels, draft_word, draft_dir, template_dir,
                copied_images=copied_images,
            )
            dst_rels = draft_word / "_rels" / f"{part.name}.rels"
            dst_rels.parent.mkdir(exist_ok=True, parents=True)
            shutil.copyfile(src_rels, dst_rels)

        # Content-type override.
        ct = CT_HEADER if kind == "header" else CT_FOOTER
        _ensure_content_type_override(
            draft_dir, f"/word/{part.name}", ct)

        # Document relationship for this header/footer.
        rel_id = _add_document_relationship(
            doc_rels,
            rel_type=REL_HEADER if kind == "header" else REL_FOOTER,
            target=part.name,
        )

        if kind == "header":
            new_header_refs.append((rel_id, ref_type))
        else:
            new_footer_refs.append((rel_id, ref_type))

    _write_relationships(
        draft_word / "_rels" / "document.xml.rels", doc_rels)

    # Insert header/footer references into every sectPr in the draft.
    _inject_header_footer_refs(
        draft_word / "document.xml", new_header_refs, new_footer_refs,
        template_word / "document.xml",
    )


def _apply_page_setup(draft_dir: Path, template_dir: Path) -> None:
    """Copy the <w:pgMar>, <w:pgSz>, and first-page flags from the
    template's first section to every section in the draft."""
    tmpl_doc = _read_xml(template_dir / "word" / "document.xml")
    draft_doc_path = draft_dir / "word" / "document.xml"
    draft_doc = _read_xml(draft_doc_path)

    tmpl_sect = tmpl_doc.find(f".//{{{W_NS}}}sectPr")
    if tmpl_sect is None:
        return

    replaceable = [
        f"{{{W_NS}}}pgSz",
        f"{{{W_NS}}}pgMar",
        f"{{{W_NS}}}cols",
        f"{{{W_NS}}}docGrid",
        f"{{{W_NS}}}titlePg",
    ]

    for draft_sect in draft_doc.iter(f"{{{W_NS}}}sectPr"):
        for tag in replaceable:
            for old in list(draft_sect.findall(tag)):
                draft_sect.remove(old)
            for src in tmpl_sect.findall(tag):
                draft_sect.append(_clone_element(src))

    _write_xml(draft_doc_path, draft_doc)


# =====================================================================
# Helpers: zip I/O
# =====================================================================


def _unzip(src: Path, dst: Path) -> None:
    dst.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as z:
        z.extractall(dst)


def _rezip(src_dir: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    # ZIP_DEFLATED to match Word's default compression.
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(src_dir.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(src_dir).as_posix())


# =====================================================================
# Helpers: XML I/O
# =====================================================================


def _read_xml(path: Path) -> etree._Element:
    return etree.parse(str(path)).getroot()


def _write_xml(path: Path, root: etree._Element) -> None:
    tree = etree.ElementTree(root)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8",
               standalone=True)


def _clone_element(el: etree._Element) -> etree._Element:
    return etree.fromstring(etree.tostring(el))


# =====================================================================
# Helpers: [Content_Types].xml
# =====================================================================


def _ensure_content_type_override(root_dir: Path, part_name: str,
                                  content_type: str) -> None:
    ct_path = root_dir / "[Content_Types].xml"
    types = _read_xml(ct_path)

    for override in types.findall(f"{{{CT_NS}}}Override"):
        if override.get("PartName") == part_name:
            override.set("ContentType", content_type)
            _write_xml(ct_path, types)
            return

    new = etree.SubElement(types, f"{{{CT_NS}}}Override")
    new.set("PartName", part_name)
    new.set("ContentType", content_type)
    _write_xml(ct_path, types)


def _ensure_content_type_default(root_dir: Path, extension: str,
                                 content_type: str) -> None:
    ct_path = root_dir / "[Content_Types].xml"
    types = _read_xml(ct_path)

    for default in types.findall(f"{{{CT_NS}}}Default"):
        if default.get("Extension") == extension:
            return  # keep what's there
    new = etree.SubElement(types, f"{{{CT_NS}}}Default")
    new.set("Extension", extension)
    new.set("ContentType", content_type)
    _write_xml(ct_path, types)


# =====================================================================
# Helpers: relationships
# =====================================================================


def _read_relationships(path: Path) -> etree._Element:
    return _read_xml(path)


def _write_relationships(path: Path, root: etree._Element) -> None:
    _write_xml(path, root)


def _ensure_document_relationship(draft_dir: Path, rel_type: str,
                                  target: str) -> str:
    """Ensure document.xml.rels has a rel of `rel_type` pointing at
    `target` (path relative to /word/). Returns the relationship id."""
    rels_path = draft_dir / "word" / "_rels" / "document.xml.rels"
    rels = _read_relationships(rels_path)

    for rel in rels.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return rel.get("Id")

    new_id = _next_rel_id(rels)
    new_rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
    new_rel.set("Id", new_id)
    new_rel.set("Type", rel_type)
    new_rel.set("Target", target)
    _write_relationships(rels_path, rels)
    return new_id


def _add_document_relationship(doc_rels: etree._Element, rel_type: str,
                               target: str) -> str:
    new_id = _next_rel_id(doc_rels)
    new_rel = etree.SubElement(doc_rels, f"{{{REL_NS}}}Relationship")
    new_rel.set("Id", new_id)
    new_rel.set("Type", rel_type)
    new_rel.set("Target", target)
    return new_id


def _next_rel_id(rels: etree._Element) -> str:
    existing = {rel.get("Id") for rel in rels.findall(f"{{{REL_NS}}}Relationship")}
    i = 1
    while True:
        candidate = f"rId{i}"
        if candidate not in existing:
            return candidate
        i += 1


# =====================================================================
# Helpers: header/footer specifics
# =====================================================================


def _strip_draft_headers_footers(draft_dir: Path) -> None:
    """Remove the draft's existing header*.xml / footer*.xml parts,
    their rels files, their entries in document.xml.rels, and any
    headerReference/footerReference elements in document.xml.
    """
    draft_word = draft_dir / "word"
    doc_rels_path = draft_word / "_rels" / "document.xml.rels"

    if not doc_rels_path.exists():
        return

    rels = _read_relationships(doc_rels_path)
    stale_ids: set[str] = set()
    stale_files: set[str] = set()

    for rel in list(rels.findall(f"{{{REL_NS}}}Relationship")):
        t = rel.get("Type")
        if t in (REL_HEADER, REL_FOOTER):
            stale_ids.add(rel.get("Id"))
            stale_files.add(rel.get("Target"))
            rels.remove(rel)

    _write_relationships(doc_rels_path, rels)

    # Delete the header/footer part files and their .rels.
    for name in stale_files:
        part = draft_word / name
        if part.exists():
            part.unlink()
        part_rels = draft_word / "_rels" / f"{name}.rels"
        if part_rels.exists():
            part_rels.unlink()

        # Also scrub [Content_Types].xml override.
        _remove_content_type_override(draft_dir, f"/word/{name}")

    # Remove header/footer references from every sectPr in document.xml.
    doc_path = draft_word / "document.xml"
    if doc_path.exists():
        doc = _read_xml(doc_path)
        for ref_tag in (f"{{{W_NS}}}headerReference", f"{{{W_NS}}}footerReference"):
            for ref in list(doc.iter(ref_tag)):
                ref.getparent().remove(ref)
        _write_xml(doc_path, doc)


def _remove_content_type_override(root_dir: Path, part_name: str) -> None:
    ct_path = root_dir / "[Content_Types].xml"
    types = _read_xml(ct_path)
    for override in list(types.findall(f"{{{CT_NS}}}Override")):
        if override.get("PartName") == part_name:
            types.remove(override)
    _write_xml(ct_path, types)


def _sect_ref_type_from_part(part_name: str, template_word: Path) -> str:
    """Return default/first/even based on how the template's sectPr
    references this part. Falls back to 'default' if we can't determine."""
    doc = _read_xml(template_word / "document.xml")
    for ref in doc.iter():
        if ref.tag not in (f"{{{W_NS}}}headerReference", f"{{{W_NS}}}footerReference"):
            continue
        rid = ref.get(f"{{{R_NS}}}id")
        if not rid:
            continue
        # Find which part that rel points to.
        doc_rels = _read_relationships(template_word / "_rels" / "document.xml.rels")
        for rel in doc_rels.findall(f"{{{REL_NS}}}Relationship"):
            if rel.get("Id") == rid and rel.get("Target") == part_name:
                return ref.get(f"{{{W_NS}}}type", "default")
    return "default"


def _copy_part_rels_and_media(
    src_rels_path: Path,
    draft_word: Path,
    draft_root: Path,
    template_root: Path,
    copied_images: list[tuple[str, str]],
) -> None:
    """For a header/footer part's .rels file, copy any relative media
    targets (typically ../media/image1.png) from template into draft."""
    rels = _read_relationships(src_rels_path)
    for rel in rels.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Type") != REL_IMAGE:
            continue
        target = rel.get("Target", "")
        src_path = (src_rels_path.parent / ".." / target).resolve()
        if not src_path.exists():
            # path resolves relative to /word/_rels/ -> /word/
            src_path = (draft_word / target).resolve()
        # Real source is in the template tree; recompute.
        src_path = (template_root / "word" / target.lstrip("/")).resolve() \
            if not target.startswith("../") else \
            (template_root / "word" / target.replace("../", "")).resolve()

        if not src_path.exists():
            continue

        dst_path = (draft_word / target.replace("../", "")).resolve() \
            if target.startswith("../") else \
            (draft_word / target).resolve()
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        if not dst_path.exists():
            shutil.copyfile(src_path, dst_path)

        # Register default content type for the extension.
        ext = dst_path.suffix.lower().lstrip(".")
        _ensure_content_type_default(
            draft_root, ext, _image_content_type(ext))

        copied_images.append((target, str(dst_path.relative_to(draft_word))))


def _image_content_type(ext: str) -> str:
    return {
        "png": CT_IMAGE_PNG,
        "jpg": CT_IMAGE_JPEG,
        "jpeg": CT_IMAGE_JPEG,
        "gif": CT_IMAGE_GIF,
        "emf": CT_IMAGE_EMF,
        "wmf": CT_IMAGE_WMF,
    }.get(ext, CT_IMAGE_PNG)


def _collect_section_prs(document_xml_path: Path) -> list[etree._Element]:
    doc = _read_xml(document_xml_path)
    return list(doc.iter(f"{{{W_NS}}}sectPr"))


def _inject_header_footer_refs(
    draft_doc_path: Path,
    header_refs: list[tuple[str, str]],
    footer_refs: list[tuple[str, str]],
    template_doc_path: Path,
) -> None:
    """Mirror the template's sectPr header/footer reference layout onto
    every sectPr in the draft, using the new relationship IDs in the
    draft."""
    doc = _read_xml(draft_doc_path)

    # Read the template's first sectPr reference layout to know what
    # types (default/first/even) exist and their order.
    tmpl_doc = _read_xml(template_doc_path)
    tmpl_sect = tmpl_doc.find(f".//{{{W_NS}}}sectPr")
    if tmpl_sect is None:
        return

    layout: list[tuple[str, str, str]] = []  # (tag, type-attr, kind)
    for child in tmpl_sect:
        if child.tag == f"{{{W_NS}}}headerReference":
            layout.append(("header",
                           child.get(f"{{{W_NS}}}type", "default"),
                           "header"))
        elif child.tag == f"{{{W_NS}}}footerReference":
            layout.append(("footer",
                           child.get(f"{{{W_NS}}}type", "default"),
                           "footer"))

    def find_rel_for(kind: str, type_attr: str,
                     pool: list[tuple[str, str]]) -> str | None:
        for rel_id, ref_type in pool:
            if ref_type == type_attr:
                return rel_id
        # Fall back to default/any.
        return pool[0][0] if pool else None

    for sect in doc.iter(f"{{{W_NS}}}sectPr"):
        # Insert refs at the start of sectPr (Word expects them early).
        insert_at = 0
        for tag_key, type_attr, kind in layout:
            pool = header_refs if kind == "header" else footer_refs
            rid = find_rel_for(kind, type_attr, pool)
            if rid is None:
                continue
            ref = etree.Element(
                f"{{{W_NS}}}{kind}Reference", nsmap={"r": R_NS, "w": W_NS})
            ref.set(f"{{{W_NS}}}type", type_attr)
            ref.set(f"{{{R_NS}}}id", rid)
            sect.insert(insert_at, ref)
            insert_at += 1

    _write_xml(draft_doc_path, doc)

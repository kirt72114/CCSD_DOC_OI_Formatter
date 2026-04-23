"""Single source of truth for every USAF OI formatting rule.

Citations live in docs/rules.md. Sources:
  - AFH 33-337, The Tongue and Quill
  - DAFMAN 90-161, Publishing Processes and Procedures
"""

from __future__ import annotations

# ---- Fonts ----------------------------------------------------------
BODY_FONT = "Times New Roman"
BODY_SIZE_PT = 12

HEADING_FONT = "Times New Roman"
HEADING_SIZE_PT = 12

TITLEBLOCK_FONT = "Arial"
TITLEBLOCK_SIZE_PT = 10

# ---- Page setup -----------------------------------------------------
MARGIN_IN = 1.0
PAGE_WIDTH_IN = 8.5
PAGE_HEIGHT_IN = 11.0

# ---- Spacing --------------------------------------------------------
SPACE_AFTER_PT = 6

# ---- Numbering ------------------------------------------------------
MAX_NUMBER_DEPTH = 5

# ---- Bullet sequence (T&Q Ch. 10) ----------------------------------
BULLET_L1 = "-"
BULLET_L2 = "•"  # •
BULLET_L3 = "–"  # – (en dash)
BULLET_L4 = "»"  # »

# ---- Title-block labels (DAFMAN 90-161 Fig A2.2) -------------------
LBL_BYORDER = "BY ORDER OF THE COMMANDER"
LBL_COMPLIANCE = "COMPLIANCE WITH THIS PUBLICATION IS MANDATORY"
LBL_ACCESSIBILITY = "ACCESSIBILITY:"
LBL_RELEASABILITY = "RELEASABILITY:"
LBL_OPR = "OPR:"
LBL_SUPERSEDES = "Supersedes:"
LBL_CERTIFIED_BY = "Certified by:"
LBL_PAGES = "Pages:"

DEFAULT_ACCESSIBILITY = (
    "Publications and forms are available for downloading or ordering on "
    "the e-Publishing website at www.e-Publishing.af.mil."
)
DEFAULT_RELEASABILITY = (
    "There are no releasability restrictions on this publication."
)

# ---- Attachment conventions ----------------------------------------
ATTACH_PREFIX = "Attachment "
ATTACH_SEP = "—"  # em dash
GLOSSARY_TITLE = "GLOSSARY OF REFERENCES AND SUPPORTING INFORMATION"

# ---- Style names ---------------------------------------------------
STY_BODY = "OI Body"
STY_H1 = "OI Heading 1"
STY_H2 = "OI Heading 2"
STY_H3 = "OI Heading 3"
STY_H4 = "OI Heading 4"
STY_H5 = "OI Heading 5"
STY_TITLE = "OI Title"
STY_TITLEBLOCK = "OI TitleBlock"
STY_ATTACH_TITLE = "OI Attachment Title"
STY_BULLET_L1 = "OI Bullet 1"
STY_BULLET_L2 = "OI Bullet 2"
STY_BULLET_L3 = "OI Bullet 3"
STY_BULLET_L4 = "OI Bullet 4"

HEADING_STYLES = (STY_H1, STY_H2, STY_H3, STY_H4, STY_H5)
BULLET_STYLES = (STY_BULLET_L1, STY_BULLET_L2, STY_BULLET_L3, STY_BULLET_L4)


def heading_style_for_level(level: int) -> str:
    """Return the OI heading style name for a 1-indexed level (1..5)."""
    if 1 <= level <= len(HEADING_STYLES):
        return HEADING_STYLES[level - 1]
    return STY_BODY


def bullet_style_for_level(level: int) -> str:
    """Return the bullet style name for a 1-indexed level, clamped to 4."""
    level = max(1, min(level, len(BULLET_STYLES)))
    return BULLET_STYLES[level - 1]


def bullet_for_level(level: int) -> str:
    """Return the canonical bullet glyph for a 1-indexed level."""
    glyphs = (BULLET_L1, BULLET_L2, BULLET_L3, BULLET_L4)
    level = max(1, min(level, len(glyphs)))
    return glyphs[level - 1]


# ---- Output naming -------------------------------------------------
OUTPUT_SUFFIX = "_formatted"
REPORT_SUFFIX = "_changes.txt"

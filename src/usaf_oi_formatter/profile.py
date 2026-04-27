"""User-tunable formatting settings.

Everything here is something a unit could legitimately want to tweak per
local custom (margins, fonts, bullet glyphs, spacing). The defaults match
DAFMAN 90-161 + AFH 33-337 verbatim — picking the built-in "Tongue and
Quill" template gives the canonical settings.

DAFMAN-mandated *structural* constants (style names, attachment prefix,
title-block labels) live in `rules.py` and are not user-tunable.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

from . import rules


@dataclass
class FormattingProfile:
    """Every knob a user might twist. Defaults = strict T&Q / DAFMAN."""

    name: str = "Tongue and Quill"
    description: str = "AFH 33-337 / DAFMAN 90-161 default settings."

    # --- page setup ---------------------------------------------------
    margin_top_in: float = 1.0
    margin_bottom_in: float = 1.0
    margin_left_in: float = 1.0
    margin_right_in: float = 1.0
    page_width_in: float = 8.5
    page_height_in: float = 11.0

    # --- fonts --------------------------------------------------------
    body_font: str = "Times New Roman"
    body_size_pt: float = 12.0
    heading_font: str = "Times New Roman"
    heading_size_pt: float = 12.0
    titleblock_font: str = "Arial"
    titleblock_size_pt: float = 10.0
    title_size_pt: float = 14.0

    # --- spacing ------------------------------------------------------
    space_after_pt: float = 6.0
    heading_space_before_pt: float = 12.0
    sub_heading_space_before_pt: float = 6.0
    bullet_space_after_pt: float = 3.0
    bullet_indent_step_in: float = 0.25

    # --- bullet sequence (AFH 33-337 Ch. 10) --------------------------
    bullet_l1: str = "-"
    bullet_l2: str = "•"  # •
    bullet_l3: str = "–"  # – (en dash)
    bullet_l4: str = "»"  # »

    # --- numbering ----------------------------------------------------
    max_number_depth: int = 5
    number_indent_step_in: float = 0.25

    # --- page numbers -------------------------------------------------
    page_numbering_position: str = "bottom-center"  # bottom-center | bottom-right
    suppress_first_page_number: bool = True

    # --- title-block defaults (filled when user leaves field blank) ---
    default_accessibility: str = (
        "Publications and forms are available for downloading or ordering on "
        "the e-Publishing website at www.e-Publishing.af.mil."
    )
    default_releasability: str = (
        "There are no releasability restrictions on this publication."
    )

    # --- toggles ------------------------------------------------------
    seed_glossary: bool = True
    apply_hygiene: bool = True
    rebuild_title_block: bool = True

    # --- bookkeeping --------------------------------------------------
    is_builtin: bool = False
    source_path: str = ""  # populated when loaded from a json file

    # ---- helpers -----------------------------------------------------

    @classmethod
    def tongue_and_quill(cls) -> "FormattingProfile":
        """Strict T&Q / DAFMAN 90-161 defaults."""
        return cls(is_builtin=True)

    def bullet_for_level(self, level: int) -> str:
        glyphs = (self.bullet_l1, self.bullet_l2, self.bullet_l3, self.bullet_l4)
        level = max(1, min(level, len(glyphs)))
        return glyphs[level - 1]

    def heading_style_for_level(self, level: int) -> str:
        if 1 <= level <= len(rules.HEADING_STYLES):
            return rules.HEADING_STYLES[level - 1]
        return rules.STY_BODY

    def bullet_style_for_level(self, level: int) -> str:
        styles = rules.BULLET_STYLES
        level = max(1, min(level, len(styles)))
        return styles[level - 1]

    # ---- persistence -------------------------------------------------

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data.pop("is_builtin", None)
        data.pop("source_path", None)
        return data

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), indent=indent, sort_keys=False)

    def save(self, path: Path) -> Path:
        path = Path(path)
        path.write_text(self.to_json(), encoding="utf-8")
        out = self.copy()
        out.source_path = str(path)
        return path

    def copy(self, **overrides: Any) -> "FormattingProfile":
        data = asdict(self)
        data.update(overrides)
        return FormattingProfile(**data)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "FormattingProfile":
        known = {f.name for f in fields(cls)}
        clean = {k: v for k, v in data.items() if k in known}
        return cls(**clean)

    @classmethod
    def load(cls, path: Path) -> "FormattingProfile":
        path = Path(path)
        data = json.loads(path.read_text(encoding="utf-8"))
        prof = cls.from_dict(data)
        prof.source_path = str(path)
        prof.is_builtin = False
        return prof

    # ---- validation --------------------------------------------------

    def validate(self) -> list[str]:
        """Return human-readable warnings about settings that violate
        DAFMAN/T&Q. Empty list = nothing to flag."""
        warnings: list[str] = []
        if any(m < 0.5 or m > 2.0 for m in (
            self.margin_top_in, self.margin_bottom_in,
            self.margin_left_in, self.margin_right_in,
        )):
            warnings.append(
                "Margins outside 0.5”–2.0” are unusual; DAFMAN 90-161 expects 1”.")
        if (self.page_width_in, self.page_height_in) != (8.5, 11.0):
            warnings.append(
                "Page size is not 8.5 × 11 in; DAFMAN 90-161 publications use Letter.")
        if self.body_size_pt < 10 or self.body_size_pt > 14:
            warnings.append(
                "Body font size outside 10–14 pt is unusual for OIs.")
        if self.body_font.lower() not in ("times new roman", "arial", "calibri"):
            warnings.append(
                f"Body font '{self.body_font}' is non-standard; T&Q assumes Times New Roman.")
        if self.max_number_depth < 1 or self.max_number_depth > 9:
            warnings.append("Numbering depth must be 1–9 levels.")
        return warnings


# Convenient module-level shorthand for "give me the default profile".
def default() -> FormattingProfile:
    return FormattingProfile.tongue_and_quill()

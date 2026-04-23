"""OIMeta: the user-supplied metadata for the DAFMAN 90-161 title block."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass
class OIMeta:
    unit: str = ""                 # "442D MAINTENANCE SQUADRON"
    unit_short: str = ""           # "442 MXS"
    oi_number: str = ""            # "CCSD OI 36-1"
    date_str: str = ""             # "23 April 2026"
    category: str = ""             # "Personnel"
    subject: str = ""              # "Personnel Actions"
    opr: str = ""                  # "CCSD/CCC"
    supersedes: str = ""           # "CCSD OI 36-1, 1 Jan 2024"
    certified_by: str = ""         # "Col Jane Doe, Commander"
    pages: str = ""                # "12"
    accessibility: str = ""        # optional, rules.DEFAULT_ACCESSIBILITY if blank
    releasability: str = ""        # optional, rules.DEFAULT_RELEASABILITY if blank

    def with_defaults(self) -> "OIMeta":
        """Return a copy with sensible defaults filled in for blank fields."""
        out = OIMeta(**self.__dict__)
        if not out.date_str.strip():
            out.date_str = date.today().strftime("%-d %B %Y") if _supports_dash_d() \
                else date.today().strftime("%#d %B %Y") if _is_windows() \
                else date.today().strftime("%d %B %Y")
        return out


def _supports_dash_d() -> bool:
    import platform
    return platform.system() != "Windows"


def _is_windows() -> bool:
    import platform
    return platform.system() == "Windows"

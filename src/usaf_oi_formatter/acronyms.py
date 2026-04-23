"""Acronym collection.

Scans body paragraphs, tracks the first occurrence of each ALL-CAPS token
2..7 characters long, and seeds a glossary dictionary. We do NOT rewrite
first-use expansions automatically (too prone to mangling sentences);
downstream (attachments.py) emits the collected acronyms into the Glossary
attachment so a human reviewer can vet them.
"""

from __future__ import annotations

import re
from docx.document import Document

from . import rules

_RE_ACRONYM = re.compile(r"\b[A-Z][A-Z0-9&]{1,6}\b")

_NOISE = {
    "A", "I", "AN", "AT", "BE", "BY", "DO", "GO", "HE", "IF", "IN", "IS",
    "IT", "ME", "MY", "NO", "OF", "ON", "OR", "SO", "TO", "UP", "US", "WE",
    "THE", "AND", "FOR", "BUT", "NOT", "YOU", "ALL", "CAN",
}

_SEED: dict[str, str] = {
    "OI": "Operating Instruction",
    "AFI": "Air Force Instruction",
    "AFMAN": "Air Force Manual",
    "AFPD": "Air Force Policy Directive",
    "AFH": "Air Force Handbook",
    "DAFI": "Department of the Air Force Instruction",
    "DAFMAN": "Department of the Air Force Manual",
    "OPR": "Office of Primary Responsibility",
    "USAF": "United States Air Force",
    "DoD": "Department of Defense",
    "POC": "Point of Contact",
    "T&Q": "Tongue and Quill",
}

_EXCLUDED_STYLES = set(rules.HEADING_STYLES) | {
    rules.STY_TITLEBLOCK, rules.STY_TITLE, rules.STY_ATTACH_TITLE,
}


def collect(doc: Document) -> dict[str, str]:
    """Return a glossary mapping acronym -> expansion (or a TBD marker)."""
    glossary: dict[str, str] = dict(_SEED)
    seen: set[str] = set()

    for p in doc.paragraphs:
        if p.style.name in _EXCLUDED_STYLES:
            continue
        for match in _RE_ACRONYM.finditer(p.text):
            token = match.group(0)
            if token in _NOISE or token in seen:
                continue
            seen.add(token)
            glossary.setdefault(token, "TBD - define on first use")

    # Trim the seed dictionary back to only the acronyms that actually appear.
    return {k: glossary[k] for k in sorted(seen)} if seen else {}

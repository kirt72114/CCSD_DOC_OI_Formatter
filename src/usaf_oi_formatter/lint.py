"""Prose-level T&Q lint pass.

The structural formatter rewrites styles, margins, bullets, and the
title block. None of that catches *prose* problems — passive voice,
sentence length, gendered pronouns, T&Q "avoid" words, or acronyms
used before they're defined. Those are flagged here as suggestions
the human reviewer can act on.

Lint is read-only — it never modifies the document.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from docx.document import Document

from . import rules


# ---- finding model -------------------------------------------------

@dataclass(frozen=True)
class LintFinding:
    severity: str           # "info" | "warn"
    rule: str               # short rule id, e.g. "passive-voice"
    message: str            # human-readable
    paragraph_index: int    # 0-based; -1 for whole-doc findings
    snippet: str = ""       # ≤80-char excerpt

    def render(self) -> str:
        prefix = f"¶{self.paragraph_index + 1}" if self.paragraph_index >= 0 else "doc"
        bits = [f"[{self.severity.upper()}] {prefix} {self.rule}: {self.message}"]
        if self.snippet:
            bits.append(f'    “{self.snippet}”')
        return "\n".join(bits)


# ---- rule implementations ------------------------------------------

_BE_VERBS = r"(?:am|is|are|was|were|be|been|being)"
_PASSIVE_RE = re.compile(
    rf"\b{_BE_VERBS}\s+(?:\w+ly\s+)?(\w+ed|written|done|made|seen|taken|given|"
    r"shown|known|held|sent|kept|brought|built|driven|left|chosen|drawn)\b",
    re.IGNORECASE,
)

_SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+(?=[A-Z(])")

_GENDERED = {
    r"\bhe\b": "consider 'they' or restructure",
    r"\bhim\b": "consider 'them' or restructure",
    r"\bhis\b": "consider 'their' or restructure",
    r"\bshe\b": "consider 'they' or restructure",
    r"\bher\b": "consider 'them' or restructure",
    r"\bairman\b": "consider 'Airmen' (plural) or 'Airman/Guardian'",
    r"\bairmen\b": "OK; flagged for human review re: inclusivity",
}

# T&Q Ch. 4 "avoid these wordy phrases" sample.
_AVOID_PHRASES = {
    "utilize": "use",
    "in order to": "to",
    "prior to": "before",
    "subsequent to": "after",
    "due to the fact that": "because",
    "in the event that": "if",
    "for the purpose of": "to",
    "with regard to": "about",
    "at this point in time": "now",
    "in spite of the fact that": "although",
}

_RE_ACRONYM = re.compile(r"\b[A-Z][A-Z0-9&]{1,6}\b")
_NOISE_ACR = {
    "A", "I", "AN", "AT", "BE", "BY", "DO", "GO", "HE", "IF", "IN", "IS",
    "IT", "ME", "MY", "NO", "OF", "ON", "OR", "SO", "TO", "UP", "US", "WE",
    "THE", "AND", "FOR", "BUT", "NOT", "YOU", "ALL", "CAN", "OI",
}

# Acronyms whose first-use expansion is implied (in our seeded glossary).
_PRE_EXPANDED = {
    "AFI", "AFMAN", "AFPD", "AFH", "DAFI", "DAFMAN", "OPR", "USAF",
    "DoD", "POC", "T&Q",
}

_LONG_SENTENCE_WORDS = 30
_LONG_PARAGRAPH_WORDS = 120


_EXCLUDED_STYLES = set(rules.HEADING_STYLES) | {
    rules.STY_TITLEBLOCK, rules.STY_TITLE, rules.STY_ATTACH_TITLE,
}


def lint_document(doc: Document) -> list[LintFinding]:
    findings: list[LintFinding] = []
    seen_acronyms: set[str] = set()

    for idx, para in enumerate(doc.paragraphs):
        if para.style.name in _EXCLUDED_STYLES:
            continue
        text = para.text.strip()
        if not text:
            continue

        findings.extend(_check_passive(text, idx))
        findings.extend(_check_long_sentences(text, idx))
        findings.extend(_check_long_paragraph(text, idx))
        findings.extend(_check_gendered(text, idx))
        findings.extend(_check_avoid_phrases(text, idx))
        findings.extend(_check_first_use_acronyms(text, idx, seen_acronyms))

    return findings


# ---- individual rules ---------------------------------------------

def _snippet(text: str, match: re.Match, span: int = 30) -> str:
    start = max(0, match.start() - span)
    end = min(len(text), match.end() + span)
    out = text[start:end]
    if start > 0:
        out = "…" + out
    if end < len(text):
        out = out + "…"
    return out.replace("\n", " ")


def _check_passive(text: str, idx: int) -> Iterable[LintFinding]:
    for m in _PASSIVE_RE.finditer(text):
        yield LintFinding(
            severity="info",
            rule="passive-voice",
            message="Possible passive voice — T&Q Ch. 6 prefers active.",
            paragraph_index=idx,
            snippet=_snippet(text, m),
        )


def _check_long_sentences(text: str, idx: int) -> Iterable[LintFinding]:
    for sentence in _SENTENCE_SPLIT.split(text):
        words = sentence.split()
        if len(words) > _LONG_SENTENCE_WORDS:
            yield LintFinding(
                severity="info",
                rule="long-sentence",
                message=f"{len(words)}-word sentence; T&Q targets ≤{_LONG_SENTENCE_WORDS}.",
                paragraph_index=idx,
                snippet=" ".join(words[:12]) + "…",
            )


def _check_long_paragraph(text: str, idx: int) -> Iterable[LintFinding]:
    word_count = len(text.split())
    if word_count > _LONG_PARAGRAPH_WORDS:
        yield LintFinding(
            severity="info",
            rule="long-paragraph",
            message=f"{word_count}-word paragraph; consider splitting.",
            paragraph_index=idx,
            snippet=text[:80] + ("…" if len(text) > 80 else ""),
        )


def _check_gendered(text: str, idx: int) -> Iterable[LintFinding]:
    for pattern, hint in _GENDERED.items():
        for m in re.finditer(pattern, text, re.IGNORECASE):
            yield LintFinding(
                severity="info",
                rule="gendered-language",
                message=f"'{m.group(0)}' — {hint}.",
                paragraph_index=idx,
                snippet=_snippet(text, m),
            )


def _check_avoid_phrases(text: str, idx: int) -> Iterable[LintFinding]:
    lowered = text.lower()
    for phrase, suggestion in _AVOID_PHRASES.items():
        start = 0
        while True:
            pos = lowered.find(phrase, start)
            if pos < 0:
                break
            snippet_text = text[max(0, pos - 20):pos + len(phrase) + 20]
            yield LintFinding(
                severity="info",
                rule="wordy-phrase",
                message=f"'{phrase}' → '{suggestion}' (T&Q Ch. 4).",
                paragraph_index=idx,
                snippet=snippet_text.replace("\n", " "),
            )
            start = pos + len(phrase)


def _check_first_use_acronyms(
    text: str, idx: int, seen: set[str],
) -> Iterable[LintFinding]:
    for m in _RE_ACRONYM.finditer(text):
        token = m.group(0)
        if token in _NOISE_ACR or token in _PRE_EXPANDED or token in seen:
            seen.add(token)
            continue
        seen.add(token)
        # Heuristic: if the acronym is followed/preceded by a parenthetical
        # expansion, assume it is being defined here.
        if _looks_defined_inline(text, m):
            continue
        yield LintFinding(
            severity="warn",
            rule="undefined-acronym",
            message=f"'{token}' used without inline expansion (T&Q Ch. 8).",
            paragraph_index=idx,
            snippet=_snippet(text, m),
        )


def _looks_defined_inline(text: str, m: re.Match) -> bool:
    # Form 1: ACRONYM (Expansion Goes Here)
    after = text[m.end():m.end() + 80]
    if re.match(r"\s*\([A-Z][^)]{2,60}\)", after):
        return True
    # Form 2: Expansion Goes Here (ACRONYM)
    if m.start() > 0 and text[m.start() - 1] == "(" \
            and m.end() < len(text) and text[m.end()] == ")":
        return True
    return False


# ---- formatter -----------------------------------------------------

def write_report(findings: Iterable[LintFinding], path: Path) -> Path:
    lines = ["=== USAF OI Formatter — Tongue & Quill lint ===\n"]
    findings = list(findings)
    if not findings:
        lines.append("No prose-level issues detected.")
    else:
        lines.append(f"{len(findings)} suggestion(s):\n")
        for f in findings:
            lines.append(f.render())
            lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")
    return path

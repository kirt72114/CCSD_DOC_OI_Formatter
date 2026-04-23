# USAF OI Formatter — Rule Citations

Every constant in `src/usaf_oi_formatter/rules.py` traces back to one of the two
publications below. Update this file whenever a rule changes so the audit trail
stays intact.

Primary sources:
- **AFH 33-337**, *The Tongue and Quill* — style, voice, bullets, acronyms.
- **DAFMAN 90-161**, *Publishing Processes and Procedures for Air Force Publications* —
  publication mechanics: title block, paragraph numbering, margins, attachments.

| Constant | Value | Source |
|---|---|---|
| `BODY_FONT` / `BODY_SIZE_PT` | Times New Roman, 12 pt | DAFMAN 90-161, publication body text |
| `HEADING_FONT` / `HEADING_SIZE_PT` | Times New Roman, 12 pt bold | DAFMAN 90-161, heading specification |
| `TITLEBLOCK_FONT` / `TITLEBLOCK_SIZE_PT` | Arial, 10 pt | DAFMAN 90-161, cover/title block example |
| `MARGIN_IN` | 1.0 in all sides | DAFMAN 90-161, page setup |
| `PAGE_WIDTH_IN` / `PAGE_HEIGHT_IN` | 8.5 × 11 in | DAFMAN 90-161 |
| `SPACE_AFTER_PT` | 6 pt | DAFMAN 90-161 paragraph spacing |
| `MAX_NUMBER_DEPTH` | 5 levels (`1.`, `1.1.`, `1.1.1.`, `1.1.1.1.`, `1.1.1.1.1.`) | DAFMAN 90-161 paragraph numbering |
| `BULLET_L1..L4` | `-`, `•`, `–`, `»` | AFH 33-337 Ch. 10, bullet conventions |
| `LBL_BYORDER` | "BY ORDER OF THE COMMANDER" | DAFMAN 90-161 Fig A2.2 |
| `LBL_COMPLIANCE` | "COMPLIANCE WITH THIS PUBLICATION IS MANDATORY" | DAFMAN 90-161 Fig A2.2 |
| `LBL_ACCESSIBILITY` / `LBL_RELEASABILITY` | "ACCESSIBILITY:", "RELEASABILITY:" | DAFMAN 90-161 |
| `LBL_OPR` / `LBL_SUPERSEDES` / `LBL_CERTIFIED_BY` / `LBL_PAGES` | Standard labels | DAFMAN 90-161 Fig A2.2 |
| `ATTACH_PREFIX` / `ATTACH_SEP` | `Attachment N—TITLE` (em dash) | DAFMAN 90-161 attachment format |
| `GLOSSARY_TITLE` | "GLOSSARY OF REFERENCES AND SUPPORTING INFORMATION" | DAFMAN 90-161 |
| Acronym first-use rule | Collected and surfaced in Glossary; not auto-rewritten | AFH 33-337 Ch. 8 |
| Straight quotes / apostrophes | Smart quotes replaced with straight | DAFMAN 90-161 text hygiene |
| Widow/orphan control | On | AFH 33-337 editing standards |
| Page numbering | Bottom center, first page suppressed | DAFMAN 90-161 pagination |

## Verifying against updated publications

When AFH 33-337 or DAFMAN 90-161 is revised:

1. Note the new specification and the paragraph reference in this table.
2. Update the corresponding constant in `src/usaf_oi_formatter/rules.py`.
3. Run `pytest` to confirm nothing regressed.
4. Re-run the tool against `tests/samples/` and open the output in Word.
5. Commit the rule change and the doc update together.

## Known gaps

These items are **not** auto-enforced — the formatter cannot assess them
reliably and a human reviewer must still check:

- Tone and voice (active vs. passive, plain language).
- Completeness of the OPR, Supersedes, and Certified-by values (we accept
  whatever the user types).
- Semantic accuracy of acronym expansions — we collect acronyms into the
  Glossary but only fill in expansions we have seeded in `acronyms.py`;
  unknown ones are marked "TBD - define on first use" for the reviewer.
- References format (AFI/AFMAN numbers, dates, URLs).
- Classification markings.
- Signature block signed by the correct individual.

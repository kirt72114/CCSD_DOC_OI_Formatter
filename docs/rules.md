# USAF OI Formatter — Rule Citations

Every constant in `src/vba/modRules.bas` traces back to one of the two
publications below. Update this file whenever a rule changes so the audit trail
stays intact.

Primary sources:
- **AFH 33-337**, *The Tongue and Quill* — style, voice, bullets, acronyms.
- **DAFMAN 90-161**, *Publishing Processes and Procedures for Air Force Publications* —
  publication mechanics: title block, paragraph numbering, margins, attachments.

| Constant | Value | Source |
|---|---|---|
| `R_BODY_FONT` / `R_BODY_SIZE` | Times New Roman, 12 pt | DAFMAN 90-161, publication body text |
| `R_HEADING_FONT` / `R_HEADING_SIZE` | Times New Roman, 12 pt bold | DAFMAN 90-161, heading specification |
| `R_TITLEBLOCK_FONT` / `R_TITLEBLOCK_SIZE` | Arial, 10 pt | DAFMAN 90-161, cover/title block example |
| `R_MARGIN_IN` | 1.0 in all sides | DAFMAN 90-161, page setup |
| `R_PAGE_WIDTH_IN` / `R_PAGE_HEIGHT_IN` | 8.5 × 11 in | DAFMAN 90-161 |
| `R_LINE_SPACING_RULE` | Single | AFH 33-337 default; DAFMAN 90-161 body spec |
| `R_SPACE_AFTER_PT` | 6 pt | DAFMAN 90-161 paragraph spacing |
| `R_MAX_NUMBER_DEPTH` | 5 levels (`1.`, `1.1.`, `1.1.1.`, `1.1.1.1.`, `1.1.1.1.1.`) | DAFMAN 90-161 paragraph numbering |
| `R_BULLET_L1..L4` | `-`, `•`, `–`, `»` | AFH 33-337 Ch. 10, bullet conventions |
| `R_LBL_BYORDER` | "BY ORDER OF THE COMMANDER" | DAFMAN 90-161 Fig A2.2 |
| `R_LBL_COMPLIANCE` | "COMPLIANCE WITH THIS PUBLICATION IS MANDATORY" | DAFMAN 90-161 Fig A2.2 |
| `R_LBL_ACCESSIBILITY` / `R_LBL_RELEASABILITY` | "ACCESSIBILITY:", "RELEASABILITY:" | DAFMAN 90-161 |
| `R_LBL_OPR` / `R_LBL_SUPERSEDES` / `R_LBL_CERTIFIED_BY` / `R_LBL_PAGES` | Standard labels | DAFMAN 90-161 Fig A2.2 |
| `R_ATTACH_PREFIX` / `R_ATTACH_SEP` | `Attachment N—TITLE` (em dash) | DAFMAN 90-161 attachment format |
| `R_GLOSSARY_TITLE` | "GLOSSARY OF REFERENCES AND SUPPORTING INFORMATION" | DAFMAN 90-161 |
| Acronym first-use rule | Spell out, then parenthesize; reuse bare | AFH 33-337 Ch. 8 |
| Straight quotes / apostrophes | Smart quotes replaced with straight | DAFMAN 90-161 text hygiene |
| Widow/orphan control | On | AFH 33-337 editing standards |
| Page numbering | Bottom center, suppress first page | DAFMAN 90-161 pagination |

## Verifying against updated publications

When AFH 33-337 or DAFMAN 90-161 is revised:

1. Note the new specification and the paragraph reference in this table.
2. Update the corresponding constant in `src/vba/modRules.bas`.
3. Rebuild with `build\build-dotm.ps1` and re-run the sample test in
   `tests\samples\` to confirm the output reflects the change.
4. Commit the rule change and the doc update together.

## Known gaps

These items are **not** auto-enforced — the formatter cannot assess them
reliably and a human reviewer must still check:

- Tone and voice (active vs. passive, plain language).
- Completeness of the OPR, Supersedes, and Certified-by values (we accept
  whatever the user types).
- Semantic accuracy of acronym expansions (we only flag and collect).
- References format (AFI/AFMAN numbers, dates, URLs).
- Classification markings.
- Signature block signed by the correct individual.

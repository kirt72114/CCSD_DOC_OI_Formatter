import pytest

docx = pytest.importorskip("docx")

from usaf_oi_formatter import lint


def _doc_with(*paragraphs: str):
    d = docx.Document()
    for p in paragraphs:
        d.add_paragraph(p)
    return d


def test_clean_paragraph_has_no_findings():
    d = _doc_with("This OI defines local procedures for personnel actions.")
    findings = lint.lint_document(d)
    assert all(f.rule != "passive-voice" for f in findings)


def test_passive_voice_flagged():
    d = _doc_with("The report was written by the airman.")
    findings = lint.lint_document(d)
    assert any(f.rule == "passive-voice" for f in findings)


def test_long_sentence_flagged():
    text = " ".join(["word"] * 35) + "."
    d = _doc_with(text)
    findings = lint.lint_document(d)
    assert any(f.rule == "long-sentence" for f in findings)


def test_wordy_phrase_flagged():
    d = _doc_with("Personnel will utilize the system in order to comply.")
    findings = lint.lint_document(d)
    rules_seen = {f.rule for f in findings}
    assert "wordy-phrase" in rules_seen


def test_undefined_acronym_flagged():
    d = _doc_with("Submit the SORTS data quarterly.")
    findings = lint.lint_document(d)
    assert any(f.rule == "undefined-acronym" for f in findings)


def test_inline_definition_suppresses_acronym_warning():
    d = _doc_with("Submit the Status of Resources and Training System (SORTS) data quarterly.")
    findings = lint.lint_document(d)
    assert not any(
        f.rule == "undefined-acronym" and "SORTS" in f.message
        for f in findings
    )

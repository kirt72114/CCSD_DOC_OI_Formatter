from pathlib import Path

from usaf_oi_formatter.profile import FormattingProfile
from usaf_oi_formatter import templates


def test_default_profile_matches_t_and_q():
    p = FormattingProfile.tongue_and_quill()
    assert p.name == "Tongue and Quill"
    assert p.body_font == "Times New Roman"
    assert p.body_size_pt == 12
    assert p.margin_top_in == 1.0
    assert p.bullet_l1 == "-"
    assert p.max_number_depth == 5
    assert p.is_builtin is True


def test_validate_flags_unusual_margins():
    p = FormattingProfile(margin_top_in=0.1)
    warnings = p.validate()
    assert any("Margins" in w for w in warnings)


def test_validate_flags_non_letter_paper():
    p = FormattingProfile(page_width_in=11.0, page_height_in=17.0)
    warnings = p.validate()
    assert any("8.5" in w for w in warnings)


def test_default_profile_has_no_warnings():
    assert FormattingProfile.tongue_and_quill().validate() == []


def test_copy_overrides_field():
    p = FormattingProfile.tongue_and_quill()
    q = p.copy(body_font="Calibri")
    assert q.body_font == "Calibri"
    assert p.body_font == "Times New Roman"  # original untouched


def test_round_trip_save_load(tmp_path: Path):
    p = FormattingProfile.tongue_and_quill().copy(
        name="Local OI", body_size_pt=11.5, bullet_l1="*",
    )
    out = tmp_path / "profile.json"
    p.save(out)
    q = FormattingProfile.load(out)
    assert q.name == "Local OI"
    assert q.body_size_pt == 11.5
    assert q.bullet_l1 == "*"
    assert q.source_path == str(out)


def test_bullet_for_level_clamps():
    p = FormattingProfile.tongue_and_quill()
    assert p.bullet_for_level(1) == p.bullet_l1
    assert p.bullet_for_level(4) == p.bullet_l4
    assert p.bullet_for_level(99) == p.bullet_l4
    assert p.bullet_for_level(0) == p.bullet_l1


def test_templates_all_loadable():
    for name in templates.builtin_names():
        p = templates.get_builtin(name)
        assert p.name
        assert p.is_builtin


def test_compact_template_has_tighter_spacing():
    base = templates.get_builtin("Tongue and Quill")
    compact = templates.get_builtin("Compact")
    assert compact.space_after_pt < base.space_after_pt
    assert compact.heading_space_before_pt < base.heading_space_before_pt

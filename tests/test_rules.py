from usaf_oi_formatter import rules


def test_heading_style_for_level_in_range():
    assert rules.heading_style_for_level(1) == rules.STY_H1
    assert rules.heading_style_for_level(5) == rules.STY_H5


def test_heading_style_for_level_out_of_range():
    assert rules.heading_style_for_level(0) == rules.STY_BODY
    assert rules.heading_style_for_level(6) == rules.STY_BODY


def test_bullet_for_level_clamps():
    assert rules.bullet_for_level(1) == rules.BULLET_L1
    assert rules.bullet_for_level(4) == rules.BULLET_L4
    assert rules.bullet_for_level(99) == rules.BULLET_L4
    assert rules.bullet_for_level(-1) == rules.BULLET_L1


def test_bullet_style_for_level_clamps():
    assert rules.bullet_style_for_level(1) == rules.STY_BULLET_L1
    assert rules.bullet_style_for_level(4) == rules.STY_BULLET_L4
    assert rules.bullet_style_for_level(99) == rules.STY_BULLET_L4

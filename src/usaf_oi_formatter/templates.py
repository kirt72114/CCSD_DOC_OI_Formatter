"""Built-in formatting templates.

A template is just a `FormattingProfile` with a non-default name and
preset values that match a particular publication style. Picking a
template in the UI loads its values into the editable settings — users
can then tweak anything before running.

Three built-ins ship today:
  - "Tongue and Quill"  : strict DAFMAN 90-161 / AFH 33-337 (default)
  - "Approved OI"       : same as T&Q + a few common unit overrides
  - "Compact"           : tighter spacing for shorter OIs

Custom templates can be saved/loaded from JSON via FormattingProfile.save / .load.
"""

from __future__ import annotations

from .profile import FormattingProfile


def _t_and_q() -> FormattingProfile:
    return FormattingProfile.tongue_and_quill()


def _approved_oi() -> FormattingProfile:
    return FormattingProfile(
        name="Approved OI",
        description=(
            "Tongue and Quill defaults with the unit OI conventions: "
            "title-block fully populated, glossary always seeded."
        ),
        seed_glossary=True,
        rebuild_title_block=True,
        is_builtin=True,
    )


def _compact() -> FormattingProfile:
    return FormattingProfile(
        name="Compact",
        description=(
            "Tighter spacing for short OIs. Same fonts/margins as T&Q, "
            "reduced paragraph spacing."
        ),
        space_after_pt=3.0,
        heading_space_before_pt=8.0,
        sub_heading_space_before_pt=4.0,
        bullet_space_after_pt=2.0,
        is_builtin=True,
    )


_BUILTIN_FACTORIES = {
    "Tongue and Quill": _t_and_q,
    "Approved OI": _approved_oi,
    "Compact": _compact,
}


def builtin_names() -> list[str]:
    return list(_BUILTIN_FACTORIES)


def get_builtin(name: str) -> FormattingProfile:
    """Return a fresh copy of the named built-in. KeyError if unknown."""
    return _BUILTIN_FACTORIES[name]()


def default() -> FormattingProfile:
    return get_builtin("Tongue and Quill")

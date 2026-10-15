"""Page setups for programmatic documents.

`tdj` is the optional bilingual journal layout. The others are everyday documents.
`plain` is the same setup as `note`, kept so older commands still work.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TemplateProfile:
    key: str
    one_block: bool
    checklist: bool
    show_udc: bool
    labeled_front: bool
    font: str
    body_pt: float
    margin_in: float
    line_spacing: float
    first_indent_in: float
    body_align: str
    title_align: str


def _profile(
    key: str,
    *,
    one_block: bool = True,
    checklist: bool = False,
    show_udc: bool = False,
    labeled_front: bool = True,
    font: str = "Calibri",
    body_pt: float = 11,
    margin_in: float = 1.0,
    line_spacing: float = 1.15,
    first_indent_in: float = 0,
    body_align: str = "left",
    title_align: str = "left",
) -> TemplateProfile:
    return TemplateProfile(
        key=key,
        one_block=one_block,
        checklist=checklist,
        show_udc=show_udc,
        labeled_front=labeled_front,
        font=font,
        body_pt=body_pt,
        margin_in=margin_in,
        line_spacing=line_spacing,
        first_indent_in=first_indent_in,
        body_align=body_align,
        title_align=title_align,
    )


_ARTICLE = _profile(
    "article",
    show_udc=True,
    font="Times New Roman",
    body_pt=14,
    margin_in=0.79,
    line_spacing=1.5,
    first_indent_in=0.49,
    body_align="justify",
    title_align="center",
)

PROFILES = {
    "note": _profile("note", show_udc=True),
    "letter": _profile(
        "letter",
        labeled_front=False,
        body_pt=12,
        margin_in=1.15,
        line_spacing=1.08,
    ),
    "memo": _profile(
        "memo",
        labeled_front=False,
        margin_in=0.85,
        line_spacing=1.08,
    ),
    "report": _profile(
        "report",
        show_udc=True,
        font="Times New Roman",
        body_pt=12,
        margin_in=0.9,
        line_spacing=1.15,
        first_indent_in=0.3,
        body_align="justify",
        title_align="center",
    ),
    "article": _ARTICLE,
    "plain": _profile("plain", show_udc=True),
    "tdj": _profile(
        "tdj",
        one_block=False,
        checklist=True,
        show_udc=True,
        font="Times New Roman",
        body_pt=14,
        margin_in=0.79,
        line_spacing=1.5,
        first_indent_in=0.49,
        body_align="justify",
        title_align="center",
    ),
}

TEMPLATES = tuple(PROFILES)


def profile_for(name: str | None) -> TemplateProfile:
    key = (name or "note").strip().lower()
    if key not in PROFILES:
        known = ", ".join(TEMPLATES)
        raise ValueError(f"Unknown template {name!r}. Use one of: {known}.")
    return PROFILES[key]

"""authors.json helpers."""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Sequence

from .constants import SUPERSCRIPTS, resolve_authors_json
from .models import Paper

def to_superscript(text: str) -> str:
    return text.translate(SUPERSCRIPTS)



def load_authors_db(path: Optional[Path] = None) -> dict:
    path = Path(path) if path else resolve_authors_json()
    if not path.exists():
        return {"authors": [], "organizations": {}}
    return json.loads(path.read_text(encoding="utf-8"))



def detect_author_ids(paper: "Paper", db: dict) -> List[str]:
    blob = " ".join(
        [
            paper.authors_ru,
            paper.authors_en,
            paper.affiliations_ru,
            paper.affiliations_en,
            " ".join(paper.authors_info),
        ]
    )
    found = []
    for author in db.get("authors", []):
        keys = [
            author.get("ru_full", ""),
            author.get("ru_short", ""),
            author.get("en_short", ""),
            (author.get("ru_full") or "").split()[0] if author.get("ru_full") else "",
        ]
        if any(k and k in blob for k in keys):
            found.append(author["id"])
    return found



def _tex_authors_look_careful(paper: "Paper") -> bool:
    """True if .tex already has superscript affiliation markers — don't clobber.

    Long affiliation text alone is not enough: without marks like ¹,² authors.json
    should still inject numbered orgs so Word shows who belongs where.
    """
    supers = set("⁰¹²³⁴⁵⁶⁷⁸⁹")
    blob = f"{paper.authors_ru or ''}{paper.affiliations_ru or ''}"
    return any(c in blob for c in supers)



def apply_authors_db(paper: "Paper", db: Optional[dict] = None, author_ids: Optional[Sequence[str]] = None) -> None:
    """Fill authors / affiliations / info from authors.json with superscript markers.

    Explicit ``author_ids`` (CLI ``--authors``) always overwrites.
    Auto-detect only fills empty fields when .tex authors already look careful.
    """
    db = db or load_authors_db()
    authors = {a["id"]: a for a in db.get("authors", [])}
    orgs = db.get("organizations", {})
    explicit = author_ids is not None
    ids = list(author_ids) if explicit else detect_author_ids(paper, db)
    ids = [i for i in ids if i in authors]
    if not ids:
        return

    # Stable org numbering across selected authors
    org_order: List[str] = []
    for aid in ids:
        for oid in authors[aid].get("org_ids", []):
            if oid not in org_order and oid in orgs:
                org_order.append(oid)
    org_num = {oid: i + 1 for i, oid in enumerate(org_order)}

    def marks(aid: str) -> str:
        nums = [org_num[oid] for oid in authors[aid].get("org_ids", []) if oid in org_num]
        return to_superscript(",".join(str(n) for n in nums)) if nums else ""

    authors_ru = ", ".join(f"{authors[aid]['ru_short']}{marks(aid)}" for aid in ids)
    authors_en = ", ".join(f"{authors[aid]['en_short']}{marks(aid)}" for aid in ids)

    aff_ru, aff_en = [], []
    for oid in org_order:
        n = to_superscript(str(org_num[oid]))
        aff_ru.append(f"{n} {orgs[oid].get('ru', '')}".strip())
        aff_en.append(f"{n} {orgs[oid].get('en', '')}".strip())
    affiliations_ru = "\n".join(aff_ru)
    affiliations_en = "\n".join(aff_en)

    infos = []
    for aid in ids:
        info = authors[aid].get("info_ru")
        if info:
            infos.append(info)

    preserve = (not explicit) and _tex_authors_look_careful(paper)
    if preserve:
        if not paper.authors_en:
            paper.authors_en = authors_en
        if not paper.affiliations_en:
            paper.affiliations_en = affiliations_en
        if not paper.authors_info and infos:
            paper.authors_info = infos
        return

    paper.authors_ru = authors_ru
    paper.authors_en = authors_en
    paper.affiliations_ru = affiliations_ru
    paper.affiliations_en = affiliations_en
    if infos:
        paper.authors_info = infos


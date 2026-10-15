"""High-level convert + TD-J checklist."""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Sequence, Union

from .authors import apply_authors_db, load_authors_db
from .constants import resolve_authors_json
from .docx_builder import DocxBuilder
from .md_parser import MarkdownParser
from .models import Heading, Paper
from .parser import LatexParser
from .util import ensure_out_prefix

def tdj_compliance_notes(paper: Paper) -> List[str]:
    """Checklist vs https://td-j.ru/ — what the converter cannot fully guarantee."""
    notes = []
    if not paper.udc:
        notes.append("missing UDC (TD-J 2.5)")
    if not paper.title_ru:
        notes.append("missing Russian title (TD-J 2.4)")
    if not paper.title_en:
        notes.append("missing English title (TD-J 2.4)")
    if not paper.authors_ru:
        notes.append("missing Russian authors (TD-J 2.4)")
    if not paper.authors_en:
        notes.append("missing English author names (TD-J 2.4); add them in the .tex or by hand")
    if not paper.affiliations_en:
        notes.append("missing English affiliation (TD-J 2.4)")
    if not paper.abstract_ru:
        notes.append("missing Russian abstract (TD-J 2.4)")
    elif len(paper.abstract_ru.split()) < 100:
        notes.append(
            f"Russian abstract is shorter than 100 words ({len(paper.abstract_ru.split())}); TD-J asks for 100–250"
        )
    elif len(paper.abstract_ru.split()) > 250:
        notes.append(f"Russian abstract is longer than 250 words ({len(paper.abstract_ru.split())})")
    if not paper.abstract_en:
        notes.append("missing English abstract (TD-J 2.4)")
    if not paper.keywords_ru:
        notes.append("missing Russian keywords (TD-J 2.4)")
    else:
        nkw = len(re.split(r"[,;]", paper.keywords_ru))
        if nkw > 10:
            notes.append(f"more than 10 Russian keywords ({nkw})")
    if not paper.keywords_en:
        notes.append("missing English keywords (TD-J 2.4)")
    if not paper.references:
        notes.append("missing reference list (TD-J 2.8)")
    if not paper.authors_info:
        notes.append(
            "missing author-information block (TD-J 2.11: degree, position, addresses, phones, e-mail)"
        )
    headings = [b.text.lower() for b in paper.body if isinstance(b, Heading)]
    if not any("введен" in h for h in headings):
        notes.append('missing a section titled "Введение" (TD-J 2.6)')
    if not any("заключен" in h or "вывод" in h for h in headings):
        notes.append('missing a section titled "Заключение" or "Выводы" (TD-J 2.6)')
    return notes



def latex_to_docx(
    tex_path: Union[str, Path],
    out_path: Optional[Union[str, Path]] = None,
    template: str = "tdj",
    authors_json: Optional[Union[str, Path]] = None,
    author_ids: Optional[Sequence[str]] = None,
    checklist: bool = True,
) -> Path:
    tex_path = Path(tex_path).resolve()
    if not tex_path.exists():
        raise FileNotFoundError(tex_path)
    out_path = (
        ensure_out_prefix(out_path)
        if out_path
        else Path(f"out_{tex_path.stem}.docx")
    )
    paper = LatexParser(tex_path).parse()

    db_path = resolve_authors_json(authors_json)
    apply_authors_db(paper, load_authors_db(db_path), author_ids=author_ids)

    builder = DocxBuilder()
    builder.build(paper, template=template)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    builder.doc.save(str(out_path))

    if checklist and template == "tdj":
        notes = tdj_compliance_notes(paper)
        if notes:
            print("TD-J checklist (fix these in the .tex, then convert again):")
            for n in notes:
                print(f"  • {n}")
        else:
            print("TD-J checklist: front matter and section headings are present.")
    return out_path


def markdown_to_docx(
    md_path: Union[str, Path],
    out_path: Optional[Union[str, Path]] = None,
) -> Path:
    """Markdown → DOCX with the same builder as LaTeX (plain template, no TD-J front matter)."""
    md_path = Path(md_path).resolve()
    if not md_path.exists():
        raise FileNotFoundError(md_path)
    out_path = (
        ensure_out_prefix(out_path)
        if out_path
        else Path(f"out_{md_path.stem}.docx")
    )
    paper = MarkdownParser(md_path).parse()
    builder = DocxBuilder()
    builder.build(paper, template="plain")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    builder.doc.save(str(out_path))
    return out_path




"""High-level convert + TD-J checklist."""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Sequence, Union

from .authors import apply_authors_db, load_authors_db
from .constants import resolve_authors_json
from .docx_builder import DocxBuilder
from .i18n import checklist_header, checklist_line, normalize_lang
from .md_parser import MarkdownParser
from .models import Heading, Paper
from .parser import LatexParser
from .util import ensure_out_prefix

def tdj_compliance_notes(paper: Paper, lang: str = "ru") -> List[str]:
    """Checklist vs https://td-j.ru/ — what the converter cannot fully guarantee."""
    lang = normalize_lang(lang)
    notes = []
    line = lambda key, **kwargs: notes.append(checklist_line(lang, key, **kwargs))
    if not paper.udc:
        line("udc")
    if not paper.title_ru:
        line("title_ru")
    if not paper.title_en:
        line("title_en")
    if not paper.authors_ru:
        line("authors_ru")
    if not paper.authors_en:
        line("authors_en")
    if not paper.affiliations_en:
        line("aff_en")
    if not paper.abstract_ru:
        line("abs_ru")
    elif len(paper.abstract_ru.split()) < 100:
        line("abs_ru_short", n=len(paper.abstract_ru.split()))
    elif len(paper.abstract_ru.split()) > 250:
        line("abs_ru_long", n=len(paper.abstract_ru.split()))
    if not paper.abstract_en:
        line("abs_en")
    if not paper.keywords_ru:
        line("kw_ru")
    else:
        nkw = len(re.split(r"[,;]", paper.keywords_ru))
        if nkw > 10:
            line("kw_ru_many", n=nkw)
    if not paper.keywords_en:
        line("kw_en")
    if not paper.references:
        line("refs")
    if not paper.authors_info:
        line("authors_info")
    headings = [b.text.lower() for b in paper.body if isinstance(b, Heading)]
    if not any("введен" in h or "introduction" in h for h in headings):
        line("intro")
    if not any(token in h for h in headings for token in ("заключен", "вывод", "conclusion")):
        line("conclusion")
    return notes



def latex_to_docx(
    tex_path: Union[str, Path],
    out_path: Optional[Union[str, Path]] = None,
    template: str = "note",
    authors_json: Optional[Union[str, Path]] = None,
    author_ids: Optional[Sequence[str]] = None,
    checklist: bool = True,
    lang: str = "ru",
    docx_template: Optional[Union[str, Path]] = None,
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

    lang = normalize_lang(lang)
    builder = DocxBuilder(lang=lang, style_template=docx_template, template=template)
    builder.build(paper, template=template)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    builder.doc.save(str(out_path))

    if checklist and builder.profile.checklist:
        notes = tdj_compliance_notes(paper, lang=lang)
        print(checklist_header(lang, ok=not notes))
        for n in notes:
            print(f"  • {n}")
    return out_path


def markdown_to_docx(
    md_path: Union[str, Path],
    out_path: Optional[Union[str, Path]] = None,
    lang: str = "ru",
    docx_template: Optional[Union[str, Path]] = None,
    template: str = "note",
) -> Path:
    """Markdown → DOCX. template picks the page setup (note, letter, memo, report, article)."""
    md_path = Path(md_path).resolve()
    if not md_path.exists():
        raise FileNotFoundError(md_path)
    out_path = (
        ensure_out_prefix(out_path)
        if out_path
        else Path(f"out_{md_path.stem}.docx")
    )
    paper = MarkdownParser(md_path).parse()
    builder = DocxBuilder(lang=lang, style_template=docx_template, template=template)
    builder.build(paper, template=template)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    builder.doc.save(str(out_path))
    return out_path




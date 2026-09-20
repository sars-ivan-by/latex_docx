"""High-level convert + TD-J checklist."""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Sequence, Union

from .authors import apply_authors_db, load_authors_db
from .constants import resolve_authors_json
from .docx_builder import DocxBuilder
from .models import Heading, Paper
from .parser import LatexParser
from .util import ensure_out_prefix

def tdj_compliance_notes(paper: Paper) -> List[str]:
    """Checklist vs https://td-j.ru/ — what the converter cannot fully guarantee."""
    notes = []
    if not paper.udc:
        notes.append("нет УДК (п. 2.5)")
    if not paper.title_ru:
        notes.append("нет заглавия на русском (п. 2.4)")
    if not paper.title_en:
        notes.append("нет заглавия на английском (п. 2.4)")
    if not paper.authors_ru:
        notes.append("нет авторов на русском (п. 2.4)")
    if not paper.authors_en:
        notes.append("нет ФИО авторов на английском (п. 2.4) — добавьте в .tex или вручную")
    if not paper.affiliations_en:
        notes.append("нет названия организации на английском (п. 2.4)")
    if not paper.abstract_ru:
        notes.append("нет аннотации на русском (п. 2.4)")
    elif len(paper.abstract_ru.split()) < 100:
        notes.append(f"русская аннотация короче 100 слов ({len(paper.abstract_ru.split())}) — TD-J: 100–250")
    elif len(paper.abstract_ru.split()) > 250:
        notes.append(f"русская аннотация длиннее 250 слов ({len(paper.abstract_ru.split())})")
    if not paper.abstract_en:
        notes.append("нет аннотации на английском (п. 2.4)")
    if not paper.keywords_ru:
        notes.append("нет ключевых слов на русском (п. 2.4)")
    else:
        nkw = len(re.split(r"[,;]", paper.keywords_ru))
        if nkw > 10:
            notes.append(f"ключевых слов RU > 10 ({nkw})")
    if not paper.keywords_en:
        notes.append("нет keywords на английском (п. 2.4)")
    if not paper.references:
        notes.append("нет списка литературы (п. 2.8)")
    if not paper.authors_info:
        notes.append(
            "нет блока «Сведения об авторах» (п. 2.11: степень, должность, адреса, телефоны, e-mail)"
        )
    headings = [b.text.lower() for b in paper.body if isinstance(b, Heading)]
    if not any("введен" in h for h in headings):
        notes.append("нет раздела «Введение» (п. 2.6)")
    if not any("заключен" in h or "вывод" in h for h in headings):
        notes.append("нет раздела «Заключение» (п. 2.6)")
    return notes



def latex_to_docx(
    tex_path: Union[str, Path],
    out_path: Optional[Union[str, Path]] = None,
    template: str = "tdj",
    authors_json: Optional[Union[str, Path]] = None,
    author_ids: Optional[Sequence[str]] = None,
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

    if template == "tdj":
        notes = tdj_compliance_notes(paper)
        if notes:
            print("TD-J checklist (нужно проверить вручную / дополнить в .tex):")
            for n in notes:
                print(f"  • {n}")
        else:
            print("TD-J checklist: базовые поля front matter / структура на месте.")
    return out_path




"""Публичная обёртка: один вызов для .tex и .md."""
from __future__ import annotations

from pathlib import Path
from typing import Optional, Sequence, Union

from .convert import latex_to_docx, markdown_to_docx
from .diff import write_docx_diff_report
from .pptx_export import markdown_to_pptx
from .word_pdf import docx_to_pdf

PathLike = Union[str, Path]
_MARKDOWN = {".md", ".markdown"}
_LATEX = {".tex", ".latex"}


def convert(
    source: PathLike,
    output: Optional[PathLike] = None,
    *,
    template: str = "tdj",
    authors_json: Optional[PathLike] = None,
    author_ids: Optional[Sequence[str]] = None,
    checklist: bool = False,
    pdf: bool = False,
    pptx: bool = False,
    pptx_template: Optional[PathLike] = None,
    title_layout: Optional[str] = None,
    content_layout: Optional[str] = None,
) -> Path:
    """Собрать DOCX из .tex или .md.

    Расширение выбирает конвертер. Для Markdown аргументы шаблона и авторов
    не используются. Чеклист TD-J по умолчанию выключен: библиотека не пишет
    в stdout. Верните checklist=True, если нужен тот же отчёт, что у CLI.
    При pdf=True Word сохраняет PDF рядом с DOCX (тот же stem, суффикс .pdf).
    При pptx=True или если output оканчивается на .pptx Markdown собирается
    в презентацию. pptx_template — файл .pptx организации.
    """
    path = Path(source)
    suffix = path.suffix.lower()
    wants_pptx = pptx or (output is not None and Path(output).suffix.lower() == ".pptx")
    if wants_pptx and pdf:
        raise ValueError("PDF и PPTX одним вызовом не собираются")
    if wants_pptx:
        if suffix not in _MARKDOWN:
            raise ValueError("PPTX собирается из Markdown (.md)")
        return markdown_to_pptx(
            path,
            output,
            template=pptx_template,
            title_layout=title_layout,
            content_layout=content_layout,
        )
    if suffix in _MARKDOWN:
        out = markdown_to_docx(path, output)
    elif suffix in _LATEX:
        out = latex_to_docx(
            path,
            output,
            template=template,
            authors_json=authors_json,
            author_ids=author_ids,
            checklist=checklist,
        )
    else:
        raise ValueError(
            f"Неизвестное расширение {suffix!r}: ожидается .tex, .latex, .md или .markdown"
        )
    if pdf:
        docx_to_pdf(out)
    return out


def diff(
    base: PathLike,
    edited: PathLike,
    output: Optional[PathLike] = None,
) -> Path:
    """Сравнить два DOCX и записать markdown-отчёт."""
    return write_docx_diff_report(base, edited, output)

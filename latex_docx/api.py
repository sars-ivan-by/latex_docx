"""Публичная обёртка: один вызов для .tex и .md."""
from __future__ import annotations

from pathlib import Path
from typing import Optional, Sequence, Union

from .convert import latex_to_docx, markdown_to_docx
from .diff import write_docx_diff_report

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
) -> Path:
    """Собрать DOCX из .tex или .md.

    Расширение выбирает конвертер. Для Markdown аргументы шаблона и авторов
    не используются. Чеклист TD-J по умолчанию выключен: библиотека не пишет
    в stdout. Верните checklist=True, если нужен тот же отчёт, что у CLI.
    """
    path = Path(source)
    suffix = path.suffix.lower()
    if suffix in _MARKDOWN:
        return markdown_to_docx(path, output)
    if suffix in _LATEX:
        return latex_to_docx(
            path,
            output,
            template=template,
            authors_json=authors_json,
            author_ids=author_ids,
            checklist=checklist,
        )
    raise ValueError(
        f"Неизвестное расширение {suffix!r}: ожидается .tex, .latex, .md или .markdown"
    )


def diff(
    base: PathLike,
    edited: PathLike,
    output: Optional[PathLike] = None,
) -> Path:
    """Сравнить два DOCX и записать markdown-отчёт."""
    return write_docx_diff_report(base, edited, output)

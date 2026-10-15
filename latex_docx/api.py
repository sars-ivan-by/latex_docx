"""Public wrapper: one call for .tex and .md."""
from __future__ import annotations

from pathlib import Path
from typing import Optional, Sequence, Union

from .convert import latex_to_docx, markdown_to_docx
from .i18n import normalize_lang
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
    template: str = "note",
    authors_json: Optional[PathLike] = None,
    author_ids: Optional[Sequence[str]] = None,
    checklist: bool = False,
    pdf: bool = False,
    pptx: bool = False,
    pptx_template: Optional[PathLike] = None,
    title_layout: Optional[str] = None,
    content_layout: Optional[str] = None,
    lang: str = "ru",
    docx_template: Optional[PathLike] = None,
) -> Path:
    """Build a DOCX from .tex or .md.

    The suffix picks the converter. template is note, letter, memo, report,
    article, or tdj. Markdown uses it for the page setup and ignores author
    arguments. The journal checklist stays quiet unless checklist=True and
    template is tdj.
    pdf=True asks Microsoft Word to write a PDF next to the DOCX.
    pptx=True, or an output path ending in .pptx, builds a presentation
    from Markdown. pptx_template is an organization .pptx file.
    lang is "ru" or "en". It sets shared labels (figure, table, references)
    and which front-matter block template="plain" or template="article"
    prints. template="tdj" still prints the Russian block, then the English
    block. template="article" is the one-block layout for a paper that is
    not a TD-J submission. docx_template is an organization .docx or .dotx:
    its styles, page setup, headers, and footers are kept, and sample body
    text is dropped.
    """
    lang = normalize_lang(lang)
    path = Path(source)
    suffix = path.suffix.lower()
    wants_pptx = pptx or (output is not None and Path(output).suffix.lower() == ".pptx")
    if wants_pptx and pdf:
        raise ValueError("PDF and PPTX cannot be built in one call")
    if wants_pptx:
        if suffix not in _MARKDOWN:
            raise ValueError("PPTX is built from Markdown (.md)")
        return markdown_to_pptx(
            path,
            output,
            template=pptx_template,
            title_layout=title_layout,
            content_layout=content_layout,
            lang=lang,
        )
    if suffix in _MARKDOWN:
        out = markdown_to_docx(path, output, lang=lang, docx_template=docx_template, template=template)
    elif suffix in _LATEX:
        out = latex_to_docx(
            path,
            output,
            template=template,
            authors_json=authors_json,
            author_ids=author_ids,
            checklist=checklist,
            lang=lang,
            docx_template=docx_template,
        )
    else:
        raise ValueError(
            f"Unknown extension {suffix!r}: expected .tex, .latex, .md, or .markdown"
        )
    if pdf:
        docx_to_pdf(out)
    return out


def diff(
    base: PathLike,
    edited: PathLike,
    output: Optional[PathLike] = None,
    *,
    lang: str = "en",
) -> Path:
    """Compare two DOCX files and write a Markdown report.

    The report stays English unless lang="ru".
    """
    return write_docx_diff_report(base, edited, output, lang=lang)

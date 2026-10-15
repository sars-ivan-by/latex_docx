"""Public wrapper: one call for .tex and .md."""
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
    """Build a DOCX from .tex or .md.

    The suffix picks the converter. Markdown ignores template and author
    arguments. The TD-J checklist stays quiet unless checklist=True.
    pdf=True asks Microsoft Word to write a PDF next to the DOCX.
    pptx=True, or an output path ending in .pptx, builds a presentation
    from Markdown. pptx_template is an organization .pptx file.
    """
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
            f"Unknown extension {suffix!r}: expected .tex, .latex, .md, or .markdown"
        )
    if pdf:
        docx_to_pdf(out)
    return out


def diff(
    base: PathLike,
    edited: PathLike,
    output: Optional[PathLike] = None,
) -> Path:
    """Compare two DOCX files and write a Markdown report."""
    return write_docx_diff_report(base, edited, output)

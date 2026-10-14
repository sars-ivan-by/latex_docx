"""LaTeX / Markdown → DOCX converter package."""
from .api import convert, diff
from .convert import latex_to_docx, markdown_to_docx
from .diff import diff_docx_paragraphs, write_docx_diff_report
from .md_parser import MarkdownParser
from .parser import LatexParser

__all__ = [
    "convert",
    "diff",
    "latex_to_docx",
    "markdown_to_docx",
    "LatexParser",
    "MarkdownParser",
    "diff_docx_paragraphs",
    "write_docx_diff_report",
]

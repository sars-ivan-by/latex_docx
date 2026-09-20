"""LaTeX → DOCX converter package."""
from .convert import latex_to_docx
from .diff import diff_docx_paragraphs, write_docx_diff_report
from .parser import LatexParser

__all__ = [
    "latex_to_docx",
    "LatexParser",
    "diff_docx_paragraphs",
    "write_docx_diff_report",
]

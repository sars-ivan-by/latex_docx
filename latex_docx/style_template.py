"""Open an organization .docx or .dotx and drop its sample body."""
from __future__ import annotations

import io
import zipfile
from pathlib import Path

import docx
from docx.oxml.ns import qn

_TEMPLATE_MAIN = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
)
_DOCUMENT_MAIN = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)


def load_style_document(path: Path):
    """Return a Document that keeps the file's styles, sections, and theme.

    Paragraphs and tables already in the file are removed. Headers, footers,
    and sectPr stay.
    """
    path = Path(path).resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    suffix = path.suffix.lower()
    if suffix == ".docx":
        document = docx.Document(str(path))
    elif suffix == ".dotx":
        document = docx.Document(_dotx_stream(path))
    else:
        raise ValueError(
            f"Style template must be .docx or .dotx, got {path.suffix!r}"
        )
    _clear_body(document)
    return document


def _dotx_stream(path: Path) -> io.BytesIO:
    """python-docx opens only the document content type, so retag a .dotx."""
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(buf, "w") as dest:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(_TEMPLATE_MAIN, _DOCUMENT_MAIN)
            dest.writestr(info, data)
    buf.seek(0)
    return buf


def _clear_body(document) -> None:
    body = document.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

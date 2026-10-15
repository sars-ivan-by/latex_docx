"""DOCX → PDF через Word: пути и ошибки без запуска Word."""
from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from latex_docx.word_pdf import PdfExportError, default_pdf_path, docx_to_pdf


def test_default_pdf_path_keeps_stem():
    assert default_pdf_path(Path("out_note.docx")).name == "out_note.pdf"


def test_missing_docx():
    with pytest.raises(FileNotFoundError):
        docx_to_pdf("/no/such/out_missing.docx")


def test_rejects_non_docx(tmp_path: Path):
    src = tmp_path / "note.md"
    src.write_text("# x\n", encoding="utf-8")
    with pytest.raises(PdfExportError):
        docx_to_pdf(src)


def test_export_writes_named_pdf(tmp_path: Path, monkeypatch):
    src = tmp_path / "note.docx"
    Document().save(src)

    def fake(docx: Path, pdf: Path, timeout: float) -> None:
        assert docx == src.resolve()
        pdf.write_bytes(b"%PDF-1.4\n")

    monkeypatch.setattr("latex_docx.word_pdf.sys.platform", "darwin")
    monkeypatch.setattr("latex_docx.word_pdf._export_macos", fake)
    out = docx_to_pdf(src, tmp_path / "report.pdf")
    assert out.name == "out_report.pdf"
    assert out.read_bytes().startswith(b"%PDF")

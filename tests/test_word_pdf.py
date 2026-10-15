"""DOCX → PDF via Word: paths and errors without launching Word."""
from __future__ import annotations

import subprocess
import sys
import types
from pathlib import Path

import pytest
from docx import Document

from latex_docx.word_pdf import PdfExportError, _export_windows, default_pdf_path, docx_to_pdf


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


def _note(tmp_path: Path) -> Path:
    src = tmp_path / "note.docx"
    Document().save(src)
    return src


def test_windows_export_uses_a_private_word_instance(tmp_path: Path, monkeypatch):
    src = _note(tmp_path)
    dest = tmp_path / "out_note.pdf"
    calls = {}

    class Doc:
        def ExportAsFixedFormat(self, path, fmt):
            calls["export"] = (path, fmt)
            Path(path).write_bytes(b"%PDF-1.4\n")

        def Close(self, save):
            calls["close"] = save

    class Documents:
        def Open(self, path, confirm, readonly, recent):
            calls["open"] = (path, confirm, readonly, recent)
            return Doc()

    class Word:
        def __init__(self):
            self.Documents = Documents()
            self.Visible = True
            self.DisplayAlerts = 1

        def Quit(self):
            calls["quit"] = True

    client = types.ModuleType("win32com.client")

    def dispatch_ex(name):
        calls["app"] = name
        return Word()

    client.DispatchEx = dispatch_ex
    package = types.ModuleType("win32com")
    package.client = client
    monkeypatch.setitem(sys.modules, "win32com", package)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(
        "latex_docx.word_pdf._word_exe_windows",
        lambda: Path("C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE"),
    )

    _export_windows(src.resolve(), dest)
    assert calls["app"] == "Word.Application"
    assert calls["open"][1:] == (False, True, False)
    assert calls["export"][1] == 17
    assert calls["close"] == 0
    assert calls["quit"] is True
    assert dest.read_bytes().startswith(b"%PDF")


def test_windows_missing_word(tmp_path: Path, monkeypatch):
    monkeypatch.setattr("latex_docx.word_pdf._word_exe_windows", lambda: None)
    with pytest.raises(PdfExportError, match="WINWORD"):
        _export_windows(_note(tmp_path), tmp_path / "out_note.pdf")


def test_windows_missing_pywin32(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(
        "latex_docx.word_pdf._word_exe_windows",
        lambda: Path("C:/Word/WINWORD.EXE"),
    )
    monkeypatch.delitem(sys.modules, "win32com.client", raising=False)
    monkeypatch.delitem(sys.modules, "win32com", raising=False)
    real_import = __import__

    def blocked(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "win32com" or name.startswith("win32com."):
            raise ImportError("blocked")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr("builtins.__import__", blocked)
    with pytest.raises(PdfExportError, match="latex-docx\\[win\\]"):
        _export_windows(_note(tmp_path), tmp_path / "out_note.pdf")


def test_linux_export_uses_libreoffice(tmp_path: Path, monkeypatch):
    src = _note(tmp_path)
    seen = {}

    def fake_run(cmd, **kwargs):
        seen["cmd"] = cmd
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        (outdir / f"{src.stem}.pdf").write_bytes(b"%PDF-1.4\n")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr("latex_docx.word_pdf.sys.platform", "linux")
    monkeypatch.setattr("latex_docx.word_pdf._soffice_bin", lambda: "/usr/bin/soffice")
    monkeypatch.setattr("latex_docx.word_pdf.subprocess.run", fake_run)
    out = docx_to_pdf(src, tmp_path / "report.pdf")
    assert seen["cmd"][0] == "/usr/bin/soffice"
    assert "pdf:writer_pdf_Export" in seen["cmd"]
    assert out.read_bytes().startswith(b"%PDF")


def test_linux_missing_libreoffice(tmp_path: Path, monkeypatch):
    monkeypatch.setattr("latex_docx.word_pdf.sys.platform", "linux")
    monkeypatch.setattr("latex_docx.word_pdf._soffice_bin", lambda: None)
    with pytest.raises(PdfExportError, match="LibreOffice"):
        docx_to_pdf(_note(tmp_path))

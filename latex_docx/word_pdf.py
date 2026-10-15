"""DOCX → PDF.

macOS drives the open Word app with AppleScript and does not quit it.
Windows starts its own Word instance through pywin32 and closes that instance.
Linux has no Word, so it asks LibreOffice (`soffice`) to write the PDF.
That page can differ from Word.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Union

from .util import ensure_out_prefix

PathLike = Union[str, Path]
_SCRIPT = Path(__file__).with_name("word_to_pdf.applescript")
_MAC_APPS = (
    Path("/Applications/Microsoft Word.app"),
    Path.home() / "Applications" / "Microsoft Word.app",
)
# Word VBA: wdExportFormatPDF, wdAlertsNone, wdDoNotSaveChanges.
_WD_EXPORT_PDF = 17
_WD_ALERTS_NONE = 0
_WD_DO_NOT_SAVE = 0
_WINWORD_KEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE"


class PdfExportError(RuntimeError):
    """Word could not write the PDF."""


def word_available() -> bool:
    if sys.platform == "darwin":
        return any(p.is_dir() for p in _MAC_APPS)
    if sys.platform == "win32":
        return _word_exe_windows() is not None
    return False


def _word_exe_windows() -> Optional[Path]:
    """WINWORD.EXE from App Paths, then the usual Office folders."""
    import winreg

    hives = (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER)
    flags = [winreg.KEY_READ]
    for extra in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        value = getattr(winreg, extra, None)
        if value is not None:
            flags.append(winreg.KEY_READ | value)
    for hive in hives:
        for flag in flags:
            try:
                with winreg.OpenKey(hive, _WINWORD_KEY, 0, flag) as key:
                    raw, _ = winreg.QueryValueEx(key, "")
            except OSError:
                continue
            if raw and Path(raw).is_file():
                return Path(raw)
    roots = [
        os.environ.get("ProgramW6432"),
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
    ]
    for root in roots:
        if not root:
            continue
        for office in ("Office16", "Office15"):
            candidate = Path(root) / "Microsoft Office" / "root" / office / "WINWORD.EXE"
            if candidate.is_file():
                return candidate
    return None


def default_pdf_path(docx_path: PathLike) -> Path:
    return Path(docx_path).with_suffix(".pdf")


def docx_to_pdf(
    docx_path: PathLike,
    pdf_path: Optional[PathLike] = None,
    *,
    timeout: float = 180,
) -> Path:
    """Save a PDF beside the DOCX, or at pdf_path.

    macOS and Windows use Microsoft Word. Linux uses LibreOffice.
    """
    src = Path(docx_path).resolve()
    if not src.exists():
        raise FileNotFoundError(src)
    if src.suffix.lower() != ".docx":
        raise PdfExportError(f"Expected a .docx file, got {src.name}")
    dest = ensure_out_prefix(pdf_path) if pdf_path else default_pdf_path(src)
    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()

    if sys.platform == "darwin":
        _export_macos(src, dest, timeout)
    elif sys.platform == "win32":
        _export_windows(src, dest)
    elif sys.platform.startswith("linux"):
        _export_linux(src, dest, timeout)
    else:
        raise PdfExportError(
            "PDF export supports macOS (Word), Windows (Word), and Linux (LibreOffice)."
        )
    if not dest.exists() or dest.stat().st_size == 0:
        raise PdfExportError(f"The PDF was not written: {dest}")
    return dest


def _export_macos(src: Path, dest: Path, timeout: float) -> None:
    if not word_available():
        raise PdfExportError(
            "Microsoft Word was not found in /Applications. "
            "The PDF is written by the same Word that opens the DOCX."
        )
    if not _SCRIPT.is_file():
        raise PdfExportError(f"Missing Word script: {_SCRIPT.name}")
    cmd = ["/usr/bin/osascript", str(_SCRIPT), str(src), str(dest)]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise PdfExportError(
            f"Word did not save the PDF within {int(timeout)}s. Close Word dialogs and retry."
        ) from exc
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        raise PdfExportError(detail or "osascript failed")


def _export_windows(src: Path, dest: Path) -> None:
    if _word_exe_windows() is None:
        raise PdfExportError(
            "Microsoft Word (WINWORD.EXE) was not found. "
            "The PDF is written by Word so it matches the DOCX."
        )
    try:
        import win32com.client
    except ImportError as exc:
        raise PdfExportError(
            'Windows PDF export needs pywin32: pip install "latex-docx[win]".'
        ) from exc
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = _WD_ALERTS_NONE
        # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False.
        doc = word.Documents.Open(str(src), False, True, False)
        doc.ExportAsFixedFormat(str(dest), _WD_EXPORT_PDF)
    except PdfExportError:
        raise
    except Exception as exc:
        raise PdfExportError(f"Word could not save the PDF: {exc}") from exc
    finally:
        if doc is not None:
            try:
                doc.Close(_WD_DO_NOT_SAVE)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def _soffice_bin() -> Optional[str]:
    return shutil.which("soffice") or shutil.which("libreoffice")


def _export_linux(src: Path, dest: Path, timeout: float) -> None:
    binary = _soffice_bin()
    if not binary:
        raise PdfExportError(
            "Linux PDF export needs LibreOffice. Install it and make sure "
            "`soffice` is on PATH. Microsoft Word is not used on Linux, "
            "so the page can differ from Word on macOS and Windows."
        )
    with tempfile.TemporaryDirectory(prefix="latex-docx-pdf-") as tmp:
        root = Path(tmp)
        profile = root / "profile"
        outdir = root / "out"
        profile.mkdir()
        outdir.mkdir()
        cmd = [
            binary,
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(outdir),
            str(src),
        ]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise PdfExportError(
                f"LibreOffice did not save the PDF within {int(timeout)}s."
            ) from exc
        produced = outdir / f"{src.stem}.pdf"
        if proc.returncode != 0 or not produced.is_file():
            detail = (proc.stderr or proc.stdout or "").strip()
            raise PdfExportError(detail or "LibreOffice did not write a PDF")
        shutil.move(str(produced), dest)

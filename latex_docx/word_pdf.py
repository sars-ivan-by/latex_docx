"""DOCX → PDF через установленный Microsoft Word.

Вёрстка берётся из Word, отдельный движок PDF не используется.
На macOS документ открывает AppleScript, на Windows — Word.Application.
Word после экспорта не закрываем: закрывается только этот документ.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from typing import Optional, Union

from .util import ensure_out_prefix

PathLike = Union[str, Path]
_SCRIPT = Path(__file__).with_name("word_to_pdf.applescript")
_MAC_APPS = (
    Path("/Applications/Microsoft Word.app"),
    Path.home() / "Applications" / "Microsoft Word.app",
)


class PdfExportError(RuntimeError):
    """Word не смог записать PDF."""


def word_available() -> bool:
    if sys.platform == "darwin":
        return any(p.is_dir() for p in _MAC_APPS)
    if sys.platform == "win32":
        return True
    return False


def default_pdf_path(docx_path: PathLike) -> Path:
    return Path(docx_path).with_suffix(".pdf")


def docx_to_pdf(
    docx_path: PathLike,
    pdf_path: Optional[PathLike] = None,
    *,
    timeout: float = 180,
) -> Path:
    """Открыть DOCX в Word и сохранить PDF рядом или по явному пути."""
    src = Path(docx_path).resolve()
    if not src.exists():
        raise FileNotFoundError(src)
    if src.suffix.lower() != ".docx":
        raise PdfExportError(f"Ожидается .docx, получен {src.name}")
    dest = ensure_out_prefix(pdf_path) if pdf_path else default_pdf_path(src)
    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()

    if sys.platform == "darwin":
        _export_macos(src, dest, timeout)
    elif sys.platform == "win32":
        _export_windows(src, dest)
    else:
        raise PdfExportError(
            "PDF собирается через Microsoft Word, чтобы совпасть с DOCX. "
            "На этой системе Word недоступен."
        )
    if not dest.exists() or dest.stat().st_size == 0:
        raise PdfExportError(f"Word не записал PDF: {dest}")
    return dest


def _export_macos(src: Path, dest: Path, timeout: float) -> None:
    if not word_available():
        raise PdfExportError(
            "Microsoft Word не найден в /Applications. "
            "PDF пишется тем же Word, что открывает DOCX."
        )
    if not _SCRIPT.is_file():
        raise PdfExportError(f"Нет скрипта Word: {_SCRIPT.name}")
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
            f"Word не сохранил PDF за {int(timeout)} с. Закройте диалоги Word и повторите."
        ) from exc
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        raise PdfExportError(detail or "osascript завершился с ошибкой")


def _export_windows(src: Path, dest: Path) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise PdfExportError(
            "На Windows нужен pywin32: pip install pywin32. "
            "PDF сохраняет установленный Microsoft Word."
        ) from exc
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(str(src))
    try:
        doc.SaveAs(str(dest), FileFormat=17)
    finally:
        doc.Close(0)

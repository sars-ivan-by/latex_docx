"""CLI entry point."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Sequence

from .constants import DEFAULT_OP_LOG, TEMPLATES, _HAS_OMML, resolve_authors_json
from .convert import latex_to_docx, markdown_to_docx
from .diff import write_docx_diff_report
from .util import append_op_log, resolve_by
from .word_pdf import docx_to_pdf


def main(argv: Optional[Sequence[str]] = None):
    parser = argparse.ArgumentParser(
        description="Convert LaTeX or Markdown to DOCX (TD-J layout for .tex, light layout for .md)"
    )
    parser.add_argument("tex", nargs="?", help="Input .tex, .md, .markdown or .docx")
    parser.add_argument(
        "-o",
        "--output",
        help="Output path (имя получит приставку out_, если её нет)",
    )
    parser.add_argument(
        "--template",
        choices=TEMPLATES,
        default="tdj",
        help="Front-matter layout: tdj (bilingual) or plain (RU-only)",
    )
    parser.add_argument(
        "--authors-json",
        default=None,
        help="Path to authors.json (default: cwd / repo / authors.example.json)",
    )
    parser.add_argument(
        "--authors",
        help="Comma-separated author ids from authors.json (auto-detect if omitted)",
    )
    parser.add_argument(
        "--diff-docx",
        nargs=2,
        metavar=("BASE.docx", "EDITED.docx"),
        help="Compare two docx files and write a markdown change report",
    )
    parser.add_argument(
        "--by",
        choices=("agent", "user"),
        default=None,
        help="Who ran the conversion (written to log). Default: LATEX_DOCX_BY or auto-detect",
    )
    parser.add_argument(
        "--log",
        default=str(DEFAULT_OP_LOG),
        help=f"Append-only operations log (default: {DEFAULT_OP_LOG.name})",
    )
    parser.add_argument(
        "--no-log",
        action="store_true",
        help="Do not append to the operations log",
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="After DOCX, save PDF with Microsoft Word (same layout)",
    )
    args = parser.parse_args(argv)
    src_suffix = Path(args.tex).suffix.lower() if args.tex else ""
    if not _HAS_OMML and not args.diff_docx and src_suffix != ".docx":
        print(
            "Warning: latex2mathml/mathml2omml not installed — formulas will be plain text.\n"
            "  pip install 'latex-docx[omml]'  or  pip install latex2mathml mathml2omml",
            file=sys.stderr,
        )

    who = resolve_by(args.by)

    if args.diff_docx:
        base, edited = args.diff_docx
        report = write_docx_diff_report(base, edited, args.output)
        print(f"Diff report: {report}")
        if not args.no_log:
            append_op_log(
                "diff",
                f"base={base} vs edited={edited} → changes={report}",
                by=who,
                log_path=args.log,
            )
            print(f"log → {args.log} (by={who})")
        return

    if not args.tex:
        parser.error("tex is required unless --diff-docx is used")

    src = Path(args.tex)
    if src.suffix.lower() == ".docx":
        pdf = docx_to_pdf(src, args.output)
        print(f"PDF: {pdf}")
        if not args.no_log:
            append_op_log(
                "pdf",
                f"docx={src} → pdf={pdf}",
                by=who,
                log_path=args.log,
            )
            print(f"log → {args.log} (by={who})")
        return

    if src.suffix.lower() in {".md", ".markdown"}:
        out = markdown_to_docx(args.tex, args.output)
        print(f"Saved: {out}")
        pdf = _write_pdf(out, args.pdf)
        if not args.no_log:
            size = out.stat().st_size if out.exists() else 0
            append_op_log(
                "convert-md",
                f"md={args.tex} → base={out} size={size}B pdf={pdf or '-'}",
                by=who,
                log_path=args.log,
            )
            print(f"log → {args.log} (by={who})")
        return

    ids = [x.strip() for x in args.authors.split(",")] if args.authors else None
    authors_path = resolve_authors_json(args.authors_json)
    out = latex_to_docx(
        args.tex,
        args.output,
        template=args.template,
        authors_json=authors_path,
        author_ids=ids,
    )
    print(f"Saved: {out}")
    pdf = _write_pdf(out, args.pdf)
    if not args.no_log:
        size = out.stat().st_size if out.exists() else 0
        append_op_log(
            "convert",
            f"tex={args.tex} → base={out} size={size}B authors={authors_path} pdf={pdf or '-'}",
            by=who,
            log_path=args.log,
        )
        print(f"log → {args.log} (by={who})")


def _write_pdf(docx_path: Path, enabled: bool) -> Optional[Path]:
    if not enabled:
        return None
    pdf = docx_to_pdf(docx_path)
    print(f"PDF: {pdf}")
    return pdf


if __name__ == "__main__":
    main()

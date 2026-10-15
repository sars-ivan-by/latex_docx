#!/usr/bin/env python3
"""Обёртка над установленным пакетом latex-docx.

Скопируйте в свой проект вместе с examples/Makefile.
Пакет должен быть уже в окружении:

    pip install "latex-docx[omml]"

Вызов:

    python wrap.py sample.md -o out_sample.docx
    python wrap.py article.tex -o out_article.docx --template plain --checklist
    python wrap.py --diff out_base.docx out_edited.docx -o out_changes.md
    python wrap.py sample.md -o out_sample.docx --pdf
    python wrap.py out_sample.docx -o out_sample.pdf
    python wrap.py slides.md -o out_slides.pptx --pptx-template org.pptx
"""
from __future__ import annotations

import argparse
import sys

from latex_docx import convert, diff, docx_to_pdf


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Library wrapper around latex-docx")
    parser.add_argument("source", nargs="?", help="Input .tex or .md")
    parser.add_argument("-o", "--output", help="Output path (получит приставку out_)")
    parser.add_argument("--template", default="tdj", choices=("tdj", "plain"))
    parser.add_argument("--authors-json", default=None)
    parser.add_argument("--authors", default=None, help="id через запятую")
    parser.add_argument(
        "--checklist",
        action="store_true",
        help="Печатать чеклист TD-J для .tex",
    )
    parser.add_argument(
        "--diff",
        nargs=2,
        metavar=("BASE.docx", "EDITED.docx"),
        help="Сравнить два DOCX и записать отчёт",
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="Сохранить PDF через Microsoft Word",
    )
    parser.add_argument("--pptx", action="store_true", help="Собрать PPTX из Markdown")
    parser.add_argument("--pptx-template", default=None, help="Шаблон организации .pptx")
    parser.add_argument("--title-layout", default=None)
    parser.add_argument("--content-layout", default=None)
    args = parser.parse_args(argv)

    if args.diff:
        report = diff(args.diff[0], args.diff[1], args.output)
        print(report)
        return

    if not args.source:
        parser.error("нужен исходник .tex/.md/.docx или флаг --diff")

    if args.source.lower().endswith(".docx"):
        print(docx_to_pdf(args.source, args.output))
        return

    ids = [x.strip() for x in args.authors.split(",")] if args.authors else None
    out = convert(
        args.source,
        args.output,
        template=args.template,
        authors_json=args.authors_json,
        author_ids=ids,
        checklist=args.checklist,
        pdf=args.pdf,
        pptx=args.pptx or bool(args.pptx_template or args.title_layout or args.content_layout),
        pptx_template=args.pptx_template,
        title_layout=args.title_layout,
        content_layout=args.content_layout,
    )
    print(out)
    if args.pdf:
        print(out.with_suffix(".pdf"))


if __name__ == "__main__":
    main(sys.argv[1:])

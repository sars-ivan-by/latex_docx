# latex-docx

Convert a typical scientific article from LaTeX to Microsoft Word, a light Markdown file to the same DOCX layout, or Markdown to PowerPoint. The Word file can be saved as PDF by Microsoft Word, so the PDF matches the DOCX.

The LaTeX path follows the [TD-J](https://td-j.ru/) journal: Times New Roman 14 pt, margins of about 2 cm, equation numbers on the right, OMML math. Labels inside that DOCX stay in Russian (`Аннотация`, `Рис.`, `Таблица`) because the journal asks for them. Tooling, errors, and this repository are in English.

This is not a pandoc replacement. The LaTeX parser covers a typical article: `equation` / `align`, `figure`, `tabular`, `thebibliography`, bilingual front matter.

Documentation: [English](https://sars-ivan-by.github.io/latex_docx/en/) · [Русский](https://sars-ivan-by.github.io/latex_docx/ru/).

## Install

```bash
pip install "latex-docx[omml,pptx]"
```

`omml` turns formulas into Word math. Without it, formulas become italic text. `pptx` is required for presentations. On Windows, PDF export also needs `pip install pywin32` and a local Word install.

From a git checkout:

```bash
git clone https://github.com/sars-ivan-by/latex_docx.git
cd latex_docx
python -m venv .venv
source .venv/bin/activate
pip install -e ".[omml,pptx,dev]"
```

## Quick start

```bash
latex-docx article.tex -o out_article.docx
latex-docx article.tex --template plain -o out_plain.docx
latex-docx article.tex --authors alice,bob -o out_article.docx

latex-docx notes.md -o out_notes.docx
latex-docx notes.md -o out_notes.docx --pdf
latex-docx out_notes.docx -o out_notes.pdf

latex-docx slides.md -o out_slides.pptx --pptx-template org.pptx
latex-docx slides.md --pptx --title-layout "Title Slide" --content-layout "Title and Content"

latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

The same commands work as `python -m latex_docx`. Output names gain an `out_` prefix when they lack one (see `.gitignore`).

```bash
make convert TEX=article.tex BASE=out_base.docx
make convert-md MD=notes.md BASE=out_notes.docx
make prepare-edit EDITED=out_edited.docx
make diff BASE=out_base.docx EDITED=out_edited.docx CHANGES=out_changes.md
make pdf
make pptx SLIDES=slides.md PPTX=out_slides.pptx PPTX_TEMPLATE=org.pptx
```

## Library

```python
from latex_docx import convert, diff, docx_to_pdf, markdown_to_pptx

convert("notes.md", "out_notes.docx")
convert("article.tex", "out_article.docx", template="tdj", checklist=True)
convert("notes.md", "out_notes.docx", pdf=True)
docx_to_pdf("out_notes.docx", "out_notes.pdf")
markdown_to_pptx("slides.md", "out_slides.pptx", template="org.pptx")
diff("out_base.docx", "out_edited.docx", "out_changes.md")
```

`convert` picks LaTeX or Markdown from the suffix. The TD-J checklist is silent unless `checklist=True`. Copy [`examples/wrap.py`](examples/wrap.py) and [`examples/Makefile`](examples/Makefile) into another project.

Agent notes for choosing a source format: [`SKILLS.md`](SKILLS.md).

## Authors

`--authors-json`, then `./authors.json`, then `authors.json` in the repo, then `authors.example.json`. Copy the example and fill `id`, names, `org_ids`, and `organizations`. `--authors id1,id2` always replaces the author block. Without that flag, a careful front matter already present in the `.tex` is left as it is.

## Markdown and slides

For DOCX, the first `#` heading becomes the centered title unless YAML `title:` is set. The parser accepts paragraphs, `**bold**`, `*italic*`, `` `code` ``, links, one-level lists, task items, GFM tables, images, fenced code, quotes, and `$...$` / `$$...$$`.

For PPTX, a line `---` starts the next slide. A single `#` is the title slide, `##` is a content slide, `-` is a bullet, and `???` starts speaker notes. See [`examples/slides.md`](examples/slides.md). `--pptx-template` keeps the file's size, theme, and layouts, and drops sample slides that were already in it.

## PDF

PDF is whatever Word saves from the DOCX. macOS uses AppleScript. Windows uses Word through `pywin32`. Word is not quit; only that document is closed. The first run may ask for permission to control Word.

## Tests

```bash
pytest
```

## Limits

The parser covers a typical article (`equation` / `align`, `figure` / `subfigure`, `tabular` / `tabularx` / `longtable`, `thebibliography`, and a sibling `.bbl`). Do not auto-import a DOCX back into TeX. Use `--diff-docx` and edit the source. Backlog: [`troubles.md`](troubles.md).

## GitHub Pages

English and Russian pages live in [`docs/`](docs/). The workflow [`.github/workflows/pages.yml`](.github/workflows/pages.yml) publishes them.

In the GitHub repository: **Settings → Pages → Build and deployment → Source: GitHub Actions**. Push to `main` deploys `https://sars-ivan-by.github.io/latex_docx/`.

## Publish to PyPI

The project name `latex-docx` is unused on PyPI as of this packaging. Version lives in `pyproject.toml`.

1. On [pypi.org](https://pypi.org/manage/account/publishing/), add a pending publisher: owner `sars-ivan-by`, repository `latex_docx`, workflow `pypi.yml`, environment `pypi`.
2. In the GitHub repository, create an environment named `pypi` (**Settings → Environments**).
3. Tag a release. [`.github/workflows/pypi.yml`](.github/workflows/pypi.yml) builds the sdist and wheel and uploads them with PyPI trusted publishing. No API token is stored in the repo.

Local check, without uploading:

```bash
pip install build twine
python -m build
twine check dist/*
```

## License

MIT. See [`LICENSE`](LICENSE).

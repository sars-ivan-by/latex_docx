# latex-docx

Turn a typical scientific article from LaTeX into Word, a Markdown note into the same DOCX, or Markdown into PowerPoint. PDF is that Word file saved by Microsoft Word, so the page matches the DOCX.

The LaTeX path follows the [TD-J](https://td-j.ru/) journal: Times New Roman 14 pt, margins of about 2 cm, equation numbers on the right, OMML math when the extra is installed.

- Documentation: [English](https://sars-ivan-by.github.io/latex_docx/en/) · [Русский](https://sars-ivan-by.github.io/latex_docx/ru/) · [中文](https://sars-ivan-by.github.io/latex_docx/zh/)
- Roadmap: [`ROADMAP.md`](ROADMAP.md) · [English](https://sars-ivan-by.github.io/latex_docx/en/roadmap.html) · [Русский](https://sars-ivan-by.github.io/latex_docx/ru/roadmap.html) · [中文](https://sars-ivan-by.github.io/latex_docx/zh/roadmap.html)
- Agent notes: [`SKILLS.md`](SKILLS.md)

## Install

```bash
pip install "latex-docx[omml,pptx]"
```

`omml` writes formulas as Word math. Without it they become italic text and the CLI warns. `pptx` is required for slides. PDF on macOS and Windows needs Microsoft Word. On Windows install the extra that pulls in `pywin32`:

```bash
pip install "latex-docx[omml,pptx,win]"
```

From a checkout:

```bash
git clone https://github.com/sars-ivan-by/latex_docx.git
cd latex_docx
python -m venv .venv
source .venv/bin/activate
pip install -e ".[omml,pptx,win,dev]"
pytest
```

## Quick start

```bash
latex-docx article.tex -o out_article.docx
latex-docx article.tex --template article --lang en -o out_article.docx
latex-docx notes.md -o out_notes.docx --lang en --docx-template org.docx
latex-docx notes.md -o out_notes.docx --pdf
latex-docx slides.md -o out_slides.pptx --pptx-template org.pptx
latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

`python -m latex_docx` is the same program. A file name without the `out_` prefix receives one (see `.gitignore`).

```bash
make convert TEX=article.tex BASE=out_base.docx
make pdf
make pptx SLIDES=slides.md PPTX=out_slides.pptx PPTX_TEMPLATE=org.pptx
```

Copy [`examples/wrap.py`](examples/wrap.py) and [`examples/Makefile`](examples/Makefile) into another project.

## Language

`lang` is `ru` or `en`. The default is `ru`.

| | `tdj` (default for `.tex`) | `plain` (Markdown, or `--template plain`) |
|---|---|---|
| Front matter | Russian block, then English block. Each block keeps its own labels (`Аннотация` / `Abstract`). | One block. `en` prefers `*_en` fields and falls back to `*_ru`, so a Markdown title still appears. |
| Shared labels | Follow `lang`: figure, table, references, author block, UDC, missing image. | Same. |

```bash
latex-docx article.tex --lang en -o out_article.docx
latex-docx notes.md --lang en -o out_notes.docx
```

```python
from latex_docx import convert, diff

convert("notes.md", "out_notes.docx", lang="en")
convert("article.tex", "out_article.docx", template="tdj", lang="ru", checklist=True)
diff("out_base.docx", "out_edited.docx", "out_changes.md", lang="ru")
```

The TD-J checklist is printed for a `.tex` file on the command line and stays quiet in `convert()` unless `checklist=True`. Its wording follows `lang`. `diff()` writes an English report unless `lang="ru"`. The CLI passes `--lang` through, so a CLI diff is Russian unless you pass `--lang en`.

## Library

```python
from latex_docx import convert, diff, docx_to_pdf, markdown_to_pptx

convert("article.tex", "out_article.docx", template="tdj", authors_json="authors.json")
convert("notes.md", "out_notes.docx", pdf=True)
docx_to_pdf("out_notes.docx", "out_notes.pdf")
markdown_to_pptx("slides.md", "out_slides.pptx", template="org.pptx")
```

`convert` chooses LaTeX or Markdown from the suffix. PDF and PPTX cannot be requested in one call.

## Styles

`--docx-template org.docx` (or a `.dotx`) keeps the organization's styles, page size, headers, and footers. Text that was already in that file is removed. Section headings use the heading styles from the file (`Heading 1`, `Заголовок 1`, and the same pattern for levels 2–9). Without a template, headings are still real Word heading styles, drawn in Times New Roman 14 pt.

`--template article` prints one front-matter block, the same way `plain` does, and does not run the TD-J checklist. Use it with `--lang en` for an English article.

## Authors, Markdown, slides, PDF

`--authors-json`, then `./authors.json`, then the copy in the repo, then `authors.example.json`. `--authors id1,id2` replaces the author block. Without that flag, a careful front matter already in the `.tex` is left as it is.

Markdown DOCX: the first `#` is the centered title unless YAML `title:` is set. Headings, bold, italic, code, links, nested lists, task items, footnotes (`[^id]` plus `[^id]: text`), GFM tables, images, fenced code, quotes, and `$...$` / `$$...$$` are accepted.

Slides: `---` starts the next slide, `#` is the title slide, `##` is content, `-` is a bullet, `???` is speaker notes. `--pptx-template` keeps size, theme, and layouts and drops slides that were already in the file. See [`examples/slides.md`](examples/slides.md).

PDF on macOS is whatever Word saves via AppleScript; Word is not quit, and the first run may ask for permission to control it. Windows uses a separate Word instance through `pywin32` and closes that instance when the PDF is written. Linux has no Word: `soffice` (LibreOffice) writes the PDF, and the page can differ from Word. Install LibreOffice and keep `soffice` on `PATH`.

## Limits

The parser covers a typical article: `equation` / `align`, `figure` / `subfigure`, `tabular` / `tabularx` / `longtable`, `thebibliography`, and a sibling `.bbl`. Do not import a DOCX back into TeX. Use `--diff-docx` and edit the source. Known limits: [`troubles.md`](troubles.md). What comes next: [`ROADMAP.md`](ROADMAP.md).

PyPI package name: `latex-docx`. A GitHub Release on `main` publishes through trusted publishing (`.github/workflows/pypi.yml`, environment `pypi`). Pages are the static files in `docs/`, deployed by `.github/workflows/pages.yml`.

## License

MIT. See [`LICENSE`](LICENSE).

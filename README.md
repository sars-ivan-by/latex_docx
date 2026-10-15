# latex-docx

Build a DOCX, a PDF, or a PowerPoint file from a source an agent or a script can write. The source is Markdown, or LaTeX when the document needs formulas and a bibliography. The program fills a document template. It does not replace the editor: Word and PowerPoint stay the place where a person makes the last pass.

An agent controls the result from the source and a few flags: which template, which language, which organization file, and, for slides, which layout. Colors and fonts of a branded file stay in that file.

A scientific article is one option (`--template article`, or `--template tdj` for a bilingual journal header). The default is an ordinary note.

- Agent instructions: [`SKILLS.md`](SKILLS.md)
- Documentation: [English](https://sars-ivan-by.github.io/latex_docx/en/) · [Русский](https://sars-ivan-by.github.io/latex_docx/ru/) · [中文](https://sars-ivan-by.github.io/latex_docx/zh/)
- Roadmap: [`ROADMAP.md`](ROADMAP.md)

## Install

```bash
pip install "latex-docx[omml,pptx]"
```

`omml` writes formulas as Word math. Without it they become italic text and the CLI warns. `pptx` is required for slides. PDF on macOS needs Microsoft Word. On Linux, LibreOffice (`soffice`) writes the PDF, and the page can differ from Word.

### Windows

Install Python 3.9 or newer with “Add python.exe to PATH”. PDF also needs Microsoft Word. In PowerShell:

```powershell
py -m pip install "latex-docx[omml,pptx,win]"
```

The `win` extra installs `pywin32`. In `cmd`, use `python` instead of `py` when Python is already on `PATH`.

From a checkout:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -e ".[omml,pptx,win,dev]"
py -m pytest
```

If PowerShell refuses `Activate.ps1`, run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once. In `cmd` the activator is `.venv\Scripts\activate.bat`.

### macOS and Linux checkout

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
latex-docx notes.md -o out_notes.docx
latex-docx letter.md --template letter --lang en -o out_letter.docx
latex-docx report.md --template report -o out_report.docx
latex-docx slides.md -o out_slides.pptx
latex-docx notes.md -o out_notes.docx --pdf
latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

`python -m latex_docx` is the same program. A file name without the `out_` prefix receives one (see `.gitignore`). Copy [`examples/wrap.py`](examples/wrap.py) into another project and call `convert()`.

```python
from latex_docx import convert, markdown_to_pptx

convert("notes.md", "out_notes.docx", template="note", lang="en")
convert("letter.md", "out_letter.docx", template="letter")
convert("slides.md", "out_slides.pptx")
```

`convert` chooses Markdown or LaTeX from the suffix. PDF and PPTX cannot be requested in one call. `diff()` writes an English report unless `lang="ru"`. The CLI passes `--lang` through, so a CLI diff is Russian unless you pass `--lang en`.

## What an agent sets

| Choice | Where |
|---|---|
| Document kind | `--template note\|letter\|memo\|report\|article\|tdj` |
| Language of labels | `--lang ru\|en` |
| Organization Word file | `--docx-template org.docx` (styles and page; sample body is dropped) |
| Slide deck | `--pptx-template org.pptx`, or no file and the built-in layouts |
| One slide's layout | `::layout Exact Name` after `latex-docx --pptx-layouts` |
| Look, when there is no deck file | `::font`, `::color RRGGBB`, `::background RRGGBB`, `::size widescreen` |

With an organization `.pptx`, those four `::` lines are ignored. The file already has the look. Full rules for agents: [`SKILLS.md`](SKILLS.md).

## Templates

| Template | Page |
|---|---|
| `note` | Default. Calibri 11 pt, left aligned, about 2.5 cm margins. |
| `letter` | Calibri 12 pt, wider margins, no UDC line and no abstract label. |
| `memo` | Same idea as a letter, with tighter margins. |
| `report` | Times New Roman 12 pt, justified, a small first-line indent. |
| `article` | One front-matter block, Times New Roman 14 pt, justified. The scientific option. |
| `tdj` | Optional bilingual header (Russian block, then English block) and a journal checklist on the command line. |
| `plain` | Same page as `note`. Kept for older commands. |

`--docx-template` replaces the page setup with the organization's Word styles. Headings are real Word heading styles either way, so a table of contents can see them.

## Markdown, slides, PDF

Markdown DOCX: the first `#` is the title unless YAML `title:` is set. Headings, bold, italic, code, links, nested lists, task items, footnotes (`[^id]` and `[^id]: text`), GFM tables, images, fenced code, quotes, and `$...$` / `$$...$$` are accepted.

Slides: `---` starts the next slide, `#` is the title slide, `##` is content, `-` is a bullet, `???` is speaker notes. See [`examples/slides.md`](examples/slides.md).

```bash
latex-docx --pptx-layouts
latex-docx slides.md -o out_slides.pptx --pptx-template org.pptx
```

PDF on macOS is whatever Word saves via AppleScript; Word is not quit. Windows uses a separate Word instance through `pywin32` and closes that instance. Linux has no Word: `soffice` (LibreOffice) writes the PDF, and the page can differ from Word.

## Authors and a journal article

`--authors-json`, then `./authors.json`, then the copy in the repo, then `authors.example.json`. `--authors id1,id2` replaces the author block. Without that flag, a careful front matter already in the `.tex` is left as it is.

```bash
latex-docx article.tex --template article --lang en -o out_article.docx
latex-docx article.tex --template tdj -o out_article.docx
```

The checklist runs for `tdj` on the command line. In `convert()` it stays quiet unless `checklist=True`.

## Limits

LaTeX covers a typical article: `equation` / `align`, `figure` / `subfigure`, `tabular` / `tabularx` / `longtable`, `thebibliography`, and a sibling `.bbl`. Do not import a DOCX back into the source. Use `--diff-docx` and edit the Markdown or TeX. Known limits: [`troubles.md`](troubles.md). What comes next: [`ROADMAP.md`](ROADMAP.md).

PyPI package name: `latex-docx`. A GitHub Release publishes through trusted publishing (`.github/workflows/pypi.yml`, environment `pypi`). Pages are the static files in `docs/`.

## License

MIT. See [`LICENSE`](LICENSE).

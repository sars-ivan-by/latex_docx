# Preparing documents with latex-docx

Instructions for an agent. You write the source and choose the template. The program builds the DOCX, PDF, or PPTX. A person may still open that file in Word or PowerPoint.

Everyday documents are Markdown: `--template note` (the default), `letter`, `memo`, or `report`. A talk is Markdown built to PPTX. A scientific article is optional: `.tex` with `--template article`, or `--template tdj` when the file needs a bilingual journal header. Pass `lang="ru"` (the default) or `lang="en"`. Shared labels (figure, table, references) follow `lang`.

## Which source

`note`, `letter`, `memo`, and `report` are Markdown or LaTeX with `--template`. `letter` and `memo` do not print a UDC line or an abstract label. `report` is justified Times 12 pt. `note` is Calibri 11 pt, left aligned.

`article` is one front-matter block in Times 14 pt. `tdj` prints the Russian block and then the English block, and the command line prints the journal checklist. `plain` is the same page as `note`.

Pass `--docx-template org.docx` (or `.dotx`) when the organization already has a Word file. Styles, page setup, headers, and footers come from that file. Sample paragraphs in it are dropped. Headings become Word heading styles either way.

Markdown lists may be indented. A footnote is `[^id]` in the text and `[^id]: the note` on its own line.

A presentation: the same Markdown, with output ending in `.pptx` or with `--pptx`. Pass the organization's deck as `--pptx-template org.pptx` so the size, theme, and layouts match. Sample slides already in that file are dropped. Install `pip install "latex-docx[pptx]"`.

If the organization has a deck, pass it as `--pptx-template`. That file owns colors, fonts, and the slide size. Do not also set `::font`, `::color`, `::background`, or `::size`; those lines are ignored when a template is present. List its layouts with `latex-docx --pptx-layouts org.pptx`.

If there is no organization deck, omit `--pptx-template`. The built-in layouts are available, and the agent sets the look in the Markdown:

```bash
latex-docx --pptx-layouts
```

```markdown
::font Georgia
::color 1F4E79
::background F7F4EF
::size widescreen
```

`::size` is `widescreen` or `16:9` (13.333 by 7.5 in), `standard` or `4:3` (10 by 7.5 in), or inches such as `13.333x7.5`. `::color` and `::background` are `RRGGBB`. These lines may sit anywhere in the file. They are not slide text.

Each layouts line is `index`, then the layout name, then its placeholders (`TITLE`, `BODY`, `PICTURE`, …). Put a directive on its own line inside a slide to use that name:

```markdown
::layout Title Only

## Figure

![scheme](fig.png)
```

The name must match the list, including spaces. Slides without `::layout` use `--title-layout` for a `#` slide and `--content-layout` for the rest. Omit those flags to let the converter pick a title layout and a content layout from the placeholders.

Word edits do not flow back into the source. Run `--diff-docx` and copy the changes into the `.tex` or `.md` yourself.

## Commands

After `pip install -e ".[omml,pptx,dev]"`:

```bash
latex-docx article.tex -o out_article.docx
latex-docx article.tex --template plain -o out_plain.docx
latex-docx notes.md -o out_notes.docx --lang en
latex-docx notes.md -o out_notes.docx --pdf
latex-docx out_notes.docx -o out_notes.pdf
latex-docx slides.md -o out_slides.pptx --pptx-template org.pptx
latex-docx slides.md --pptx --title-layout "Title Slide" --content-layout "Title and Content"
latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

`python -m latex_docx` is the same entry point.

```python
from latex_docx import convert, diff, docx_to_pdf, markdown_to_pptx

convert("notes.md", "out_notes.docx", lang="en")
convert("notes.md", "out_notes.docx", pdf=True)
convert("article.tex", "out_article.docx", template="plain", checklist=True)
docx_to_pdf("out_notes.docx")
markdown_to_pptx(
    "slides.md",
    "out_slides.pptx",
    template="org.pptx",
    title_layout="Title Slide",
)
diff("out_base.docx", "out_edited.docx", "out_changes.md")
```

`convert` chooses LaTeX or Markdown from the suffix and does not print the TD-J checklist unless asked. A copy-paste wrapper lives in `examples/wrap.py` and `examples/Makefile`.

```bash
make convert TEX=article.tex BASE=out_base.docx BY=agent
make convert-md MD=notes.md BASE=out_notes.docx BY=agent
make prepare-edit EDITED=out_edited.docx
make diff BASE=out_base.docx EDITED=out_edited.docx CHANGES=out_changes.md BY=agent
make pptx SLIDES=slides.md PPTX=out_slides.pptx PPTX_TEMPLATE=org.pptx BY=agent
```

A result name without the `out` prefix gets one (`letter.docx` becomes `out_letter.docx`). Do not commit large generated DOCX, PDF, or PPTX files unless the person asked.

Formulas without the `omml` extra become italic text. Word math needs `pip install "latex-docx[omml]"`. PDF on macOS and Windows needs Microsoft Word. On Windows use `pip install "latex-docx[win]"` so `pywin32` is installed. On Linux, PDF is written by LibreOffice (`soffice` on `PATH`); that page can differ from Word. DOCX and PPTX do not need Word or LibreOffice.

## LaTeX

The parser is narrow: a typical article, not arbitrary TeX. It understands `equation` / `align`, `figure` / `subfigure`, `tabular` / `tabularx` / `longtable`, `thebibliography` and a sibling `.bbl`, and `\input` / `\include`. Authors can come from `authors.json` (`--authors id1,id2`). Without that flag, JSON does not overwrite a careful author block already in the `.tex`.

Before a TD-J submission the converter prints a checklist: UDC, titles, abstracts, keywords, introduction, conclusion, references, author information. Fix those items in the `.tex`.

Limits: `troubles.md`.

## Markdown for DOCX

The first `# ...` heading becomes the centered document title when no earlier block exists. Later `#`, `##`, and `###` stay headings. A YAML title keeps the following `#` as a heading:

```markdown
---
title: Memo
---

# Reason
```

Only the `title:` line is read. Other YAML keys are ignored.

The parser accepts paragraphs separated by a blank line, `**bold**`, `*italic*`, `***both***`, `` `code` ``, links as text plus the URL in parentheses, one-level `-` / `*` / `+` and `1.` lists, `- [ ]` and `- [x]`, a GFM table whose first row is a bold header, an image on its own line `![caption](fig.png)` with a path relative to the `.md`, a fenced code block in Consolas 11 pt with no highlighting, `>` quotes, inline `$E=mc^2$`, unnumbered `$$...$$`, and a `---` rule that is skipped.

`foo_bar` is not italic. Nested lists collapse to one level. Footnotes, HTML, and definition lists are skipped. Markdown images are not numbered `Рис.` / `Таблица`.

## Markdown for PPTX

```markdown
# Talk title
Subtitle

---

## First slide

- point
- another point

???
Say this aloud.

---

## Figure

![scheme](fig.png)
```

## After Word

1. Build the base DOCX.
2. Copy it to the file the person will edit (`make prepare-edit`).
3. After the edits: `latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md`.
4. Copy the meaning of those edits into the source and rebuild.

The diff report includes OMML text. Do not try a full DOCX import into TeX or Markdown.

## What to avoid

Do not invent measurements, authors, or bibliographic data. Do not overwrite the person's `.tex`, `.md`, or edited DOCX. Keep this converter instead of pandoc when the TD-J layout matters: the layout lives in `DocxBuilder`. If Markdown cannot express numbered formulas, a complex table, or a bilingual header, write LaTeX.

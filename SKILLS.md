# Preparing documents with latex-docx

Instructions for an agent. Write the source. The DOCX, PDF, or PPTX is a build product, plus the file a person edits in Word or PowerPoint.

A TD-J article is LaTeX. A note, report, letter, or draft is Markdown. A talk is Markdown too, built to PPTX. The DOCX builder uses Times New Roman 14 pt, margins of about 2 cm, line spacing 1.5, a first-line indent, and a centered page number. Journal captions inside that DOCX stay Russian.

## Which source

TD-J article, with formulas, figures, references, and a bilingual header: `.tex`, template `tdj` (the default). Russian header only: `--template plain`.

A document without the journal header: `.md` or `.markdown`. `--template` does not apply. The builder uses `plain` and does not add UDC, abstracts, or the English block.

A presentation: the same Markdown, with output ending in `.pptx` or with `--pptx`. Pass the organization's deck as `--pptx-template org.pptx` so the size, theme, and layouts match. Sample slides already in that file are dropped. If the layout names are not `Title Slide` and `Title and Content`, pass the exact names with `--title-layout` and `--content-layout`. Install `pip install "latex-docx[pptx]"`.

Word edits do not flow back into the source. Run `--diff-docx` and copy the changes into the `.tex` or `.md` yourself.

## Commands

After `pip install -e ".[omml,pptx,dev]"`:

```bash
latex-docx article.tex -o out_article.docx
latex-docx article.tex --template plain -o out_plain.docx
latex-docx notes.md -o out_notes.docx
latex-docx notes.md -o out_notes.docx --pdf
latex-docx out_notes.docx -o out_notes.pdf
latex-docx slides.md -o out_slides.pptx --pptx-template org.pptx
latex-docx slides.md --pptx --title-layout "Title Slide" --content-layout "Title and Content"
latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

`python -m latex_docx` is the same entry point.

```python
from latex_docx import convert, diff, docx_to_pdf, markdown_to_pptx

convert("notes.md", "out_notes.docx")
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

Formulas without the `omml` extra become italic text. Word math needs `pip install "latex-docx[omml]"`. PDF needs Microsoft Word on the machine. Windows also needs `pywin32`.

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

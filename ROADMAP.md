# Roadmap

What 0.2.0 already does, and what is worth doing next. This is not a promise of later releases. The open list in [`troubles.md`](troubles.md) is empty: the items below are product steps, not leftover bugs.

## Shipped

- A typical LaTeX article becomes a TD-J Word file: Times New Roman 14 pt, margins of about 2 cm, OMML when the `omml` extra is installed.
- Light Markdown uses the same DOCX builder.
- Markdown becomes a PowerPoint file. An organization `.pptx` supplies size, theme, and layouts.
- PDF is the DOCX saved by Microsoft Word, so the page matches Word.
- A paragraph diff of two DOCX files, written as Markdown.
- Library call `convert()`, CLI `latex-docx`, and `lang="ru"` or `lang="en"`.
- English, Russian, and Chinese GitHub Pages.
- An organization `.docx` or `.dotx` (`--docx-template`) keeps styles, page setup, headers, and footers. Sample body text is dropped. Times New Roman is not forced on top of that file.
- Section titles use Word heading styles (`Heading 1` and the localized equivalents), so a table of contents and the navigation pane can see them. Without a style template the heading styles are still Times New Roman 14 pt.
- Markdown lists can be nested. `[^id]` becomes a Word footnote when `[^id]:` is defined.
- `--template article` is the one-block layout for a paper that is not a TD-J submission. Pair it with `--lang en` for English labels.

## Next

Nothing is queued. A full DOCX→TeX import stays out of scope: the diff report is how Word edits come back to the source.

## Out of scope

- Replacing pandoc.
- Reading an arbitrary DOCX and reconstructing the original LaTeX.
- A custom PDF engine. macOS and Windows still ask Word to save the DOCX. Linux uses LibreOffice because Word is not there, and that page can differ.

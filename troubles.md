# latex-docx issues and backlog

Known limits. Closed items are marked `[x]`.

## Done

- [x] `**` is not ignored
- [x] roots do not force an arbitrary degree by default
- [x] affiliation markers use superscripts
  (regression: long affiliations from `.tex` without ¹… no longer block `authors.json`)
- [x] author names in `authors.json` / `authors.example.json`
- [x] several front-matter templates (`--template tdj|plain`)
- [x] tables: no fill, black borders
- [x] table / figure / equation numbers (not `tab:` / `fig:` / `eq:` leftovers)
- [x] `--diff-docx` writes a Markdown report (full DOCX→TeX import stays unreliable)
- [x] tables no longer leak `|c|c|c|p5.8cm|` into the first cell
- [x] equation numbers, centering, number on the right
- [x] inline formulas become OMML
- [x] `cases` / `matrix`: `\\` and `&` survive prep
- [x] `align` becomes `aligned` (it does not collapse to one line)
- [x] `\ref{sec:…}` / `\label` on section and subsection
- [x] `\_` inside `\mathrm` does not become a subscript
- [x] `apply_authors_db` does not overwrite a careful author block from TeX
- [x] font `w:rFonts` ascii/hAnsi (macOS Word)
- [x] the package is split into modules, with pytest regressions
- [x] `\textbf{…$math$…}` keeps OMML (not a Unicode fallback)
- [x] table cells are runs plus OMML, not plain text only
- [x] tall numbered displays (`cases` / fractions) use display OMML in a 2-column table
- [x] PDF figures are rasterized with `pdftoppm` / ImageMagick
- [x] `--diff-docx` reads OMML text (`m:t`) and sees formula edits
- [x] abstracts built with `paper.cite_map` turn `\cite{key}` into `[N]`
- [x] `by=agent` detection: more env markers, plus argv heuristics
- [x] `\input` / `\include` (recursive, with a cycle guard)
- [x] `tabularx` / `longtable` / `\multicolumn` / `\multirow` / booktabs
- [x] BibTeX / biblatex via a sibling `.bbl` (`\bibliography` / `\addbibresource`)
- [x] several `\includegraphics` in one `figure` share one number; `subfigure` becomes (a)/(b)
- [x] `find_balanced` inside `\vspace` uses `_safe_balanced`, so broken TeX does not crash the parser
- [x] `latex2mathml` / OMML: warnings instead of a silent `None`
- [x] `_unwrap_omml_boxes` also unwraps `<m:box>` that has `boxPr`
- [x] `\bmod` becomes a spaced `mod` in flat text and `\operatorname{mod}` in prep
- [x] table caption sits above the table, right-aligned (TD-J)
- [x] light Markdown → DOCX (same `DocxBuilder`, CLI by `.md` suffix)
- [x] PDF from DOCX through Microsoft Word (`--pdf` / `docx_to_pdf`)
- [x] Markdown → PPTX, organization template (`--pptx-template`)

## Open

_(empty)_

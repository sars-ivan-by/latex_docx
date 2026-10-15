# Changelog

## 0.2.0

- `lang="ru"` or `lang="en"` on the library, the CLI (`--lang`), and the example wrapper. The default for documents is Russian. `template="tdj"` still prints both front-matter blocks. Shared labels and the TD-J checklist follow `lang`. `diff()` stays English unless `lang="ru"`.
- `--docx-template` accepts an organization `.docx` or `.dotx` and keeps its styles and page setup.
- Headings are Word heading styles. Nested Markdown lists and `[^id]` footnotes are written into the DOCX.
- `--template article` is a one-block layout for a paper that is not aimed at TD-J.
- GitHub Pages in Chinese, next to English and Russian.
- Windows PDF export starts its own Word instance via `pywin32` (`pip install "latex-docx[win]"`), writes the PDF with `ExportAsFixedFormat`, and closes that instance. macOS still uses AppleScript and leaves Word open.
- Linux PDF export uses LibreOffice (`soffice`). The page can differ from Word. DOCX and PPTX are unchanged on Linux.

## 0.1.0

First packaged release.

- LaTeX article to a TD-J DOCX, with OMML when the `omml` extra is installed.
- Light Markdown to the same DOCX builder.
- Markdown to PPTX, including an organization `.pptx` template.
- DOCX to PDF through the local Microsoft Word install.
- DOCX diff report.
- English and Russian GitHub Pages under `docs/`.

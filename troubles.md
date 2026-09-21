# Проблемы и бэклог latex-docx

Известные ограничения конвертера. Закрытые пункты помечены `[x]`.

## Сделано

- [x] нет игнорирования **
- [x] знак корня по дефолту ставит с произвольной степенью
- [x] нет надстрочных цифр для указания какой автор откуда
  (регресс: длинные аффилиации из `.tex` без ¹… больше не блокируют `authors.json`)
- [x] ФИО авторов в `authors.json` / `authors.example.json`
- [x] поддержка нескольких форматов шаблонов (`--template tdj|plain`)
- [x] таблички: без заливки, чёрные границы
- [x] номера таблиц / картинок / формул (не артефакты tab:/fig:/eq:)
- [x] `--diff-docx` → markdown-отчёт (полный автоимпорт DOCX→TeX ненадёжен)
- [x] в таблицах артефакт "|c|c|c|p5.8cm|" в первой ячейке
- [x] номера формул, центрирование, номер справа
- [x] средитекстовые формулы → OMML
- [x] `cases` / `matrix`: `\\` и `&` защищены в prep
- [x] `align` → `aligned` (не схлопывается в одну строку)
- [x] `\ref{sec:…}` / `\label` на section/subsection
- [x] `\_` в `\mathrm` не превращается в subscript
- [x] `apply_authors_db` не затирает аккуратных авторов из tex
- [x] шрифт `w:rFonts` ascii/hAnsi (macOS Word)
- [x] пакет разбит на модули; есть pytest-регрессии
- [x] `\textbf{…$math$…}` сохраняет OMML (не Unicode)
- [x] ячейки таблиц — Runs + OMML (не только plain text)
- [x] высокие numbered display-формулы (`cases`/дроби) → display OMML в 2-col table
- [x] PDF-рисунки растеризуются через `pdftoppm` / ImageMagick
- [x] `--diff-docx` читает текст OMML (`m:t`), видит правки формул
- [x] аннотация при сборке с `paper.cite_map` — `\cite{key}` → `[N]`
- [x] автодетект `by=agent`: больше env-маркеров + argv/psutil heuristics
- [x] `\input` / `\include` (рекурсивно, с защитой от циклов)
- [x] `tabularx` / `longtable` / `\multicolumn` / `\multirow` / booktabs
- [x] BibTeX / biblatex через соседний `.bbl` (`\bibliography` / `\addbibresource`)
- [x] несколько `\includegraphics` в одном `figure` — один номер; `subfigure` → (a)/(b)
- [x] `find_balanced` в `\vspace` через `_safe_balanced` — битый tex не роняет парсер
- [x] `latex2mathml` / OMML: предупреждения вместо молчаливого `None`
- [x] `_unwrap_omml_boxes` снимает и `<m:box>` с `boxPr`
- [x] `\bmod` → spaced `mod` в flat text / `\operatorname{mod}` в prep
- [x] подпись таблицы: перед таблицей, выравнивание справа (TD-J)

## Открыто

_(пусто)_

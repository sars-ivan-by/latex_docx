# Проблемы и бэклог latex-docx

Известные ограничения конвертера. Закрытые пункты помечены `[x]`.

## Сделано

- [x] нет игнорирования **
- [x] знак корня по дефолту ставит с произвольной степенью
- [x] нет надстрочных цифр для указания какой автор откуда
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

## Открыто

### Критичные / заметные в DOCX

- [ ] `\textbf{…$math$…}` теряет OMML — внутренний math уходит в Unicode
- [ ] ячейки таблиц — только plain text (без OMML)
- [ ] display-формулы с номером всегда inline OMML (высокие `cases`/дроби мельче)
- [ ] PDF-рисунки не вставляются — только raster

### Diff / лог / процесс

- [ ] `--diff-docx` почти не видит правки формул (в `.text` у eq только `(1)`)
- [ ] аннотация при сборке без `cite_map` — `\cite{key}` может остаться ключом
- [ ] автодетект `by=agent` хрупкий (env)

### Парсер / покрытие LaTeX

- [ ] нет `\input` / `\include`
- [ ] нет `tabularx` / `longtable` / `\multicolumn` / `\multirow` / booktabs
- [ ] `\bibitem` только `thebibliography` — BibTeX / biblatex нет
- [ ] несколько `\includegraphics` в одном `figure` — один номер на все
- [ ] `find_balanced` в `\vspace` без try — битый tex может уронить парсер

### OMML / качество формул

- [ ] `latex2mathml` глотает исключения молча
- [ ] `_unwrap_omml_boxes` не трогает `<m:box>` с `boxPr`
- [ ] `\bmod` / spacing в плоском виде

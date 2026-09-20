# latex-docx

Конвертер научных статей LaTeX → Microsoft Word (DOCX) с оформлением близким к требованиям журнала [TD-J](https://td-j.ru/): Times New Roman 14 pt, поля ~2 см, номера формул справа, OMML-математика.

Не замена pandoc: узкий парсер под типичные статьи (equation/align, figure, tabular, thebibliography, bilingual front matter).

## Установка

```bash
git clone https://github.com/sars-ivan-by/latex_docx.git
cd latex_docx
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e ".[omml,dev]"
```

Без OMML формулы уйдут в plain italic; для нормальных формул Word ставьте extra `omml`.

## Быстрый старт

```bash
# .tex → DOCX
latex-docx article.tex -o out_article.docx

# шаблон без английской шапки
latex-docx article.tex --template plain -o out_plain.docx

# авторы из JSON (см. authors.example.json)
cp authors.example.json authors.json   # отредактируйте
latex-docx article.tex --authors alice,bob -o out_article.docx

# diff двух DOCX → markdown-отчёт (правки из Word переносите в .tex вручную)
latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

То же через модуль: `python -m latex_docx …`.

### Makefile (опционально)

```bash
make convert TEX=article.tex BASE=out_base.docx
make prepare-edit EDITED=out_edited.docx
# правки в Word…
make diff BASE=out_base.docx EDITED=out_edited.docx CHANGES=out_changes.md
```

Все выходные имена с приставкой `out` (см. `.gitignore`).

## Авторы (`authors.json`)

Поиск файла: `--authors-json` → `./authors.json` (cwd) → `authors.json` в корне репо → `authors.example.json`.

Скопируйте пример и заполните `id`, ФИО, `org_ids`, блок `organizations`. Явный `--authors id1,id2` всегда перезаписывает авторов; без флага JSON не затирает уже аккуратный front matter из `.tex` (надстрочные индексы организаций).

## API

```python
from pathlib import Path
from latex_docx import latex_to_docx, LatexParser

latex_to_docx("article.tex", "out_article.docx", template="tdj")
paper = LatexParser(Path("article.tex")).parse()
```

## Тесты

```bash
pytest
```

## Ограничения

Парсер не тянет `\input`, BibTeX, `longtable`, PDF-рисунки и часть OMML-краёв. Актуальный бэклог: [`troubles.md`](troubles.md).

## Лицензия

MIT. См. [`LICENSE`](LICENSE).

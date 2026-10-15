# latex-docx

Конвертер научных статей LaTeX → Microsoft Word (DOCX) с оформлением близким к требованиям журнала [TD-J](https://td-j.ru/): Times New Roman 14 pt, поля ~2 см, номера формул справа, OMML-математика.

Не замена pandoc: узкий парсер под типичные статьи (equation/align, figure, tabular, thebibliography, bilingual front matter).

## Установка

Как библиотека в уже существующее окружение:

```bash
pip install "latex-docx[omml]"
# из git, если пакета нет на PyPI:
pip install "latex-docx[omml] @ git+https://github.com/sars-ivan-by/latex_docx.git"
```

Разработка этого репозитория:

```bash
git clone https://github.com/sars-ivan-by/latex_docx.git
cd latex_docx
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e ".[omml,dev]"
```

Без extra `omml` формулы уйдут в plain italic. В Word для формул ставьте `omml`.

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

# лёгкий Markdown → тот же DOCX (записка, отчёт, письмо)
latex-docx notes.md -o out_notes.docx

# PDF тем же Word из готового DOCX (вёрстка совпадает с документом)
latex-docx notes.md -o out_notes.docx --pdf
latex-docx out_notes.docx -o out_notes.pdf
```

То же через модуль: `python -m latex_docx …`.

### Makefile (опционально)

```bash
make convert TEX=article.tex BASE=out_base.docx
make convert-md MD=notes.md BASE=out_notes.docx
make prepare-edit EDITED=out_edited.docx
# правки в Word…
make diff BASE=out_base.docx EDITED=out_edited.docx CHANGES=out_changes.md
```

Все выходные имена с приставкой `out` (см. `.gitignore`).

## Авторы (`authors.json`)

Поиск файла: `--authors-json` → `./authors.json` (cwd) → `authors.json` в корне репо → `authors.example.json`.

Скопируйте пример и заполните `id`, ФИО, `org_ids`, блок `organizations`. Явный `--authors id1,id2` всегда перезаписывает авторов; без флага JSON не затирает уже аккуратный front matter из `.tex` (надстрочные индексы организаций).

## Библиотека

Один вызов смотрит на расширение и собирает DOCX. Чеклист TD-J в этом режиме молчит, пока не передать `checklist=True`.

```python
from latex_docx import convert, diff, docx_to_pdf

convert("notes.md", "out_notes.docx")
convert("article.tex", "out_article.docx", template="tdj", checklist=True)
convert("notes.md", "out_notes.docx", pdf=True)  # рядом появится out_notes.pdf
docx_to_pdf("out_notes.docx", "out_notes.pdf")
diff("out_base.docx", "out_edited.docx", "out_changes.md")
```

PDF пишет установленный Microsoft Word: на macOS через osascript, на Windows через Word и `pywin32`. Отдельный рендер не используется, поэтому страница совпадает с DOCX. Word при этом не закрывается, закрывается только открытый файл.

Чужой проект может скопировать [`examples/wrap.py`](examples/wrap.py) и [`examples/Makefile`](examples/Makefile):

```bash
cd examples
make convert SRC=sample.md
```

Низкий уровень, если нужен разбор без записи файла:

```python
from pathlib import Path
from latex_docx import latex_to_docx, markdown_to_docx, LatexParser

latex_to_docx("article.tex", "out_article.docx", template="tdj")
markdown_to_docx("notes.md", "out_notes.docx")
paper = LatexParser(Path("article.tex")).parse()
```

Как агенту выбрать `.tex` или `.md` и что парсер Markdown принимает: [`SKILLS.md`](SKILLS.md).

## Тесты

```bash
pytest
```

## Ограничения

Парсер покрывает типичные статьи (equation/align, figure/subfigure, tabular/tabularx/longtable,
thebibliography и `.bbl`). Полный автоимпорт DOCX→TeX по-прежнему ненадёжен — используйте
`--diff-docx` и переносите правки в `.tex` вручную. Актуальный бэклог: [`troubles.md`](troubles.md).

## Лицензия

MIT. См. [`LICENSE`](LICENSE).

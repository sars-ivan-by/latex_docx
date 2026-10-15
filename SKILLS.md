# Подготовка документов через latex-docx

Инструкция для агента. Пакет собирает DOCX из исходника. Научная статья живёт в LaTeX. Записка, отчёт, письмо, методичка и черновик живут в Markdown. Оба пути заканчиваются одним сборщиком: Times New Roman 14 pt, поля около 2 см, межстрочный 1,5, абзацный отступ, номер страницы внизу по центру.

Пишите исходник. DOCX только результат конвертации и файл, который человек правит в Word.

## Какой формат выбрать

Статья в журнал TD-J, с формулами, рисунками, списком литературы и двуязычной шапкой: файл `.tex`, шаблон `tdj` (он стоит по умолчанию). Русская шапка без английского блока: `--template plain`.

Документ без журнальной шапки: файл `.md` или `.markdown`. CLI смотрит на расширение и вызывает лёгкий конвертер. Флаг `--template` на Markdown не действует: сборщик берёт `plain`, УДК, аннотацию и английский блок не добавляет.

Правки, которые человек внёс в Word, обратно в исходник сами не переносятся. Снимите отчёт `--diff-docx` и перенесите правки в `.tex` или `.md` руками.

## Команды

Из каталога пакета, после `pip install -e ".[omml,dev]"`:

```bash
latex-docx article.tex -o out_article.docx
latex-docx article.tex --template plain -o out_plain.docx
latex-docx notes.md -o out_notes.docx
latex-docx notes.md -o out_notes.docx --pdf
latex-docx out_notes.docx -o out_notes.pdf
latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md
```

PDF получается из DOCX через Microsoft Word, не отдельным движком. На машине должен стоять Word. На Windows для этого вызова ещё нужен `pywin32`. Команда `make pdf` в каталоге пакета берёт уже собранный `BASE`.

То же самое: `python -m latex_docx …`.

Из кода, когда пакет уже установлен (`pip install "latex-docx[omml]"`):

```python
from latex_docx import convert, diff, docx_to_pdf

convert("notes.md", "out_notes.docx")
convert("notes.md", "out_notes.docx", pdf=True)
convert("article.tex", "out_article.docx", template="plain", checklist=True)
docx_to_pdf("out_notes.docx")
diff("out_base.docx", "out_edited.docx", "out_changes.md")
```

`convert` сам выбирает LaTeX или Markdown по расширению и по умолчанию не печатает чеклист TD-J. Пример обёртки и Makefile, которые можно скопировать в другой проект: `examples/wrap.py`, `examples/Makefile` (`cd examples && make convert SRC=sample.md`).

Makefile:

```bash
make convert TEX=article.tex BASE=out_base.docx BY=agent
make convert-md MD=notes.md BASE=out_notes.docx BY=agent
make prepare-edit EDITED=out_edited.docx
make diff BASE=out_base.docx EDITED=out_edited.docx CHANGES=out_changes.md BY=agent
```

Имя результата без приставки `out` получит её само (`letter.docx` → `out_letter.docx`). Так файлы попадают под `.gitignore`. В репозитории aspa выходы ещё и `out*.docx` / `out*.md` в корневом ignore. Крупный DOCX не коммитьте, пока человек явно не попросил.

Если пакет стоит из aspa:

```bash
pip install -e "./latex_docx[omml,dev]"
python -m latex_docx papers/.../main.tex -o out_base.docx --by agent
python -m latex_docx path/notes.md -o out_notes.docx --by agent
```

Формулы без extra `omml` уходят в обычный курсив. Для нормальных формул Word нужен `pip install 'latex-docx[omml]'`.

## LaTeX

Парсер узкий: типичная статья, не произвольный TeX. Он понимает `equation` / `align`, `figure` / `subfigure`, `tabular` / `tabularx` / `longtable`, `thebibliography` и соседний `.bbl`, `\input` / `\include`. Авторы можно подставить из `authors.json` (`--authors id1,id2`). Без флага JSON не затирает уже аккуратную шапку из `.tex`.

Перед сдачей в TD-J конвертер печатает чеклист: УДК, заглавия, аннотации, ключевые слова, введение, заключение, литература, сведения об авторах. Закройте пункты в `.tex`, а не в DOCX.

Ограничения и закрытый бэклог: `troubles.md`.

## Markdown

Первый заголовок `# ...`, если перед ним нет другого блока, становится названием документа по центру. Остальные `#`, `##`, `###` остаются заголовками в теле. Если название задано в шапке, первый `#` остаётся заголовком:

```markdown
---
title: Служебная записка
---

# Основание
```

Из шапки читается только строка `title:`. Остальные поля игнорируются.

Дальше парсер принимает:

- абзацы, разделённые пустой строкой;
- `**жирный**`, `*курсив*`, `***жирный курсив***`, `` `код` ``;
- ссылку `[текст](https://example.com)` как текст и адрес в скобках;
- маркированный список `-` / `*` / `+` и нумерованный `1.` в один уровень;
- задачу `- [ ]` и `- [x]` как ☐ и ☑;
- таблицу GitHub Flavored Markdown, первая строка становится шапкой и пишется полужирным;
- картинку отдельной строкой `![подпись](fig.png)`, путь относительно `.md`;
- блок кода в тройных обратных кавычках, шрифт Consolas 11 pt, язык подсветки не используется;
- цитату строками `> ...`;
- формулу в строке `$E=mc^2$` и выключную `$$...$$` без номера;
- горизонтальную черту `---` как пустое место, её в DOCX нет.

`foo_bar` не становится курсивом. Вложенные списки схлопываются в один уровень. Сноски, HTML, определения и автоматическая нумерация «Рис.» / «Таблица» для Markdown не делаются. Подпись картинки берётся из квадратных скобок и пишется под рисунком курсивом, без слова «Рис.».

## Как править после Word

1. Соберите базовый DOCX.
2. Скопируйте его в файл, который откроет человек (`make prepare-edit` или свой `out_edited.docx`).
3. После правок: `latex-docx --diff-docx out_base.docx out_edited.docx -o out_changes.md`.
4. Перенесите смысл правок в исходник и соберите DOCX заново.

Отчёт diff читает и текст формул OMML. Полный автоимпорт DOCX в TeX или Markdown не делайте.

## Чего не делать

Не выдумывайте числа измерений, авторов и выходные данные. Не затирайте сырой `.tex`, `.md` и снятые DOCX человека. Не ставьте pandoc вместо этого конвертера в том же репозитории: оформление держится на `DocxBuilder`. Если разметки Markdown не хватает на формулы с номерами, таблицы со сложной вёрсткой или двуязычную шапку, пишите LaTeX.

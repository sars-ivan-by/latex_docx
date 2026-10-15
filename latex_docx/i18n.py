"""Russian and English strings inserted by the converter."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple


@dataclass(frozen=True)
class Labels:
    """Boilerplate for one language. Front-matter slots keep their own labels."""

    udc: str
    abstract: str
    keywords: str
    figure: str
    table: str
    references: str
    authors: str
    image_missing: str
    image_failed: str
    abstract_skip: Tuple[str, ...]
    keywords_skip: Tuple[str, ...]


RU = Labels(
    udc="УДК",
    abstract="Аннотация.",
    keywords="Ключевые слова:",
    figure="Рис.",
    table="Таблица",
    references="Список литературы",
    authors="Сведения об авторах",
    image_missing="рисунок не найден: {name}",
    image_failed="не удалось вставить {name}: {error}",
    abstract_skip=("аннотация",),
    keywords_skip=("ключевые",),
)

EN = Labels(
    udc="UDC",
    abstract="Abstract.",
    keywords="Keywords:",
    figure="Fig.",
    table="Table",
    references="References",
    authors="About the authors",
    image_missing="image not found: {name}",
    image_failed="could not insert {name}: {error}",
    abstract_skip=("abstract",),
    keywords_skip=("keywords",),
)

_BY_CODE = {"ru": RU, "en": EN}

# Captions that already carry a number should not gain a second prefix.
CAPTION_SKIP = ("рис", "fig", "табл", "table")


def normalize_lang(lang: str | None) -> str:
    value = (lang or "ru").strip().lower()
    aliases = {"rus": "ru", "russian": "ru", "eng": "en", "english": "en"}
    value = aliases.get(value, value)
    if value not in _BY_CODE:
        raise ValueError(f"Unsupported lang {lang!r}. Use 'ru' or 'en'.")
    return value


def labels_for(lang: str | None) -> Labels:
    return _BY_CODE[normalize_lang(lang)]


def checklist_header(lang: str, *, ok: bool) -> str:
    if normalize_lang(lang) == "ru":
        if ok:
            return "Чеклист TD-J: шапка и разделы на месте."
        return "Чеклист TD-J (поправьте это в .tex и соберите снова):"
    if ok:
        return "TD-J checklist: front matter and section headings are present."
    return "TD-J checklist (fix these in the .tex, then convert again):"


def checklist_line(lang: str, key: str, **kwargs) -> str:
    ru = {
        "udc": "нет УДК (п. 2.5)",
        "title_ru": "нет заглавия на русском (п. 2.4)",
        "title_en": "нет заглавия на английском (п. 2.4)",
        "authors_ru": "нет авторов на русском (п. 2.4)",
        "authors_en": "нет ФИО авторов на английском (п. 2.4); добавьте их в .tex",
        "aff_en": "нет названия организации на английском (п. 2.4)",
        "abs_ru": "нет аннотации на русском (п. 2.4)",
        "abs_ru_short": "русская аннотация короче 100 слов ({n}); TD-J просит 100–250",
        "abs_ru_long": "русская аннотация длиннее 250 слов ({n})",
        "abs_en": "нет аннотации на английском (п. 2.4)",
        "kw_ru": "нет ключевых слов на русском (п. 2.4)",
        "kw_ru_many": "ключевых слов на русском больше 10 ({n})",
        "kw_en": "нет ключевых слов на английском (п. 2.4)",
        "refs": "нет списка литературы (п. 2.8)",
        "authors_info": "нет блока «Сведения об авторах» (п. 2.11: степень, должность, адреса, телефоны, e-mail)",
        "intro": "нет раздела «Введение» или Introduction (п. 2.6)",
        "conclusion": "нет раздела «Заключение», «Выводы» или Conclusion (п. 2.6)",
    }
    en = {
        "udc": "missing UDC (TD-J 2.5)",
        "title_ru": "missing Russian title (TD-J 2.4)",
        "title_en": "missing English title (TD-J 2.4)",
        "authors_ru": "missing Russian authors (TD-J 2.4)",
        "authors_en": "missing English author names (TD-J 2.4); add them in the .tex",
        "aff_en": "missing English affiliation (TD-J 2.4)",
        "abs_ru": "missing Russian abstract (TD-J 2.4)",
        "abs_ru_short": "Russian abstract is shorter than 100 words ({n}); TD-J asks for 100–250",
        "abs_ru_long": "Russian abstract is longer than 250 words ({n})",
        "abs_en": "missing English abstract (TD-J 2.4)",
        "kw_ru": "missing Russian keywords (TD-J 2.4)",
        "kw_ru_many": "more than 10 Russian keywords ({n})",
        "kw_en": "missing English keywords (TD-J 2.4)",
        "refs": "missing reference list (TD-J 2.8)",
        "authors_info": "missing author-information block (TD-J 2.11: degree, position, addresses, phones, e-mail)",
        "intro": 'missing a section titled "Введение" or "Introduction" (TD-J 2.6)',
        "conclusion": 'missing a section titled "Заключение", "Выводы", or "Conclusion" (TD-J 2.6)',
    }
    table = ru if normalize_lang(lang) == "ru" else en
    return table[key].format(**kwargs)


def diff_copy(lang: str | None) -> dict:
    if normalize_lang(lang) == "ru":
        return {
            "changed": "изменено",
            "removed": "удалено",
            "added": "добавлено",
            "count": "Изменений: **{n}** (сопоставление по тексту абзацев; пустые абзацы пропущены).",
            "import": "Полный перенос DOCX обратно в TeX ненадёжен (формулы OMML, таблицы, поля).",
            "by_hand": "Перенесите правки ниже в `.tex` или `.md` вручную.",
            "base": "`base` должен быть свежей сборкой того же исходника (`make convert`),",
            "edited": "а `edited` — этот файл после правок в Word.",
            "noise": "Сравнение со старым DOCX другого шаблона добавит шум.",
            "before": "**было:**",
            "after": "**стало:**",
            "empty": "_(пусто)_",
            "none": "_Отличий нет._",
        }
    return {
        "changed": "changed",
        "removed": "removed",
        "added": "added",
        "count": "Changes: **{n}** (matched by paragraph text; empty paragraphs are ignored).",
        "import": "A full DOCX→TeX import is unreliable (OMML formulas, tables, fields).",
        "by_hand": "Copy the edits below back into the `.tex` or `.md` by hand.",
        "base": "`base` should be a fresh conversion of the same source (`make convert`),",
        "edited": "and `edited` should be that file after the Word edits.",
        "noise": "Comparing against an older DOCX from another template adds noise.",
        "before": "**before:**",
        "after": "**after:**",
        "empty": "_(empty)_",
        "none": "_No differences._",
    }

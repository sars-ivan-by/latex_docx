"""Лёгкий Markdown → модель Paper (тот же конвейер, что у LaTeX)."""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

from .math_omml import make_equation, make_math_run
from .models import Figure, Heading, ListBlock, Paper, Para, Run, TableBlock

CODE_FONT = "Consolas"

_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_UL = re.compile(r"^(\s*)([-+*])\s+(.*)$")
_OL = re.compile(r"^(\s*)(\d+)[.)]\s+(.*)$")
_HR = re.compile(r"^(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$")
_FENCE = re.compile(r"^(`{3,}|~{3,})(.*)$")
_TABLE_SEP = re.compile(
    r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$"
)
_IMAGE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")
_TASK = re.compile(r"^\[([ xX])\]\s+(.*)$")


def _emit(runs: List[Run], text: str, *, bold=False, italic=False, font=None) -> None:
    if not text:
        return
    if (
        runs
        and runs[-1].math is None
        and runs[-1].bold == bold
        and runs[-1].italic == italic
        and runs[-1].font == font
        and runs[-1].size is None
    ):
        runs[-1].text += text
    else:
        runs.append(Run(text, bold=bold, italic=italic, font=font))


def _word_char(ch: str) -> bool:
    return ch.isalnum() or ch == "_"


def _wrapped(s: str, i: int, marker: str) -> Optional[Tuple[str, int]]:
    if not s.startswith(marker, i):
        return None
    j = s.find(marker, i + len(marker))
    if j == -1:
        return None
    return s[i + len(marker) : j], j + len(marker)


def _apply_style(runs: List[Run], inner: str, *, bold=False, italic=False) -> None:
    for r in parse_md_inline(inner):
        r.bold = r.bold or bold
        r.italic = r.italic or italic
        runs.append(r)


def parse_md_inline(text: str) -> List[Run]:
    """**жирный**, *курсив*, `код`, ссылки и inline-формулы $...$."""
    runs: List[Run] = []
    s = text
    i, n = 0, len(s)
    while i < n:
        if s[i] == "\\" and i + 1 < n:
            _emit(runs, s[i + 1])
            i += 2
            continue
        if s[i] == "`":
            j = s.find("`", i + 1)
            if j != -1:
                _emit(runs, s[i + 1 : j], font=CODE_FONT)
                i = j + 1
                continue
        if s[i] == "$":
            j = s.find("$", i + 1)
            if j != -1 and j > i + 1:
                runs.append(make_math_run(s[i + 1 : j]))
                i = j + 1
                continue
        if s.startswith("![", i):
            m = re.match(r"!\[([^\]]*)\]\(([^)]+)\)", s[i:])
            if m:
                _emit(runs, m.group(1) or m.group(2))
                i += m.end()
                continue
        if s.startswith("[", i):
            m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", s[i:])
            if m:
                label, url = m.group(1), m.group(2)
                runs.extend(parse_md_inline(label))
                if url and url != label:
                    _emit(runs, f" ({url})")
                i += m.end()
                continue
        for marker in ("***", "___"):
            got = _wrapped(s, i, marker)
            if got:
                inner, i = got
                _apply_style(runs, inner, bold=True, italic=True)
                break
        else:
            got = None
        if got:
            continue
        for marker in ("**", "__"):
            got = _wrapped(s, i, marker)
            if got:
                inner, i = got
                _apply_style(runs, inner, bold=True)
                break
        else:
            got = None
        if got:
            continue
        if s[i] in "*_":
            marker = s[i]
            prev_ok = i == 0 or not _word_char(s[i - 1])
            j = s.find(marker, i + 1)
            nxt = s[j + 1] if j != -1 and j + 1 < n else ""
            close_ok = j is not None and j != -1 and j > i + 1 and not (nxt and _word_char(nxt))
            if marker == "_" and not prev_ok:
                close_ok = False
            if close_ok and (marker == "*" or prev_ok):
                _apply_style(runs, s[i + 1 : j], italic=True)
                i = j + 1
                continue
        _emit(runs, s[i])
        i += 1
    return runs


def _plain(text: str) -> str:
    return "".join(r.text or "" for r in parse_md_inline(text)).strip()


def _split_row(line: str) -> List[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _is_table_start(lines: List[str], i: int) -> bool:
    if i + 1 >= len(lines):
        return False
    if "|" not in lines[i]:
        return False
    return bool(_TABLE_SEP.match(lines[i + 1].strip()))


def _list_match(line: str):
    m = _UL.match(line)
    if m:
        return False, m.group(3)
    m = _OL.match(line)
    if m:
        return True, m.group(3)
    return None


def _yaml_end(lines: List[str]) -> Optional[int]:
    """Индекс закрывающего ---, если файл начинается с короткого YAML-блока."""
    if not lines or lines[0].strip() != "---":
        return None
    for j in range(1, len(lines)):
        if lines[j].strip() == "---":
            return j
        if lines[j].strip() and ":" not in lines[j]:
            return None
    return None


def _blockquote_body(line: str) -> Optional[str]:
    if not line.startswith(">"):
        return None
    rest = line[1:]
    if rest.startswith(" "):
        rest = rest[1:]
    return rest


class MarkdownParser:
    """Подмножество CommonMark: заголовки, абзацы, списки, таблицы, код, картинки, $формулы$."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.base = self.path.parent

    def parse(self) -> Paper:
        raw = self.path.read_text(encoding="utf-8")
        lines = raw.splitlines()
        paper = Paper()
        i = 0
        yaml_end = _yaml_end(lines)
        if yaml_end is not None:
            for line in lines[1:yaml_end]:
                m = re.match(r"^title:\s*(.*?)\s*$", line)
                if m and not paper.title_ru:
                    paper.title_ru = _plain(m.group(1).strip().strip("\"'"))
            i = yaml_end + 1

        blocks: list = []
        took_title = bool(paper.title_ru)
        while i < len(lines):
            line = lines[i]
            if not line.strip():
                i += 1
                continue

            fence = _FENCE.match(line.strip())
            if fence:
                marker = fence.group(1)[0]
                body: List[str] = []
                i += 1
                while i < len(lines) and not lines[i].startswith(marker * 3):
                    body.append(lines[i])
                    i += 1
                if i < len(lines):
                    i += 1
                blocks.append(
                    Para(
                        runs=[Run("\n".join(body), font=CODE_FONT, size=11)],
                        align="left",
                        first_line=0,
                        space_before=6,
                        space_after=6,
                        size=11,
                    )
                )
                continue

            if _HR.match(line.strip()):
                i += 1
                continue

            hm = _HEADING.match(line)
            if hm:
                level = min(len(hm.group(1)), 3)
                text = _plain(hm.group(2))
                if not took_title and level == 1 and not blocks:
                    paper.title_ru = text
                    took_title = True
                else:
                    blocks.append(Heading(text=text, level=level))
                i += 1
                continue

            if _blockquote_body(line) is not None:
                parts: List[str] = []
                while i < len(lines) and _blockquote_body(lines[i]) is not None:
                    parts.append(_blockquote_body(lines[i]) or "")
                    i += 1
                quote = " ".join(p for p in parts if p.strip())
                runs = parse_md_inline(quote)
                for r in runs:
                    r.italic = True
                blocks.append(
                    Para(runs=runs, align="left", first_line=0, left_indent=0.4, space_after=6)
                )
                continue

            listed = _list_match(line)
            if listed is not None:
                ordered, _ = listed
                items: List[List[Run]] = []
                while i < len(lines):
                    hit = _list_match(lines[i])
                    if hit is None:
                        if lines[i].startswith("  ") and items:
                            extra = lines[i].strip()
                            if extra:
                                items[-1].append(Run(" " + extra))
                            i += 1
                            continue
                        break
                    same_ordered, item_text = hit
                    if same_ordered != ordered and items:
                        break
                    ordered = same_ordered
                    task = _TASK.match(item_text)
                    if task:
                        mark = "☑ " if task.group(1).lower() == "x" else "☐ "
                        item_text = mark + task.group(2)
                    items.append(parse_md_inline(item_text))
                    i += 1
                blocks.append(ListBlock(items=items, ordered=ordered))
                continue

            if _is_table_start(lines, i):
                rows: List[List[List[Run]]] = []
                header = _split_row(lines[i])
                rows.append([parse_md_inline(c) for c in header])
                i += 2  # header + separator
                while i < len(lines) and "|" in lines[i] and lines[i].strip():
                    rows.append([parse_md_inline(c) for c in _split_row(lines[i])])
                    i += 1
                blocks.append(TableBlock(rows=rows))
                continue

            img = _IMAGE.match(line.strip())
            if img:
                src = img.group(2).strip()
                path = Path(src)
                if not path.is_absolute():
                    path = (self.base / path).resolve()
                blocks.append(Figure(paths=[path], caption=_plain(img.group(1))))
                i += 1
                continue

            if line.strip().startswith("$$"):
                rest = line.strip()[2:]
                if rest.endswith("$$") and len(rest) > 2:
                    latex = rest[:-2].strip()
                    i += 1
                else:
                    chunk: List[str] = []
                    if rest:
                        chunk.append(rest)
                    i += 1
                    while i < len(lines) and "$$" not in lines[i]:
                        chunk.append(lines[i])
                        i += 1
                    if i < len(lines):
                        tail = lines[i].split("$$", 1)[0]
                        if tail.strip():
                            chunk.append(tail)
                        i += 1
                    latex = "\n".join(chunk).strip()
                if latex:
                    blocks.append(make_equation(latex, number=None))
                continue

            para_lines = [line.strip()]
            i += 1
            while i < len(lines) and lines[i].strip():
                nxt = lines[i]
                if (
                    _HEADING.match(nxt)
                    or _FENCE.match(nxt.strip())
                    or _list_match(nxt)
                    or _blockquote_body(nxt) is not None
                    or _is_table_start(lines, i)
                    or _IMAGE.match(nxt.strip())
                    or nxt.strip().startswith("$$")
                    or _HR.match(nxt.strip())
                ):
                    break
                para_lines.append(nxt.strip())
                i += 1
            blocks.append(Para(runs=parse_md_inline(" ".join(para_lines))))

        paper.body = blocks
        return paper

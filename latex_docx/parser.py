"""LaTeX document parser."""
from __future__ import annotations

import re
from pathlib import Path
from typing import List

from .inline import (
    extract_caption,
    extract_includegraphics,
    parse_author_block,
    parse_inline,
    parse_tabular,
    resolve_bibliography,
    runs_to_plain,
)
from .math_omml import make_equation
from .models import (
    Block,
    Equation,
    Figure,
    Heading,
    ListBlock,
    Paper,
    Para,
    Run,
    TableBlock,
)
from .util import (
    expand_inputs,
    extract_command_arg,
    extract_environment,
    find_balanced,
    strip_comments,
)

class LatexParser:
    def __init__(self, tex_path: Path):
        self.path = tex_path
        self.base = tex_path.parent
        raw = strip_comments(tex_path.read_text(encoding="utf-8"))
        self.tex = expand_inputs(raw, self.base)
        self.cite_map: dict = {}
        self.ref_labels: dict = {}
        self.fig_n = self.tab_n = self.eq_n = 0
        self.sec_counters = [0, 0, 0]  # section, subsection, subsubsection

    def parse(self) -> Paper:
        paper = Paper()
        refs, cite_map = resolve_bibliography(self.tex, self.path)
        paper.references = refs
        paper.cite_map = cite_map
        self.cite_map = cite_map

        title = extract_command_arg(self.tex, "title")
        if title:
            paper.title_ru = runs_to_plain(parse_inline(title, self.cite_map))

        author = extract_command_arg(self.tex, "author")
        if author:
            paper.authors_ru, paper.affiliations_ru = parse_author_block(author)

        doc_env = extract_environment(self.tex, "document")
        if not doc_env:
            raise ValueError("\\begin{document} not found")
        body, _, _ = doc_env

        bib = extract_environment(body, "thebibliography")
        if bib:
            body = body[: bib[1]] + body[bib[2] :]
        # Drop \\bibliography / biblatex commands from body walk
        body = re.sub(r"\\bibliography\s*\{[^}]*\}", "", body)
        body = re.sub(r"\\addbibresource\s*\{[^}]*\}", "", body)
        body = re.sub(r"\\printbibliography\b(?:\[[^\]]*\])?", "", body)

        body = re.sub(r"\\maketitle\b", "", body)
        body = re.sub(r"\\tableofcontents\b", "", body)

        # Drop title/author/date blocks if they sit inside document (common in our papers)
        for cmd in ("title", "author", "date", "thanks"):
            while True:
                m = re.search(rf"\\{cmd}\s*\{{", body)
                if not m:
                    break
                try:
                    _, end = find_balanced(body, m.end() - 1)
                except ValueError:
                    body = body[: m.start()] + body[m.end() :]
                    break
                body = body[: m.start()] + body[end:]

        self._walk(body, paper)
        self._extract_authors_info(paper)
        self._resolve_refs(paper)
        return paper

    @staticmethod
    def _extract_authors_info(paper: Paper):
        """Move «Сведения об авторах» section from body into paper.authors_info."""
        new_body: List[Block] = []
        i = 0
        while i < len(paper.body):
            block = paper.body[i]
            if isinstance(block, Heading) and "сведения об авторах" in block.text.lower():
                i += 1
                while i < len(paper.body) and isinstance(paper.body[i], Para):
                    text = runs_to_plain(paper.body[i].runs)
                    if text:
                        paper.authors_info.append(text)
                    i += 1
                continue
            new_body.append(block)
            i += 1
        paper.body = new_body

    def _walk(self, body: str, paper: Paper):
        i, n = 0, len(body)
        para_buf: List[str] = []

        def flush_para():
            nonlocal para_buf
            if not para_buf:
                return
            chunk = "".join(para_buf).strip()
            para_buf = []
            if chunk:
                self._handle_paragraph(chunk, paper)

        while i < n:
            m_env = re.match(r"\\begin\{([a-zA-Z*]+)\}", body[i:])
            if m_env:
                flush_para()
                name = m_env.group(1)
                env = extract_environment(body, name, i)
                if not env:
                    i += m_env.end()
                    continue
                env_body, _, end = env
                self._handle_env(name, env_body, paper)
                i = end
                continue

            m_sec = re.match(r"\\(section|subsection|subsubsection)(\*?)\s*\{", body[i:])
            if m_sec:
                flush_para()
                level = {"section": 1, "subsection": 2, "subsubsection": 3}[m_sec.group(1)]
                starred = bool(m_sec.group(2))
                title, end = find_balanced(body, i + m_sec.end() - 1)
                labels = re.findall(r"\\label\{([^}]+)\}", title)
                title = re.sub(r"\\label\{[^}]*\}", "", title)
                m_lab = re.match(r"\s*\\label\{([^}]+)\}", body[end:])
                if m_lab:
                    labels.append(m_lab.group(1))
                    end += m_lab.end()
                if not starred:
                    idx = level - 1
                    self.sec_counters[idx] += 1
                    for j in range(idx + 1, 3):
                        self.sec_counters[j] = 0
                    num = ".".join(str(self.sec_counters[j]) for j in range(level))
                    for lab in labels:
                        self.ref_labels[lab.strip()] = num
                paper.body.append(Heading(runs_to_plain(parse_inline(title, self.cite_map)), level=level))
                i = end
                continue

            if body.startswith(r"\[", i):
                flush_para()
                j = body.find(r"\]", i + 2)
                math = body[i + 2 : j] if j != -1 else body[i + 2 :]
                paper.body.append(make_equation(math, number=None))
                i = j + 2 if j != -1 else n
                continue

            if body.startswith("$$", i):
                flush_para()
                j = body.find("$$", i + 2)
                math = body[i + 2 : j] if j != -1 else body[i + 2 :]
                paper.body.append(make_equation(math, number=None))
                i = j + 2 if j != -1 else n
                continue

            if body[i] == "\n":
                if i + 1 < n and body[i + 1] == "\n":
                    flush_para()
                    i += 2
                    while i < n and body[i] == "\n":
                        i += 1
                    continue
                para_buf.append(" ")
                i += 1
                continue

            para_buf.append(body[i])
            i += 1
        flush_para()

    def _handle_env(self, name: str, env_body: str, paper: Paper):
        if name in ("equation", "equation*", "align", "align*", "gather", "gather*", "multline", "multline*"):
            labels = re.findall(r"\\label\{([^}]+)\}", env_body)
            math = re.sub(r"\\label\{[^}]*\}", "", env_body)
            math = re.sub(r"\\nonumber\b", "", math)
            numbered = not name.endswith("*")
            num = None
            if numbered:
                self.eq_n += 1
                num = self.eq_n
                for lab in labels:
                    self.ref_labels[lab] = str(self.eq_n)
            paper.body.append(make_equation(math, number=num))
            return

        if name == "figure":
            self._handle_figure(env_body, paper)
            return

        if name == "table":
            self.tab_n += 1
            cap = extract_caption(env_body)
            for lab in re.findall(r"\\label\{([^}]+)\}", env_body):
                self.ref_labels[lab] = str(self.tab_n)
            rows = self._extract_table_rows(env_body)
            paper.body.append(TableBlock(rows=rows, caption=cap, number=self.tab_n))
            return

        # Standalone tabular* / longtable outside table env
        if name in ("tabular", "tabular*", "tabularx", "longtable"):
            self.tab_n += 1
            rows = parse_tabular(env_body, self.cite_map)
            paper.body.append(TableBlock(rows=rows, caption="", number=self.tab_n))
            return

        if name in ("itemize", "enumerate"):
            items = []
            for part in re.split(r"\\item\b", env_body)[1:]:
                part = part.strip()
                if part:
                    items.append(parse_inline(part, self.cite_map))
            paper.body.append(ListBlock(items=items, ordered=(name == "enumerate")))
            return

        if name == "abstract":
            text = runs_to_plain(parse_inline(env_body, self.cite_map))
            if not paper.abstract_ru:
                paper.abstract_ru = text
                paper.abstract_ru_tex = env_body.strip()
            else:
                paper.abstract_en = text
                paper.abstract_en_tex = env_body.strip()
            return

        if name in ("center", "flushleft", "flushright", "minipage", "subfigure"):
            cleaned = re.sub(r"\\begin\{[^}]+\}", "", env_body)
            cleaned = re.sub(r"\\end\{[^}]+\}", "", cleaned)
            cleaned = re.sub(r"\\includegraphics(?:\[[^\]]*\])?\s*\{[^}]+\}", "", cleaned)
            cleaned = cleaned.strip()
            if cleaned:
                self._handle_paragraph(cleaned, paper)
            return

        cleaned = re.sub(r"\\begin\{[^}]+\}|\\end\{[^}]+\}", "", env_body).strip()
        if cleaned:
            self._handle_paragraph(cleaned, paper)

    def _extract_table_rows(self, env_body: str):
        for tab_name in ("tabularx", "longtable", "tabular*", "tabular"):
            tab = extract_environment(env_body, tab_name)
            if tab:
                return parse_tabular(tab[0], self.cite_map)
        return []

    def _handle_figure(self, env_body: str, paper: Paper):
        """One figure number for the environment; subfigure → (a)/(b) panels."""
        self.fig_n += 1
        fig_num = self.fig_n

        subfigs = list(re.finditer(r"\\begin\{subfigure\}", env_body))
        if subfigs:
            letter_ord = 0
            for m in subfigs:
                env = extract_environment(env_body, "subfigure", m.start())
                if not env:
                    continue
                sub_body, _, _ = env
                letter_ord += 1
                letter = chr(ord("a") + letter_ord - 1)
                paths = extract_includegraphics(sub_body, self.base)
                cap = extract_caption(sub_body)
                cap = f"({letter}) {cap}".strip() if cap else f"({letter})"
                for lab in re.findall(r"\\label\{([^}]+)\}", sub_body):
                    self.ref_labels[lab] = f"{fig_num}{letter}"
                if paths:
                    paper.body.append(Figure(paths=paths, caption=cap, number=fig_num))
            # Parent caption outside subfigures
            stripped = env_body
            while True:
                env = extract_environment(stripped, "subfigure")
                if not env:
                    break
                _, a, b = env
                stripped = stripped[:a] + stripped[b:]
            parent_cap = extract_caption(stripped)
            for lab in re.findall(r"\\label\{([^}]+)\}", stripped):
                self.ref_labels[lab] = str(fig_num)
            if parent_cap:
                paper.body.append(Figure(paths=[], caption=parent_cap, number=fig_num))
            return

        paths = extract_includegraphics(env_body, self.base)
        cap = extract_caption(env_body)
        for lab in re.findall(r"\\label\{([^}]+)\}", env_body):
            self.ref_labels[lab] = str(fig_num)
        # Several \\includegraphics → one number, all images under one caption
        paper.body.append(Figure(paths=paths, caption=cap, number=fig_num))

    def _handle_paragraph(self, chunk: str, paper: Paper):
        # Drop standalone labels / leftover preamble-ish commands
        chunk = re.sub(r"\\label\{[^}]*\}", "", chunk)
        chunk = re.sub(r"\\vspace\*?\s*\{[^}]*\}", "", chunk)
        chunk = re.sub(r"\s+", " ", chunk).strip()
        if not chunk or chunk in {r"\noindent", "{}"}:
            return
        raw = chunk

        if re.search(r"УДК", raw) and len(raw) < 160:
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.udc = re.sub(r"^УДК:?\s*", "", text, flags=re.I).strip()
            return

        if re.search(r"\\textbf\{Ключевые слова\}|Ключевые слова:", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.keywords_ru = re.sub(r"^Ключевые слова:?\s*", "", text, flags=re.I).strip()
            return

        if re.search(r"\\textbf\{Keywords\}|Keywords:", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.keywords_en = re.sub(r"^Keywords:?\s*", "", text, flags=re.I).strip()
            return

        if re.search(r"\\textbf\{Title\}|(?:\\noindent\s*)?Title:", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.title_en = re.sub(r"^Title:?\s*", "", text, flags=re.I).strip()
            return

        if re.search(r"\\textbf\{Authors\}|(?:\\noindent\s*)?Authors:", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.authors_en = re.sub(r"^Authors:?\s*", "", text, flags=re.I).strip()
            return

        if re.search(r"\\textbf\{Affiliations\}|(?:\\noindent\s*)?Affiliations:", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.affiliations_en = re.sub(r"^Affiliations:?\s*", "", text, flags=re.I).strip()
            return

        if re.search(r"\\textbf\{Аннотация\}|Аннотация\.", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.abstract_ru = re.sub(r"^Аннотация\.?\s*", "", text, flags=re.I).strip()
            paper.abstract_ru_tex = re.sub(
                r"^(?:\\noindent\s*)?(?:\\textbf\{Аннотация\}\.?|Аннотация\.)\s*",
                "",
                raw,
                flags=re.I,
            ).strip()
            return

        if re.search(r"\\textbf\{Abstract\}|Abstract\.", raw):
            text = runs_to_plain(parse_inline(raw, self.cite_map))
            paper.abstract_en = re.sub(r"^Abstract\.?\s*", "", text, flags=re.I).strip()
            paper.abstract_en_tex = re.sub(
                r"^(?:\\noindent\s*)?(?:\\textbf\{Abstract\}\.?|Abstract\.)\s*",
                "",
                raw,
                flags=re.I,
            ).strip()
            return

        text_runs = parse_inline(raw, self.cite_map)
        plain = runs_to_plain(text_runs)
        if not plain:
            return

        if not paper.body and not paper.abstract_ru and len(plain) > 80 and self._looks_russian(plain):
            paper.abstract_ru = plain
            paper.abstract_ru_tex = raw
            return

        if not paper.body and paper.abstract_ru and not paper.abstract_en and len(plain) > 80 and self._looks_english(plain):
            paper.abstract_en = plain
            paper.abstract_en_tex = raw
            return

        paper.body.append(Para(runs=text_runs, align="justify", first_line=0.49))

    @staticmethod
    def _looks_russian(text: str) -> bool:
        cyr = sum(1 for c in text if "а" <= c.lower() <= "я" or c in "ёЁ")
        lat = sum(1 for c in text if "a" <= c.lower() <= "z")
        return cyr > lat

    @staticmethod
    def _looks_english(text: str) -> bool:
        cyr = sum(1 for c in text if "а" <= c.lower() <= "я" or c in "ёЁ")
        lat = sum(1 for c in text if "a" <= c.lower() <= "z")
        return lat > cyr * 2

    def _resolve_refs(self, paper: Paper):
        labels = self.ref_labels

        def fix_runs(runs: List[Run]) -> None:
            for r in runs:
                if r.ref_key and r.ref_key in labels:
                    num = labels[r.ref_key]
                    r.text = f"({num})" if r.eqref else num
                    r.ref_key = None
                    continue
                # Safety net: label leaked into merged plain text
                if r.math or not r.text:
                    continue
                for key, num in sorted(labels.items(), key=lambda kv: -len(kv[0])):
                    if key in r.text:
                        # Bare eq:* leftovers from \eqref → (N); fig:/tab: → number
                        repl = f"({num})" if key.startswith("eq:") else num
                        r.text = r.text.replace(key, repl)

        for block in paper.body:
            if isinstance(block, Para):
                fix_runs(block.runs)
            elif isinstance(block, ListBlock):
                for item in block.items:
                    fix_runs(item)
            elif isinstance(block, TableBlock):
                for row in block.rows:
                    for cell in row:
                        fix_runs(cell)
                if block.caption:
                    for key, num in sorted(labels.items(), key=lambda kv: -len(kv[0])):
                        if key in block.caption:
                            block.caption = block.caption.replace(key, num)
            elif isinstance(block, Figure) and block.caption:
                for key, num in sorted(labels.items(), key=lambda kv: -len(kv[0])):
                    if key in block.caption:
                        block.caption = block.caption.replace(key, num)

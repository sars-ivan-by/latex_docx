"""Inline TeX parsing, tables, bibliography."""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from .constants import GREEK, MATH_OPS
from .math_omml import latex_math_to_text, make_math_run
from .models import Run
from .authors import to_superscript
from .util import _safe_balanced, extract_environment, find_balanced

def parse_inline(tex: str, cite_map: Optional[dict] = None) -> List[Run]:
    cite_map = cite_map or {}
    runs: List[Run] = []
    s, i = tex, 0

    def emit(text: str, bold=False, italic=False):
        if not text:
            return
        if (
            runs
            and runs[-1].math is None
            and runs[-1].ref_key is None
            and runs[-1].bold == bold
            and runs[-1].italic == italic
            and runs[-1].size is None
        ):
            runs[-1].text += text
        else:
            runs.append(Run(text, bold=bold, italic=italic))

    while i < len(s):
        if s[i] == "$" and not (i > 0 and s[i - 1] == "\\"):
            if i + 1 < len(s) and s[i + 1] == "$":
                j = s.find("$$", i + 2)
                math = s[i + 2 : j] if j != -1 else s[i + 2 :]
                i = j + 2 if j != -1 else len(s)
            else:
                j = s.find("$", i + 1)
                math = s[i + 1 : j] if j != -1 else s[i + 1 :]
                i = j + 1 if j != -1 else len(s)
            runs.append(make_math_run(math))
            continue
        if s.startswith(r"\(", i):
            j = s.find(r"\)", i + 2)
            math = s[i + 2 : j] if j != -1 else s[i + 2 :]
            i = j + 2 if j != -1 else len(s)
            runs.append(make_math_run(math))
            continue

        # Markdown-style **bold** (appears in some .tex drafts)
        if s.startswith("**", i):
            j = s.find("**", i + 2)
            if j != -1:
                inner = s[i + 2 : j]
                for r in parse_inline(inner, cite_map):
                    if r.math:
                        runs.append(Run(text=r.text, bold=True, italic=r.italic, math=r.math))
                    else:
                        emit(r.text, bold=True, italic=r.italic)
                i = j + 2
                continue
            i += 2  # unpaired ** — ignore
            continue

        m = re.match(r"\\textsuperscript\s*\{", s[i:])
        if m:
            got = _safe_balanced(s, i + m.end() - 1)
            if not got:
                i += m.end()
                continue
            inner, end = got
            emit(to_superscript(runs_to_plain(parse_inline(inner, cite_map))))
            i = end
            continue

        m = re.match(r"\\(textbf|textit|emph|texttt|underline)\s*\{", s[i:])
        if m:
            cmd = m.group(1)
            got = _safe_balanced(s, i + m.end() - 1)
            if not got:
                i += m.end()
                continue
            inner, end = got
            for r in parse_inline(inner, cite_map):
                if cmd == "textbf":
                    r.bold = True
                elif cmd in ("textit", "emph"):
                    r.italic = True
                emit(r.text, bold=r.bold, italic=r.italic)
            i = end
            continue

        m = re.match(r"\\cite\s*\{", s[i:])
        if m:
            got = _safe_balanced(s, i + m.end() - 1)
            if not got:
                i += m.end()
                continue
            keys, end = got
            nums = []
            for k in keys.split(","):
                k = k.strip()
                nums.append(str(cite_map[k]) if k in cite_map else k)
            emit("[" + ", ".join(nums) + "]")
            i = end
            continue

        m = re.match(r"\\(eqref|ref)\s*\{", s[i:])
        if m:
            got = _safe_balanced(s, i + m.end() - 1)
            if not got:
                i += m.end()
                continue
            key, end = got
            key = key.strip()
            # Separate run — must not merge with neighbors, else resolve fails
            runs.append(Run(text=key, ref_key=key, eqref=(m.group(1) == "eqref")))
            i = end
            continue

        m = re.match(r"\\url\s*\{", s[i:])
        if m:
            got = _safe_balanced(s, i + m.end() - 1)
            if not got:
                i += m.end()
                continue
            inner, end = got
            emit(inner)
            i = end
            continue

        simple = {
            r"\&": "&", r"\%": "%", r"\$": "$", r"\#": "#", r"\_": "_",
            r"\{": "{", r"\}": "}", r"\\": "\n", r"\,": " ", r"\;": " ",
            r"\!": "", r"\quad": " ", r"\qquad": "  ", "~": " ",
            "---": "—", "--": "–", "``": "«", "''": "»",
            r"\ldots": "…", r"\dots": "…", r"\newline": "\n",
            r"\noindent": "", r"\centering": "", r"\hfill": " ", r"\and": "\n",
        }
        matched = False
        for k, v in sorted(simple.items(), key=lambda x: -len(x[0])):
            if s.startswith(k, i):
                emit(v)
                i += len(k)
                matched = True
                break
        if matched:
            continue

        m = re.match(r"\\(vspace|hspace|label|phantom|vspace\*|hspace\*)\s*\{", s[i:])
        if m:
            _, end = find_balanced(s, i + m.end() - 1)
            i = end
            continue

        m = re.match(r"\\([a-zA-Z]+\*?)\s*", s[i:])
        if m:
            name = m.group(1)
            if name in ("small", "footnotesize", "normalsize", "large", "Large",
                        "itshape", "bfseries", "ttfamily", "rmfamily", "centerline"):
                i += m.end()
                continue
            if name in GREEK:
                emit(GREEK[name], italic=True)
                i += m.end()
                continue
            if name in MATH_OPS:
                emit(MATH_OPS[name])
                i += m.end()
                continue
            i += m.end()
            continue

        emit(s[i])
        i += 1

    cleaned: List[Run] = []
    for r in runs:
        if r.math or r.ref_key:
            cleaned.append(r)
            continue
        t = re.sub(r"[ \t]+", " ", r.text)
        if t:
            cleaned.append(Run(t, bold=r.bold, italic=r.italic, size=r.size))
    return cleaned



def runs_to_plain(runs: Sequence[Run]) -> str:
    parts = []
    for r in runs:
        if r.math:
            parts.append(r.text or latex_math_to_text(r.math))
        else:
            parts.append(r.text)
    return "".join(parts).strip()



def parse_tabular(body: str) -> List[List[str]]:
    body = body.strip()
    # Drop optional [pos] and required {column spec} after \begin{tabular}
    if body.startswith("["):
        end = body.find("]")
        if end != -1:
            body = body[end + 1 :].lstrip()
    if body.startswith("{"):
        try:
            _, i = find_balanced(body, 0)
            body = body[i:].lstrip()
        except ValueError:
            pass

    body = re.sub(r"\\hline\b", "", body)
    body = re.sub(r"\\cline\{[^}]*\}", "", body)
    body = re.sub(r"\\(small|footnotesize|centering|center)\b", "", body)
    rows = []
    for part in re.split(r"\\\\", body):
        part = part.strip()
        if not part:
            continue
        cells = [c.strip() for c in part.split("&")]
        if all(not c for c in cells):
            continue
        # Skip leftover column-spec rows like |c|c|c|p{5.8cm}|
        joined = "".join(cells)
        if re.fullmatch(r"[|lcrp@\{\}0-9.\\,\s]+", joined) and "p{" in joined.replace(" ", ""):
            continue
        if re.fullmatch(r"[|lcr\s]+", joined):
            continue
        cleaned = []
        for c in cells:
            c = re.sub(r"\\textbf\s*\{([^}]*)\}", r"\1", c)
            c = re.sub(r"\\textit\s*\{([^}]*)\}", r"\1", c)
            if "$" in c or "\\" in c:
                c = runs_to_plain(parse_inline(c))
            c = c.replace("{", "").replace("}", "")
            cleaned.append(c.strip())
        rows.append(cleaned)
    return rows



def extract_caption(env_body: str) -> str:
    m = re.search(r"\\caption\s*\{", env_body)
    if not m:
        return ""
    cap, _ = find_balanced(env_body, m.end() - 1)
    return runs_to_plain(parse_inline(cap))



def extract_includegraphics(env_body: str, base: Path) -> List[Path]:
    paths = []
    for m in re.finditer(r"\\includegraphics(?:\[[^\]]*\])?\s*\{([^}]+)\}", env_body):
        rel = m.group(1).strip()
        p = (base / rel).resolve()
        if not p.exists():
            for ext in (".png", ".jpg", ".jpeg", ".pdf"):
                cand = Path(str(p) + ext) if not p.suffix else p.with_suffix(ext)
                if cand.exists():
                    p = cand
                    break
        paths.append(p)
    return paths



def parse_bibliography(tex: str):
    env = extract_environment(tex, "thebibliography")
    if not env:
        return [], {}
    body, _, _ = env
    refs, cite_map = [], {}
    items = list(re.finditer(r"\\bibitem\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}", body))
    for idx, m in enumerate(items):
        key = m.group(1)
        start = m.end()
        end = items[idx + 1].start() if idx + 1 < len(items) else len(body)
        text = re.sub(r"\s+", " ", runs_to_plain(parse_inline(body[start:end]))).strip()
        num = idx + 1
        cite_map[key] = num
        refs.append(f"{num}. {text}")
    return refs, cite_map



def parse_author_block(author_tex: str) -> Tuple[str, str]:
    parts = re.split(r"\\and\b", author_tex.strip())
    names, affs = [], []
    for part in parts:
        lines = [ln.strip() for ln in re.split(r"\\\\|\n", part) if ln.strip()]
        cleaned = []
        for ln in lines:
            ln = re.sub(r"\\texttt\s*\{([^}]*)\}", r"\1", ln)
            ln = runs_to_plain(parse_inline(ln))
            if ln:
                cleaned.append(ln)
        if not cleaned:
            continue
        names.append(cleaned[0])
        if len(cleaned) > 1:
            affs.append("\n".join(cleaned[1:]))
    return ", ".join(names), "\n".join(affs)




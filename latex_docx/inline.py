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
                # Preserve OMML / ref runs — emit() would flatten math to Unicode
                if r.math or r.ref_key:
                    runs.append(
                        Run(
                            text=r.text,
                            bold=r.bold,
                            italic=r.italic,
                            size=r.size,
                            math=r.math,
                            ref_key=r.ref_key,
                            eqref=r.eqref,
                        )
                    )
                else:
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
            got = _safe_balanced(s, i + m.end() - 1)
            if not got:
                i += m.end()
                continue
            _, end = got
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



def parse_tabular(body: str, cite_map: Optional[dict] = None) -> List[List[List[Run]]]:
    """Parse tabular / tabularx / longtable body → rows of cell Runs (OMML-capable)."""
    cite_map = cite_map or {}
    body = body.strip()
    # Drop optional [pos] and required {column spec}
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

    # booktabs / hline / cline — visual rules only
    body = re.sub(r"\\(toprule|midrule|bottomrule|hline)\b(?:\[[^\]]*\])?", "", body)
    body = re.sub(r"\\cmidrule(?:\[[^\]]*\])?\{[^}]*\}", "", body)
    body = re.sub(r"\\cline\{[^}]*\}", "", body)
    body = re.sub(r"\\(small|footnotesize|centering|center)\b", "", body)
    # \multirow{n}{*}{text} / \multicolumn{n}{c}{text} → keep inner text
    body = re.sub(
        r"\\multirow\s*\{[^}]*\}\s*\{[^}]*\}\s*\{",
        r"\\MULTISTART{",
        body,
    )
    body = re.sub(
        r"\\multicolumn\s*\{[^}]*\}\s*\{[^}]*\}\s*\{",
        r"\\MULTISTART{",
        body,
    )
    # Resolve \MULTISTART{...} to inner content
    while True:
        m = re.search(r"\\MULTISTART\{", body)
        if not m:
            break
        got = _safe_balanced(body, m.end() - 1)
        if not got:
            body = body[: m.start()] + body[m.end() :]
            break
        inner, end = got
        body = body[: m.start()] + inner + body[end:]

    rows: List[List[List[Run]]] = []
    for part in re.split(r"\\\\", body):
        part = part.strip()
        if not part:
            continue
        # Drop trailing \\ / \tabularnewline leftovers already handled by split
        part = re.sub(r"\\tabularnewline\b", "", part).strip()
        if not part:
            continue
        cells_raw = [c.strip() for c in part.split("&")]
        if all(not c for c in cells_raw):
            continue
        joined = "".join(cells_raw)
        if re.fullmatch(r"[|lcrp@\{\}0-9.\\,\sX]+", joined) and "p{" in joined.replace(" ", ""):
            continue
        if re.fullmatch(r"[|lcrX\s]+", joined):
            continue
        row: List[List[Run]] = []
        for c in cells_raw:
            c = c.strip()
            if not c:
                row.append([Run("")])
                continue
            cell_runs = parse_inline(c, cite_map)
            if not cell_runs:
                row.append([Run("")])
            else:
                row.append(cell_runs)
        rows.append(row)
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
    return _bibitems_from_body(body)


def _bibitems_from_body(body: str):
    refs, cite_map = [], {}
    # Ignore trailing \end{...} wrappers from .bbl files
    body = re.sub(r"\\end\{[^}]+\}\s*$", "", body.strip())
    items = list(re.finditer(r"\\bibitem\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}", body))
    for idx, m in enumerate(items):
        key = m.group(1)
        start = m.end()
        end = items[idx + 1].start() if idx + 1 < len(items) else len(body)
        chunk = body[start:end]
        chunk = re.sub(r"\\(begin|end)\{[^}]+\}", "", chunk)
        text = re.sub(r"\s+", " ", runs_to_plain(parse_inline(chunk))).strip()
        # Drop biblatex field noise leftovers
        text = re.sub(r"\\newblock\b", " ", text)
        text = re.sub(r"\s+", " ", text).strip(" {}")
        num = idx + 1
        cite_map[key] = num
        refs.append(f"{num}. {text}")
    return refs, cite_map


def parse_bbl_file(bbl_path: Path):
    """Parse a BibTeX/biblatex .bbl file (\\bibitem entries)."""
    if not bbl_path.exists():
        return [], {}
    text = bbl_path.read_text(encoding="utf-8", errors="replace")
    return _bibitems_from_body(text)


def resolve_bibliography(tex: str, tex_path: Path):
    """Prefer thebibliography in .tex; else sibling .bbl (BibTeX / biblatex)."""
    refs, cite_map = parse_bibliography(tex)
    if refs:
        return refs, cite_map

    candidates = [tex_path.with_suffix(".bbl")]
    # Also try names from \\bibliography{a,b} → a.bbl (rare) and main stem
    m = re.search(r"\\bibliography\s*\{([^}]+)\}", tex)
    if m:
        for name in m.group(1).split(","):
            name = name.strip()
            if name:
                candidates.append(tex_path.parent / f"{name}.bbl")
    m = re.search(r"\\addbibresource\s*\{([^}]+)\}", tex)
    if m:
        # biblatex still writes <jobname>.bbl next to the .tex
        pass

    seen = set()
    for cand in candidates:
        cand = cand.resolve() if cand.exists() else cand
        key = str(cand)
        if key in seen:
            continue
        seen.add(key)
        if cand.exists():
            refs, cite_map = parse_bbl_file(cand)
            if refs:
                return refs, cite_map
    return [], {}



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




"""LaTeX math → OMML / plaintext."""
from __future__ import annotations

import re
from typing import Optional

from docx.oxml import parse_xml

from .constants import GREEK, MATH_OPS, OMML_NS, _HAS_OMML
from .authors import to_superscript
from .models import Equation, Run
from .util import _safe_balanced

try:
    import latex2mathml.converter as _latex2mathml
    from mathml2omml import convert as _mathml2omml
except ImportError:
    _latex2mathml = None  # type: ignore
    _mathml2omml = None  # type: ignore

def _prep_latex_math(latex: str) -> str:
    """Normalize LaTeX math for latex2mathml."""
    s = latex.strip()
    s = re.sub(r"\\label\{[^}]*\}", "", s)
    s = re.sub(r"\\nonumber\b", "", s)
    s = re.sub(r"\\tag\{[^}]*\}", "", s)
    # align* → aligned so \\ / & survive for latex2mathml (not flattened to spaces)
    s = re.sub(r"\\begin\{align\*?\}", r"\\begin{aligned}", s)
    s = re.sub(r"\\end\{align\*?\}", r"\\end{aligned}", s)
    s = re.sub(r"\\begin\{gather\*?\}", r"\\begin{gathered}", s)
    s = re.sub(r"\\end\{gather\*?\}", r"\\end{gathered}", s)
    # other display wrappers → plain math body
    s = re.sub(r"\\begin\{(?:multline|eqnarray)\*?\}", "", s)
    s = re.sub(r"\\end\{(?:multline|eqnarray)\*?\}", "", s)

    # Keep \\ and & inside cases/matrix/aligned — they are row/column separators.
    # Elsewhere \\ is a line break we flatten, and & is alignment noise.
    _KEEP = (
        "cases", "matrix", "pmatrix", "bmatrix", "vmatrix", "Vmatrix",
        "array", "smallmatrix", "aligned", "gathered",
    )

    def _protect(m: re.Match) -> str:
        return m.group(0).replace("\\\\", "\0R\0").replace("&", "\0A\0")

    for name in _KEEP:
        s = re.sub(
            rf"\\begin\{{{name}\*?\}}.*?\\end\{{{name}\*?\}}",
            _protect,
            s,
            flags=re.DOTALL,
        )

    s = s.replace("&=", "=").replace("&", " ")
    s = s.replace("\\\\", " ")
    s = s.replace("\0R\0", "\\\\").replace("\0A\0", "&")
    s = s.replace(r"\!", "")
    # keep \_ as-is for latex2mathml (plain _ inside \mathrm becomes a subscript)
    # function names: thin space before argument so Word doesn't glue ln+a_k
    s = re.sub(
        r"\\(ln|log|sin|cos|tan|exp|max|min|det|arg|ker|dim|deg)\b\s*",
        r"\\operatorname{\1}\\,",
        s,
    )
    s = re.sub(r"\s+", " ", s).strip()
    return s



def _fix_omml_radicals(omml: str) -> str:
    """Word shows a spurious root index unless degHide is set (pandoc does this)."""

    def repl(match: re.Match) -> str:
        inner = match.group(1)
        if "degHide" in inner:
            return match.group(0)
        deg = re.search(r"<m:deg>(.*?)</m:deg>", inner, flags=re.DOTALL)
        if deg and re.search(r"<m:t>[^<]+</m:t>", deg.group(1)):
            return match.group(0)
        inner = re.sub(r"<m:deg\s*/>|<m:deg>\s*</m:deg>", "", inner)
        return (
            "<m:rad>"
            "<m:radPr><m:degHide m:val=\"1\"/></m:radPr>"
            "<m:deg/>"
            f"{inner}"
            "</m:rad>"
        )

    return re.sub(r"<m:rad>(.*?)</m:rad>", repl, omml, flags=re.DOTALL)



def _unwrap_omml_boxes(omml: str) -> str:
    """mathml2omml wraps almost everything in m:box — Word then shifts baseline."""
    prev = None
    while prev != omml:
        prev = omml
        omml = re.sub(
            r"<m:box>\s*<m:e>(.*?)</m:e>\s*</m:box>",
            r"\1",
            omml,
            flags=re.DOTALL,
        )
    return omml



def latex_to_omml_element(latex: str, display: bool = False):
    """Convert LaTeX math to an lxml OMML element for python-docx, or None."""
    if not _HAS_OMML:
        return None
    src = _prep_latex_math(latex)
    if not src:
        return None
    try:
        mml = _latex2mathml.convert(src, display="block" if display else "inline")
        omml = _mathml2omml(mml)
    except Exception:
        return None
    if not omml or "<m:oMath" not in omml:
        return None
    omml = _fix_omml_radicals(omml)
    omml = _unwrap_omml_boxes(omml)
    if "xmlns:m=" not in omml[:120]:
        omml = omml.replace("<m:oMath>", f'<m:oMath xmlns:m="{OMML_NS}">', 1)
    try:
        if display:
            # Word display equation
            if omml.startswith("<m:oMath "):
                # drop xmlns from inner; put on wrapper
                omml_inner = re.sub(r'\sxmlns:m="[^"]*"', "", omml, count=1)
                xml = f'<m:oMathPara xmlns:m="{OMML_NS}">{omml_inner}</m:oMathPara>'
            elif omml.startswith("<m:oMath>"):
                xml = f'<m:oMathPara xmlns:m="{OMML_NS}">{omml}</m:oMathPara>'
            else:
                xml = f'<m:oMathPara xmlns:m="{OMML_NS}"><m:oMath>{omml}</m:oMath></m:oMathPara>'
            return parse_xml(xml)
        return parse_xml(omml)
    except Exception:
        return None



def _script_char(ch: str, kind: str) -> str:
    if kind == "sub":
        table = str.maketrans(
            "0123456789+-=()aehijklmnoprstuvx",
            "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ",
        )
    else:
        table = str.maketrans(
            "0123456789+-=()abdefghijklmnoprstuvwxyz",
            "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ᵃᵇᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ",
        )
    return ch.translate(table)



def latex_math_to_text(math: str) -> str:
    s = math.strip()
    for cmd in ("displaystyle", "textstyle", "scriptstyle", "left", "right", "big", "Big", "bigg", "Bigg"):
        s = re.sub(rf"\\{cmd}\b\s*", "", s)

    for cmd in ("text", "mathrm", "mathbf", "mathit", "operatorname", "mathsf", "texttt", "textrm", "textbf", "textit"):
        while True:
            m = re.search(rf"\\{cmd}\s*\{{", s)
            if not m:
                break
            got = _safe_balanced(s, m.end() - 1)
            if not got:
                s = s[: m.start()] + s[m.end() :]
                break
            inner, end = got
            s = s[: m.start()] + inner + s[end:]

    while True:
        m = re.search(r"\\frac\s*\{", s)
        if not m:
            break
        got = _safe_balanced(s, m.end() - 1)
        if not got:
            s = s[: m.start()] + s[m.end() :]
            break
        num, i1 = got
        while i1 < len(s) and s[i1].isspace():
            i1 += 1
        if i1 >= len(s) or s[i1] != "{":
            s = s[: m.start()] + f"({num})/" + s[i1:]
            continue
        got2 = _safe_balanced(s, i1)
        if not got2:
            s = s[: m.start()] + f"({num})/" + s[i1 + 1 :]
            continue
        den, i2 = got2
        s = s[: m.start()] + f"({latex_math_to_text(num)})/({latex_math_to_text(den)})" + s[i2:]

    while True:
        m = re.search(r"\\sqrt\s*\{", s)
        if not m:
            break
        got = _safe_balanced(s, m.end() - 1)
        if not got:
            s = s[: m.start()] + "√" + s[m.end() :]
            break
        inner, end = got
        s = s[: m.start()] + f"√({latex_math_to_text(inner)})" + s[end:]

    s = re.sub(r"\\begin\{cases\}", "{ ", s)
    s = re.sub(r"\\end\{cases\}", " }", s)
    s = s.replace("\\\\", "; ").replace("&", ", ")
    # Protect \_ so it stays a literal underscore (not a subscript trigger)
    s = s.replace(r"\_", "\0U\0")

    for name, sym in GREEK.items():
        # (?![A-Za-z]) — so \tau_k matches (\b fails before _)
        s = re.sub(rf"\\{name}(?![A-Za-z])", sym, s)
    for name, sym in MATH_OPS.items():
        s = re.sub(rf"\\{name}(?![A-Za-z])", sym, s)

    while True:
        m = re.search(r"\\overline\s*\{", s)
        if not m:
            break
        got = _safe_balanced(s, m.end() - 1)
        if not got:
            s = s[: m.start()] + s[m.end() :]
            break
        inner, end = got
        s = s[: m.start()] + latex_math_to_text(inner) + "̅" + s[end:]

    out, i = [], 0
    while i < len(s):
        if s[i] in "_^" and i + 1 < len(s):
            kind = "sub" if s[i] == "_" else "sup"
            i += 1
            if s[i] == "{":
                got = _safe_balanced(s, i)
                if not got:
                    out.append(("_" if kind == "sub" else "^"))
                    continue
                inner, i = got
                inner_t = latex_math_to_text(inner)
                mapped = "".join(_script_char(c, kind) for c in inner_t)
                fully = all(
                    (not c.isalnum()) or mapped[j] != c for j, c in enumerate(inner_t)
                )
                # Short indices → Unicode; multi-letter labels → comma/underscore
                # (t_k^{start} → tₖ,start; f_{ringdown,k} → f_ringdown,k)
                if fully and 0 < len(inner_t) <= 2 and "," not in inner_t:
                    out.append(mapped)
                elif kind == "sup" and re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", inner_t):
                    out.append("," + inner_t)
                else:
                    out.append(("_" if kind == "sub" else "^") + inner_t)
            else:
                ch = s[i]
                i += 1
                mapped = _script_char(ch, kind)
                out.append(mapped if mapped != ch else (("_" if kind == "sub" else "^") + ch))
            continue
        out.append(s[i])
        i += 1
    s = "".join(out)
    # Keep a space after stripped commands: \ln a → "ln a", not "lna"
    s = re.sub(r"\\([a-zA-Z]+)(\s*)", lambda m: m.group(1) + (" " if m.group(2) else ""), s)
    s = s.replace("{", "").replace("}", "").replace("~", " ")
    s = re.sub(r"\s+", " ", s).strip()
    # Glue fn-arg (ln a) so Word doesn't wrap between them; keep ", " breakable
    s = re.sub(
        r"\b(ln|log|sin|cos|tan|exp|max|min|det|mod)\s+",
        lambda m: m.group(1) + "\u00A0",
        s,
    )
    s = s.replace("\0U\0", "_")
    return s



def make_math_run(latex: str) -> Run:
    # Affiliation markers like ^{1} / ^{1,2} → unicode superscripts (not OMML)
    m = re.fullmatch(r"\^\{([0-9,\-\s]+)\}|\^([0-9])", latex.strip())
    if m:
        digits = (m.group(1) or m.group(2)).replace(" ", "")
        return Run(text=to_superscript(digits))
    plain = latex_math_to_text(latex)
    # Inline OMML like new_Q.docx (sub/sup, \mathrm) — not Unicode shortcuts
    return Run(text=plain, italic=True, math=latex)



def make_equation(latex: str, number: Optional[int] = None) -> Equation:
    return Equation(latex=latex, number=number, text=latex_math_to_text(latex))




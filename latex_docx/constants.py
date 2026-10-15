"""Shared constants for latex_docx."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

try:
    import latex2mathml.converter as _latex2mathml  # noqa: F401
    from mathml2omml import convert as _mathml2omml  # noqa: F401
    _HAS_OMML = True
except ImportError:
    _HAS_OMML = False

FONT = "Times New Roman"
BODY_PT = 14  # TD-J: Times New Roman, 14 pt
PAGE_W, PAGE_H = Inches(8.27), Inches(11.69)
MARGIN = Inches(0.79)  # ~2 cm
FIRST_INDENT = Inches(0.49)
FIG_WIDTH = Inches(6.2)
# Content width in twips (1 in = 1440 twips) for equation tab stops
_CONTENT_TWIPS = int(round((PAGE_W.inches - 2 * MARGIN.inches) * 1440))
_CENTER_TAB_TWIPS = _CONTENT_TWIPS // 2
OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
# Repo root (parent of the `latex_docx` package directory)
ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_OP_LOG = Path("out_latex_docx.log")
DEFAULT_DIFF_REPORT = Path("out_changes.md")
SUPERSCRIPTS = str.maketrans("0123456789+-=()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾")

# Front-matter layout templates
# tdj: bilingual journal header. plain and article: one front-matter block.
# article is the non-journal layout (pair it with lang="en" for English labels).
TEMPLATES = ("tdj", "plain", "article")


def resolve_authors_json(explicit: Optional[Path | str] = None) -> Path:
    """Pick authors.json: CLI path → cwd → repo root → authors.example.json."""
    if explicit:
        return Path(explicit)
    cwd = Path.cwd() / "authors.json"
    if cwd.exists():
        return cwd
    repo = ROOT_DIR / "authors.json"
    if repo.exists():
        return repo
    return ROOT_DIR / "authors.example.json"


# Back-compat alias (resolved at import; prefer resolve_authors_json() at call sites)
AUTHORS_JSON = resolve_authors_json()

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ",
    "phi": "φ", "varphi": "ϕ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Pi": "Π",
    "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

MATH_OPS = {
    "cdot": "·", "times": "×", "pm": "±", "leq": "≤", "geq": "≥", "neq": "≠",
    "approx": "≈", "propto": "∝", "infty": "∞", "partial": "∂", "sum": "∑",
    "prod": "∏", "int": "∫", "rightarrow": "→", "leftarrow": "←", "to": "→",
    "ldots": "…", "cdots": "⋯", "dots": "…", "circ": "∘", "mod": "mod",
    "bmod": "mod", "quad": " ", "qquad": "  ", "ell": "ℓ", "Re": "Re", "Im": "Im",
    "subseteq": "⊆", "in": "∈", "forall": "∀", "exists": "∃", "nabla": "∇",
}

def set_run_font(run, name: str = FONT, size_pt: Optional[float] = None) -> None:
    """Set font including w:rFonts ascii/hAnsi (needed on macOS Word)."""
    run.font.name = name
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)



def set_style_font(style, name: str = FONT) -> None:
    """Set style font + w:rFonts ascii/hAnsi."""
    style.font.name = name
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)




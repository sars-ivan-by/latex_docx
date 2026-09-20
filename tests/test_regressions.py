"""Regression tests for latex_docx (cases, \\_, eqref, align, sec refs)."""
from __future__ import annotations

from pathlib import Path

from docx.oxml.ns import qn

from latex_docx.authors import apply_authors_db, load_authors_db
from latex_docx.constants import FONT
from latex_docx.docx_builder import DocxBuilder
from latex_docx.inline import parse_inline
from latex_docx.math_omml import _prep_latex_math, latex_math_to_text
from latex_docx.models import Paper, Para
from latex_docx.parser import LatexParser

FIXTURES = Path(__file__).resolve().parent / "fixtures"
AUTHORS_DB = FIXTURES / "authors.json"


def test_prep_keeps_cases_row_separators():
    out = _prep_latex_math(r"\begin{cases} a & b \\ c & d \end{cases}")
    assert r"\begin{cases}" in out
    assert r"\\" in out
    assert "&" in out
    assert "\0R\0" not in out


def test_prep_align_becomes_aligned_not_flattened():
    out = _prep_latex_math(r"\begin{align} a &= b \\ c &= d \end{align}")
    assert r"\begin{aligned}" in out
    assert r"\\" in out
    assert "a = b c = d" not in out.replace(" ", "")


def test_mathrm_underscore_not_subscript():
    assert latex_math_to_text(r"\mathrm{exp\_fit}") == "exp_fit"
    assert r"\_" in _prep_latex_math(r"\mathrm{exp\_fit}") or "exp" in _prep_latex_math(
        r"\mathrm{exp\_fit}"
    )


def test_eqref_stays_separate_run():
    runs = parse_inline(r"see \eqref{eq:foo} and more")
    ref = next(r for r in runs if r.ref_key)
    assert ref.ref_key == "eq:foo"
    assert ref.eqref is True
    assert ref.text == "eq:foo"
    assert runs[0].text == "see "
    assert runs[-1].text == " and more"


def test_sec_and_eqref_resolve(tmp_path: Path):
    tex = tmp_path / "sample.tex"
    tex.write_text(
        r"""
\documentclass{article}
\begin{document}
\section{Quality}\label{sec:quality}
\subsection{Details}\label{sec:details}
See \ref{sec:quality} and \ref{sec:details}.
\begin{equation}\label{eq:one}
x=1
\end{equation}
See \eqref{eq:one}.
\end{document}
""",
        encoding="utf-8",
    )
    paper = LatexParser(tex).parse()
    texts = [
        "".join(r.text for r in b.runs) for b in paper.body if isinstance(b, Para)
    ]
    joined = " ".join(texts)
    assert "See 1 and 1.1." in joined
    assert "See (1)." in joined


def test_apply_authors_preserves_careful_tex():
    paper = Paper(
        authors_ru="И. А. Сукач¹",
        affiliations_ru="¹ Длинная организация из tex — более двадцати символов",
    )
    apply_authors_db(paper, load_authors_db(AUTHORS_DB))
    assert paper.authors_ru == "И. А. Сукач¹"
    assert paper.affiliations_ru.startswith("¹")
    assert paper.authors_en


def test_apply_authors_explicit_overwrites():
    paper = Paper(
        authors_ru="И. А. Сукач¹",
        affiliations_ru="¹ Длинная организация из tex — более двадцати символов",
    )
    apply_authors_db(
        paper, load_authors_db(AUTHORS_DB), author_ids=["sukach", "merkulov"]
    )
    assert "Меркулов" in paper.authors_ru
    assert "Sukach" in paper.authors_en or "I. A." in paper.authors_en


def test_run_font_sets_rFonts_ascii_hansi():
    builder = DocxBuilder()
    p = builder.add_p("hello")
    run = p.runs[0]
    r_pr = run._element.find(qn("w:rPr"))
    assert r_pr is not None
    r_fonts = r_pr.find(qn("w:rFonts"))
    assert r_fonts is not None
    assert r_fonts.get(qn("w:ascii")) == FONT
    assert r_fonts.get(qn("w:hAnsi")) == FONT

"""Document model dataclasses."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Union


@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    size: Optional[float] = None
    math: Optional[str] = None  # LaTeX → native Word OMML
    ref_key: Optional[str] = None  # unresolved \ref/\eqref label
    eqref: bool = False  # True → render as (N)


@dataclass
class Para:
    runs: List[Run] = field(default_factory=list)
    align: str = "justify"
    first_line: float = 0.49
    space_before: float = 0
    space_after: float = 0
    line_spacing: float = 1.5
    left_indent: float = 0
    hanging: float = 0
    size: Optional[float] = None


@dataclass
class Heading:
    text: str
    level: int = 1
    space_before: float = 12
    space_after: float = 6


@dataclass
class Equation:
    latex: str
    number: Optional[int] = None
    text: str = ""  # plaintext fallback


@dataclass
class Figure:
    paths: List[Path]
    caption: str = ""
    number: Optional[int] = None


@dataclass
class TableBlock:
    # Each cell is a list of Runs (may include OMML math)
    rows: List[List[List[Run]]]
    caption: str = ""
    number: Optional[int] = None


@dataclass
class ListBlock:
    items: List[List[Run]]
    ordered: bool = False


Block = Union[Para, Heading, Equation, Figure, TableBlock, ListBlock]


@dataclass
class Paper:
    udc: str = ""
    title_ru: str = ""
    title_en: str = ""
    authors_ru: str = ""
    authors_en: str = ""
    affiliations_ru: str = ""
    affiliations_en: str = ""
    abstract_ru: str = ""
    abstract_en: str = ""
    abstract_ru_tex: str = ""  # original with $...$ for OMML
    abstract_en_tex: str = ""
    keywords_ru: str = ""
    keywords_en: str = ""
    body: List[Block] = field(default_factory=list)
    references: List[str] = field(default_factory=list)
    authors_info: List[str] = field(default_factory=list)
    cite_map: dict = field(default_factory=dict)  # bib key → number for \cite

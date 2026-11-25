"""Intermediate model between the .docx reader and the .hwpx writer.

Lengths are kept in Word units (twips = 1/20 pt, EMU for drawings) so the reader stays a plain
reading of the file; the writer converts to HWPUNIT (1/100 pt = twips * 5).
Run and paragraph properties are already resolved (docDefaults -> styles -> direct formatting).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Union


@dataclass
class Text:
    text: str
    rpr: dict


@dataclass
class Field:
    kind: str  # "PAGE" | "NUMPAGES"
    rpr: dict
    fmt: str = "DIGIT"


@dataclass
class NoteRef:
    kind: str  # "footnote" | "endnote"
    blocks: list
    rpr: dict


@dataclass
class Hyperlink:
    url: str | None  # None: an internal anchor (kept as text)
    items: list


@dataclass
class Image:
    data: bytes
    fmt: str
    cx: int  # EMU
    cy: int  # EMU
    rpr: dict
    floating: bool = False


Inline = Union[Text, Field, NoteRef, Hyperlink, Image]


@dataclass
class Paragraph:
    ppr: dict
    mark_rpr: dict
    items: list = field(default_factory=list)
    style_id: str | None = None
    num: tuple | None = None  # (numId, ilvl)
    page_break: bool = False  # starts on a new page (hp:p pageBreak)
    column_break: bool = False
    after: list = field(default_factory=list)  # blocks pulled out of the paragraph (text box contents)

    def is_empty(self) -> bool:
        return not self.items and not self.after


@dataclass
class Cell:
    blocks: list
    span: int = 1
    vmerge: str | None = None  # "restart" | "continue" | None
    width: int | None = None  # twips (tcW dxa)
    borders: dict = field(default_factory=dict)  # side -> (val, sz, color) or None
    fill: str | None = None
    valign: str | None = None  # top|center|bottom
    margins: dict = field(default_factory=dict)  # side -> twips


@dataclass
class Row:
    cells: list
    height: int | None = None
    height_rule: str | None = None
    header: bool = False


@dataclass
class Table:
    grid: list  # column widths, twips
    rows: list
    width: tuple | None = None  # (w, type)
    jc: str | None = None
    indent: int = 0
    cell_margins: dict = field(default_factory=dict)  # side -> twips


@dataclass
class HeaderFooterSet:
    default: list | None = None
    first: list | None = None
    even: list | None = None


@dataclass
class Section:
    page_w: int = 11906
    page_h: int = 16838
    landscape: bool = False
    margins: dict = field(default_factory=lambda: {"top": 1440, "bottom": 1440, "left": 1440, "right": 1440,
                                                   "header": 720, "footer": 720, "gutter": 0})
    cols: int = 1
    col_space: int = 720
    col_sep: bool = False
    col_widths: list | None = None  # [(w, space)] when not equal
    start_type: str = "nextPage"
    title_pg: bool = False
    pg_start: int | None = None
    pg_fmt: str | None = None
    headers: HeaderFooterSet = field(default_factory=HeaderFooterSet)
    footers: HeaderFooterSet = field(default_factory=HeaderFooterSet)
    blocks: list = field(default_factory=list)


@dataclass
class Document:
    sections: list
    even_odd_headers: bool = False
    numbering: dict = field(default_factory=dict)  # numId -> {ilvl: level dict}

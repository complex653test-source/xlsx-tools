"""Immutable, openpyxl-free snapshots of a workbook.

This is the *escape hatch* surface: ``read()`` returns these objects, and
``ExtractResult.cells`` hands individual ``Cell`` objects back to callers who
need styling. Records produced by ``extract()`` never contain them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Mapping

from openpyxl.utils import get_column_letter, coordinate_to_tuple


@dataclass(frozen=True, slots=True)
class Color:
    rgb: str | None = None
    theme: int | None = None
    tint: float | None = None
    indexed: int | None = None


@dataclass(frozen=True, slots=True)
class Font:
    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strikethrough: bool = False
    color: Color | None = None


@dataclass(frozen=True, slots=True)
class Fill:
    pattern_type: str | None = None
    foreground: Color | None = None
    background: Color | None = None


@dataclass(frozen=True, slots=True)
class Side:
    style: str | None = None
    color: Color | None = None


@dataclass(frozen=True, slots=True)
class Border:
    left: Side = Side()
    right: Side = Side()
    top: Side = Side()
    bottom: Side = Side()
    diagonal: Side = Side()
    diagonal_direction: int | None = None


@dataclass(frozen=True, slots=True)
class Alignment:
    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool = False
    text_rotation: int | None = None
    indent: int | None = None
    shrink_to_fit: bool = False


@dataclass(frozen=True, slots=True)
class Protection:
    locked: bool = True
    hidden: bool = False


@dataclass(frozen=True, slots=True)
class Cell:
    """One addressable cell.

    ``value`` is what the spreadsheet *shows*: for a formula cell it is the
    cached result (``None`` when the file stored no cached result). ``formula``
    is the formula text, without a leading ``=`` only if openpyxl omitted it —
    in practice it includes the ``=``.
    """

    row: int
    column: int
    value: Any = None
    formula: str | None = None
    number_format: str = "General"
    font: Font = Font()
    fill: Fill = Fill()
    border: Border = Border()
    alignment: Alignment = Alignment()
    protection: Protection = Protection()
    from_merge: bool = False
    merge_anchor: str | None = None
    has_cached_value: bool = True

    @property
    def coordinate(self) -> str:
        return f"{get_column_letter(self.column)}{self.row}"

    def __str__(self) -> str:
        return self.coordinate


@dataclass(frozen=True, slots=True)
class Sheet:
    """A rectangular snapshot: ``rows[r][c]`` always exists for the sheet's range."""

    title: str
    rows: tuple[tuple[Cell, ...], ...] = ()
    merged_ranges: tuple[str, ...] = ()
    column_widths: Mapping[str, float] = field(default_factory=dict)
    row_heights: Mapping[int, float] = field(default_factory=dict)
    max_row: int = 0
    max_column: int = 0

    @property
    def width(self) -> int:
        return self.max_column

    @property
    def height(self) -> int:
        return self.max_row

    def cell_at(self, row: int, column: int) -> Cell:
        """Return the cell at 1-based ``row``/``column``.

        Coordinates outside the sheet range yield an empty ``Cell`` rather than
        raising: the sheet is rectangular, but callers (and schemas) may point
        beyond it.
        """
        if 1 <= row <= self.max_row and 1 <= column <= self.max_column:
            return self.rows[row - 1][column - 1]
        return Cell(row=row, column=column)

    def cell(self, coordinate: str) -> Cell:
        row, column = coordinate_to_tuple(coordinate)
        return self.cell_at(row, column)


@dataclass(frozen=True, slots=True)
class Workbook:
    sheets: tuple[Sheet, ...] = ()
    epoch: datetime = datetime(1899, 12, 30)
    active_index: int = 0
    source: str = "<memory>"

    @property
    def active(self) -> Sheet:
        return self.sheets[self.active_index]

    def sheet(self, title: str | None = None) -> Sheet:
        if title is None:
            return self.active
        for sheet in self.sheets:
            if sheet.title == title:
                return sheet
        raise KeyError(f"no sheet named {title!r}")

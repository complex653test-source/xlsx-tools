"""Parse the repository's `test.xlsx` fixture.

Run from the repo root:

    uv run python examples/parse_test_xlsx.py

The fixture is deliberately awkward: mixed types per column, a blank row in the
middle, a merged range, styled-but-empty cells, and a date stored as a serial.
"""

from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Annotated

from pydantic import BaseModel

from xlsx_tools import Column, extract, read

FIXTURE = Path(__file__).resolve().parents[1] / "test.xlsx"


# --- schemas -------------------------------------------------------------


class TypedRow(BaseModel):
    """Row 1: the only row where columns C/D/E hold typed values."""

    a: Annotated[str, Column("A")]
    c: Annotated[datetime, Column("C")]
    d: Annotated[int, Column("D")]
    e: Annotated[float, Column("E")]


class MoneyRow(BaseModel):
    """Row 5: a currency-formatted number."""

    d: Annotated[Decimal, Column("D")]


class LabelRow(BaseModel):
    """Row 6: labels, including a merged header spanning E-H."""

    a: Annotated[str, Column("A")]
    g: Annotated[str | None, Column("G")] = None
    j: Annotated[str | None, Column("J")] = None


class StrictMixed(BaseModel):
    """Deliberately too strict for this sheet, to show lenient mode."""

    a: Annotated[str | None, Column("A")] = None
    c: Annotated[datetime | None, Column("C")] = None
    d: Annotated[int | None, Column("D")] = None


def show_cell_model() -> None:
    workbook = read(FIXTURE)
    sheet = workbook.active

    print("== cell model ==")
    print(f"sheet={sheet.title!r} epoch={workbook.epoch.date()} "
          f"used_range=A1:{sheet.cell_at(sheet.max_row, sheet.max_column).coordinate} "
          f"({sheet.height}x{sheet.width})")
    print(f"merged ranges: {list(sheet.merged_ranges)}")
    print(f"row heights: {dict(sheet.row_heights)}")
    print()

    print("non-empty cells:")
    for row in sheet.rows:
        for cell in row:
            if cell.value is None:
                continue
            style = []
            if cell.font.bold:
                style.append("bold")
            if cell.font.color and cell.font.color.rgb:
                style.append(f"font={cell.font.color.rgb}")
            if cell.fill.pattern_type and cell.fill.foreground and cell.fill.foreground.rgb:
                style.append(f"fill={cell.fill.foreground.rgb}")
            if cell.font.size != 11:
                style.append(f"size={cell.font.size:g}")
            note = " ".join(style)
            print(
                f"  {cell.coordinate:<4} {cell.value!r:<34} "
                f"{cell.number_format:<44} {note}"
            )
    print()

    anchor = sheet.cell("E6")
    covered = sheet.cell("G6")
    print(f"merge: anchor {anchor.coordinate}={anchor.value!r}; "
          f"{covered.coordinate}={covered.value!r} "
          f"from_merge={covered.from_merge} anchor={covered.merge_anchor}")
    print()


def show_extraction() -> None:
    print("== extract: row 1 (typed) ==")
    typed = extract(FIXTURE, TypedRow, first_row=1, last_row=1)
    print(f"  {typed.records[0]}")
    print(f"  cell model for c: {typed.cells[0]['c'].coordinate} "
          f"{typed.cells[0]['c'].number_format}")
    print()

    print("== extract: row 5 (Decimal) ==")
    money = extract(FIXTURE, MoneyRow, first_row=5, last_row=5)
    print(f"  {money.records[0]}  ({type(money.records[0].d).__name__})")
    print()

    print("== extract: row 6 (labels + merged cell) ==")
    labels = extract(FIXTURE, LabelRow, first_row=6, last_row=6)
    print(f"  {labels.records[0]}")
    print(f"  G6 came from merge anchor {labels.cells[0]['g'].merge_anchor}")
    print()

    print("== extract: whole sheet, lenient ==")
    result = extract(FIXTURE, StrictMixed, strict=False)
    print(f"  records={len(result.records)} issues={len(result.issues)}")
    for issue in result.issues[:4]:
        print(f"    {issue}")
    print()


if __name__ == "__main__":
    show_cell_model()
    show_extraction()

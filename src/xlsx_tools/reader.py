"""Loading a workbook into immutable snapshots."""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .errors import XlsxToolsError
from .model import (
    Alignment,
    Border,
    Cell,
    Color,
    Fill,
    Font,
    Protection,
    Sheet,
    Side,
    Workbook,
)

def _payload(source: Any) -> tuple[Any, str]:
    """Normalise any accepted source to a path or raw bytes.

    Bytes are read up front because a workbook may be parsed twice (formulas and
    cached values); re-reading a caller-owned stream twice is not safe.
    """
    if isinstance(source, (str, Path)):
        return str(source), str(source)
    if isinstance(source, bytes):
        return source, "<bytes>"
    if hasattr(source, "read"):
        try:
            source.seek(0)
        except (AttributeError, OSError):
            pass
        return source.read(), getattr(source, "name", "<stream>")
    raise TypeError(f"unsupported source {type(source).__name__}; expected path, bytes or file-like")


def _load(payload: Any, *, data_only: bool):
    handle = io.BytesIO(payload) if isinstance(payload, bytes) else payload
    return load_workbook(handle, read_only=False, data_only=data_only)


def _rgb(color: Any) -> str | None:
    rgb = getattr(color, "rgb", None)
    return rgb if isinstance(rgb, str) else None


def _color(color: Any) -> Color | None:
    if color is None:
        return None
    theme = getattr(color, "theme", None)
    tint = getattr(color, "tint", None)
    indexed = getattr(color, "indexed", None)
    return Color(
        rgb=_rgb(color),
        theme=theme if isinstance(theme, int) else None,
        tint=tint if isinstance(tint, (int, float)) else None,
        indexed=indexed if isinstance(indexed, int) else None,
    )


def _side(side: Any) -> Side:
    if side is None:
        return Side()
    return Side(style=getattr(side, "style", None), color=_color(getattr(side, "color", None)))


def _font(font: Any) -> Font:
    return Font(
        name=getattr(font, "name", None),
        size=getattr(font, "sz", None),
        bold=bool(getattr(font, "b", False)),
        italic=bool(getattr(font, "i", False)),
        underline=getattr(font, "u", None),
        strikethrough=bool(getattr(font, "strike", False)),
        color=_color(getattr(font, "color", None)),
    )


def _fill(fill: Any) -> Fill:
    return Fill(
        pattern_type=getattr(fill, "patternType", None),
        foreground=_color(getattr(fill, "fgColor", None)),
        background=_color(getattr(fill, "bgColor", None)),
    )


def _border(border: Any) -> Border:
    return Border(
        left=_side(getattr(border, "left", None)),
        right=_side(getattr(border, "right", None)),
        top=_side(getattr(border, "top", None)),
        bottom=_side(getattr(border, "bottom", None)),
        diagonal=_side(getattr(border, "diagonal", None)),
        diagonal_direction=getattr(border, "diagonal_direction", None),
    )


def _alignment(alignment: Any) -> Alignment:
    return Alignment(
        horizontal=getattr(alignment, "horizontal", None),
        vertical=getattr(alignment, "vertical", None),
        wrap_text=bool(getattr(alignment, "wrap_text", False)),
        text_rotation=getattr(alignment, "text_rotation", None),
        indent=getattr(alignment, "indent", None),
        shrink_to_fit=bool(getattr(alignment, "shrink_to_fit", False)),
    )


def _protection(protection: Any) -> Protection:
    return Protection(
        locked=bool(getattr(protection, "locked", True)),
        hidden=bool(getattr(protection, "hidden", False)),
    )


def _formula_caches(payload: Any) -> dict[tuple[str, int, int], Any]:
    """Second pass: cached results for formula cells, keyed by (sheet, row, col)."""
    workbook = _load(payload, data_only=True)
    caches: dict[tuple[str, int, int], Any] = {}
    try:
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f" or cell.value is not None:
                        caches[(worksheet.title, cell.row, cell.column)] = cell.value
    finally:
        workbook.close()
    return caches


def _merged_map(
    worksheet: Worksheet,
) -> dict[tuple[int, int], tuple[str, int, int]]:
    """Map every covered coordinate to its anchor's (coordinate, row, column)."""
    covered: dict[tuple[int, int], tuple[str, int, int]] = {}
    for merged in worksheet.merged_cells.ranges:
        anchor_coord = f"{get_column_letter(merged.min_col)}{merged.min_row}"
        for row in range(merged.min_row, merged.max_row + 1):
            for column in range(merged.min_col, merged.max_col + 1):
                if (row, column) == (merged.min_row, merged.min_col):
                    continue
                covered[(row, column)] = (anchor_coord, merged.min_row, merged.min_col)
    return covered


def _sheet(
    worksheet: Worksheet,
    caches: dict[tuple[str, int, int], Any],
) -> Sheet:
    merged = _merged_map(worksheet)
    max_row = max(worksheet.max_row, 1)
    max_column = max(worksheet.max_column, 1)

    rows: list[tuple[Cell, ...]] = []
    for row_cells in worksheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_column):
        cells: list[Cell] = []
        for raw in row_cells:
            key = (worksheet.title, raw.row, raw.column)
            value = raw.value
            formula = raw.value if raw.data_type == "f" else None
            has_cached_value = True
            if formula is not None:
                if key in caches:
                    value = caches[key]
                else:
                    value, has_cached_value = None, False

            from_merge = False
            merge_anchor = None
            if (raw.row, raw.column) in merged:
                merge_anchor, anchor_row, anchor_column = merged[(raw.row, raw.column)]
                anchor_key = (worksheet.title, anchor_row, anchor_column)
                anchor = worksheet.cell(row=anchor_row, column=anchor_column)
                anchor_value = anchor.value
                anchor_formula = anchor.value if anchor.data_type == "f" else None
                if anchor_formula is not None:
                    formula = anchor_formula
                    has_cached_value = anchor_key in caches
                    value = caches.get(anchor_key)
                else:
                    value = anchor_value
                    has_cached_value = True
                from_merge = True

            cells.append(
                Cell(
                    row=raw.row,
                    column=raw.column,
                    value=value,
                    formula=formula,
                    number_format=raw.number_format or "General",
                    font=_font(raw.font),
                    fill=_fill(raw.fill),
                    border=_border(raw.border),
                    alignment=_alignment(raw.alignment),
                    protection=_protection(raw.protection),
                    from_merge=from_merge,
                    merge_anchor=merge_anchor,
                    has_cached_value=has_cached_value,
                )
            )
        rows.append(tuple(cells))

    widths = {
        str(key): dimension.width
        for key, dimension in worksheet.column_dimensions.items()
        if dimension.width is not None
    }
    heights = {
        int(key): dimension.height
        for key, dimension in worksheet.row_dimensions.items()
        if dimension.height is not None
    }

    return Sheet(
        title=worksheet.title,
        rows=tuple(rows),
        merged_ranges=tuple(str(r) for r in worksheet.merged_cells.ranges),
        column_widths=widths,
        row_heights=heights,
        max_row=max_row,
        max_column=max_column,
    )


def read(source: Any) -> Workbook:
    """Read a workbook into immutable snapshots.

    Read-only in the sense that matters: the input is never written to. The
    loader itself is eager (``read_only=False``) because streaming mode cannot
    expose merged ranges, column widths or row heights.
    """
    payload, label = _payload(source)

    workbook = _load(payload, data_only=False)
    try:
        has_formulas = any(
            cell.data_type == "f"
            for worksheet in workbook.worksheets
            for row in worksheet.iter_rows()
            for cell in row
        )
        caches = _formula_caches(payload) if has_formulas else {}
        sheets = tuple(_sheet(worksheet, caches) for worksheet in workbook.worksheets)
        active_index = workbook.index(workbook.active)
        epoch = workbook.epoch
    finally:
        workbook.close()

    if not sheets:
        raise XlsxToolsError(f"workbook {label} contains no worksheets")

    return Workbook(sheets=sheets, epoch=epoch, active_index=active_index, source=label)

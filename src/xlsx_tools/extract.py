"""Schema-driven extraction of worksheet rows into pydantic records."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping

from pydantic import BaseModel, ValidationError

from .errors import ExtractionError, FieldError, FormulaNotCachedError, SchemaError
from .model import Cell, Sheet
from .reader import read
from .schema import FieldSpec, field_specs


@dataclass(frozen=True, slots=True)
class ExtractResult:
    """Records, their source cells (index-aligned), and any issues found."""

    records: tuple[BaseModel, ...]
    cells: tuple[Mapping[str, Cell], ...]
    issues: tuple[FieldError, ...]
    sheet: str


def extract(
    source: Any,
    model_cls: type[BaseModel],
    *,
    sheet: str | None = None,
    strict: bool = True,
    header_row: int | None = None,
    first_row: int = 1,
    last_row: int | None = None,
) -> ExtractResult:
    """Extract rows of one sheet into ``model_cls`` instances.

    - Columns come from ``Column(...)`` markers on the model's fields (required
      for every field, including ones with defaults).
    - Rows run from ``first_row`` to ``last_row``. When ``last_row`` is omitted
      it is the last row where *every* declared field is non-empty; if no such
      row exists, ``ExtractionError`` is raised (an empty range is a broken
      schema or wrong sheet, not an empty dataset).
    - Fully blank rows inside the range are skipped.
    - Merged ranges propagate the anchor value into covered cells.
    - ``strict`` raises one ``ExtractionError`` carrying every issue; otherwise
      the bad rows are omitted and returned in ``issues``.
    """
    specs = field_specs(model_cls)

    if first_row < 1:
        raise SchemaError(f"first_row must be >= 1, got {first_row}")
    if last_row is not None and last_row < first_row:
        raise ExtractionError(f"last_row ({last_row}) is before first_row ({first_row})")
    if header_row is not None and header_row < 1:
        raise SchemaError(f"header_row must be >= 1, got {header_row}")
    if header_row is None and any(spec.header is not None for spec in specs):
        raise SchemaError(
            "a field declares header= but extract() was called without header_row="
        )

    workbook = read(source)
    worksheet = workbook.sheet(sheet)

    if first_row > worksheet.max_row:
        raise ExtractionError(
            f"no data rows in {worksheet.title!r}: first_row {first_row} is past max_row {worksheet.max_row}"
        )

    if header_row is not None:
        _assert_headers(worksheet, specs, header_row)

    resolved_last = last_row if last_row is not None else _resolve_last_row(worksheet, specs, first_row)

    records: list[BaseModel] = []
    cells: list[Mapping[str, Cell]] = []
    issues: list[FieldError] = []
    saw_data = False

    for row in range(first_row, resolved_last + 1):
        row_cells = {spec.name: worksheet.cell_at(row, spec.index) for spec in specs}
        if not any(cell.value is not None or cell.formula is not None for cell in row_cells.values()):
            continue
        saw_data = True

        row_issues: list[FieldError] = []
        for spec in specs:
            cell = row_cells[spec.name]
            if cell.formula is not None and not cell.has_cached_value:
                row_issues.append(
                    FormulaNotCachedError(
                        sheet=worksheet.title,
                        cell=cell.coordinate,
                        field=spec.name,
                        reason=(
                            "formula has no cached value; recalculate and save it in a "
                            "spreadsheet application"
                        ),
                    )
                )
        if row_issues:
            issues.extend(row_issues)
            continue

        values = {name: cell.value for name, cell in row_cells.items()}
        try:
            record = model_cls(**values)
        except ValidationError as exc:
            for error in exc.errors():
                location = error.get("loc") or ()
                field = str(location[0]) if location else "<model>"
                cell = row_cells.get(field)
                issues.append(
                    FieldError(
                        sheet=worksheet.title,
                        cell=cell.coordinate if cell is not None else f"row {row}",
                        field=field,
                        reason=str(error.get("msg", "invalid value")),
                        value=values.get(field),
                    )
                )
            continue

        records.append(record)
        cells.append(dict(row_cells))

    if not saw_data:
        raise ExtractionError(
            f"no data rows in {worksheet.title!r} between rows {first_row} and {resolved_last} "
            f"for the declared columns"
        )

    if strict and issues:
        raise ExtractionError(
            f"{len(issues)} issue(s) extracting {model_cls.__name__} from sheet "
            f"{worksheet.title!r}",
            tuple(issues),
        )

    return ExtractResult(
        records=tuple(records),
        cells=tuple(cells),
        issues=tuple(issues),
        sheet=worksheet.title,
    )


def _resolve_last_row(worksheet: Sheet, specs: tuple[FieldSpec, ...], first_row: int) -> int:
    for row in range(worksheet.max_row, first_row - 1, -1):
        if all(worksheet.cell_at(row, spec.index).value is not None for spec in specs):
            return row
    raise ExtractionError(
        f"no data rows in {worksheet.title!r}: no row from {first_row} onward has every "
        "declared field populated (pass last_row= explicitly to override)"
    )


def _assert_headers(worksheet: Sheet, specs: tuple[FieldSpec, ...], header_row: int) -> None:
    if header_row > worksheet.max_row:
        raise SchemaError(
            f"header_row {header_row} is past max_row {worksheet.max_row} of {worksheet.title!r}"
        )
    for spec in specs:
        if spec.header is None:
            continue
        cell = worksheet.cell_at(header_row, spec.index)
        actual = "" if cell.value is None else str(cell.value)
        if actual.strip() != spec.header:
            raise SchemaError(
                f"header mismatch at {worksheet.title}!{cell.coordinate} for field "
                f"{spec.name!r}: expected {spec.header!r}, found {actual!r}"
            )

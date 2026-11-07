from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from typing import Annotated

import pytest
from pydantic import BaseModel

from xlsx_tools import (
    Column,
    ExtractionError,
    FieldError,
    FormulaNotCachedError,
    SchemaError,
    extract,
    read,
)


class RowA(BaseModel):
    a: Annotated[str | None, Column("A")] = None


class TypedRow(BaseModel):
    a: Annotated[str, Column("A")]
    c: Annotated[datetime, Column("C")]
    d: Annotated[int, Column("D")]
    e: Annotated[float, Column("E")]


class MoneyRow(BaseModel):
    d: Annotated[Decimal, Column("D")]


class MergedRow(BaseModel):
    a: Annotated[str | None, Column("A")] = None
    g: Annotated[str | None, Column("G")] = None


class RequiredD(BaseModel):
    d: Annotated[int, Column("D")]


class HeaderRow(BaseModel):
    a: Annotated[str, Column("A", header="a")]
    b: Annotated[str, Column("B", header="b")]


class FormulaRow(BaseModel):
    a: Annotated[float | None, Column("A")] = None


def test_extracts_typed_records(fixture_path):
    result = extract(fixture_path, TypedRow, first_row=1, last_row=1)

    assert result.sheet == "Sheet1"
    assert result.issues == ()
    record = result.records[0]
    assert record.a == "a"
    assert record.c == datetime(2000, 1, 1)
    assert record.d == 98
    assert record.e == 123.456


def test_blank_row_is_skipped(fixture_path):
    result = extract(fixture_path, RowA)

    assert [record.a for record in result.records] == [
        "a",
        "b",
        "blue back",
        "another",
        "A ",
    ]


def test_records_have_parallel_source_cells(fixture_path):
    result = extract(fixture_path, TypedRow, first_row=1, last_row=1)

    assert result.cells[0]["a"].coordinate == "A1"
    assert result.cells[0]["c"].number_format == "mm-dd-yy"
    assert result.cells[0]["a"] == read(fixture_path).active.cell("A1")


def test_merged_cells_propagate_to_fields(fixture_path):
    result = extract(fixture_path, MergedRow)

    assert len(result.records) == 5  # last_row is bounded by row 6, partial rows are kept
    assert [record.g for record in result.records] == [None, None, None, None, "merged column e-h"]
    assert result.cells[-1]["g"].from_merge is True
    assert result.cells[-1]["g"].merge_anchor == "E6"


def test_header_assertions_pass(fixture_path):
    result = extract(fixture_path, HeaderRow, header_row=1, last_row=1)

    assert result.records[0].a == "a"
    assert result.records[0].b == "b"


def test_header_mismatch_raises_schema_error(fixture_path):
    class WrongHeader(BaseModel):
        a: Annotated[str, Column("A", header="z")]

    with pytest.raises(SchemaError, match="header mismatch"):
        extract(fixture_path, WrongHeader, header_row=1, last_row=1)


def test_header_without_header_row_raises_schema_error(fixture_path):
    with pytest.raises(SchemaError, match="header_row"):
        extract(fixture_path, HeaderRow)


def test_field_without_column_raises_schema_error(fixture_path):
    class MissingColumn(BaseModel):
        a: str = "default"

    with pytest.raises(SchemaError, match="no Column"):
        extract(fixture_path, MissingColumn)


def test_unknown_column_letter_raises_schema_error():
    with pytest.raises(SchemaError, match="invalid column"):
        Column("nope")


def test_strict_mode_raises_aggregate_error(fixture_path):
    with pytest.raises(ExtractionError) as excinfo:
        extract(fixture_path, RequiredD)

    issues = excinfo.value.issues
    assert len(issues) == 2
    assert {issue.cell for issue in issues} == {"D5", "D6"}
    assert all(isinstance(issue, FieldError) for issue in issues)
    assert {issue.field for issue in issues} == {"d"}
    assert {issue.sheet for issue in issues} == {"Sheet1"}


def test_lenient_mode_returns_issues(fixture_path):
    result = extract(fixture_path, RequiredD, strict=False)

    assert [record.d for record in result.records] == [98]
    assert {issue.cell for issue in result.issues} == {"D5", "D6"}
    assert result.issues[0].value == 999.9


def test_decimal_keeps_precision(fixture_path):
    result = extract(fixture_path, MoneyRow, first_row=5, last_row=5)

    assert result.records[0].d == Decimal("999.9")


def test_empty_range_raises(fixture_path):
    with pytest.raises(ExtractionError, match="no data rows"):
        extract(fixture_path, RowA, first_row=4, last_row=4)


def test_last_row_before_first_row_raises(fixture_path):
    with pytest.raises(ExtractionError, match="before first_row"):
        extract(fixture_path, RowA, first_row=5, last_row=4)


def test_last_row_is_conjunctive(ragged_workbook):
    class Ragged(BaseModel):
        a: Annotated[str, Column("A")]
        b: Annotated[int, Column("B")]

    result = extract(ragged_workbook, Ragged)

    assert len(result.records) == 1
    assert result.records[0].a == "first"
    assert result.records[0].b == 1


def test_first_row_is_respected(fixture_path):
    result = extract(fixture_path, RowA, first_row=3)

    assert [record.a for record in result.records] == ["blue back", "another", "A "]


def test_sheet_can_be_named(fixture_path):
    assert extract(fixture_path, RowA, sheet="Sheet1").sheet == "Sheet1"


def test_unknown_sheet_raises_key_error(fixture_path):
    with pytest.raises(KeyError, match="nope"):
        extract(fixture_path, RowA, sheet="nope")


def test_accepts_bytes(fixture_path):
    result = extract(fixture_path.read_bytes(), RowA)

    assert result.records[0].a == "a"


def test_formula_without_cache_raises_in_strict_mode(formula_workbook):
    with pytest.raises(ExtractionError) as excinfo:
        extract(formula_workbook, FormulaRow, sheet="S", first_row=2, last_row=2)

    assert isinstance(excinfo.value.issues[0], FormulaNotCachedError)


def test_formula_without_cache_reports_issue_in_lenient_mode(formula_workbook):
    result = extract(
        formula_workbook, FormulaRow, sheet="S", first_row=2, last_row=2, strict=False
    )

    assert result.records == ()
    assert isinstance(result.issues[0], FormulaNotCachedError)
    assert result.issues[0].cell == "A2"


def test_formula_cached_value_is_extracted(formula_workbook_cached):
    result = extract(
        formula_workbook_cached, FormulaRow, sheet="S", first_row=2, last_row=2
    )

    assert result.records[0].a == 2.0

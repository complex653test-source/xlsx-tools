from __future__ import annotations

from datetime import datetime

from xlsx_tools import read


def test_sheet_is_rectangular(fixture_path):
    sheet = read(fixture_path).active

    assert sheet.title == "Sheet1"
    assert (sheet.max_row, sheet.max_column) == (6, 10)
    assert sheet.height == 6
    assert sheet.width == 10
    assert len(sheet.rows) == 6
    assert all(len(row) == 10 for row in sheet.rows)


def test_typed_values_and_dates(fixture_path):
    sheet = read(fixture_path).active

    assert sheet.cell("A1").value == "a"
    assert sheet.cell("C1").value == datetime(2000, 1, 1)
    assert sheet.cell("D1").value == 98
    assert sheet.cell("E1").value == 123.456


def test_number_formats_are_preserved(fixture_path):
    sheet = read(fixture_path).active

    assert sheet.cell("C1").number_format == "mm-dd-yy"
    assert "#,##0.00" in sheet.cell("D5").number_format


def test_styles_are_snapshotted(fixture_path):
    sheet = read(fixture_path).active

    assert sheet.cell("A3").fill.foreground.rgb == "FF2D7FCF"
    assert sheet.cell("B3").font.color.rgb == "FFC00000"
    assert sheet.cell("C3").font.bold is True
    assert sheet.cell("C3").font.color.rgb == "FFC00000"
    assert sheet.cell("F5").font.size == 28


def test_merged_range_propagates_anchor_value(fixture_path):
    sheet = read(fixture_path).active

    assert sheet.merged_ranges == ("E6:H6",)
    for coordinate in ("F6", "G6", "H6"):
        cell = sheet.cell(coordinate)
        assert cell.value == "merged column e-h"
        assert cell.from_merge is True
        assert cell.merge_anchor == "E6"
    assert sheet.cell("E6").from_merge is False


def test_row_heights_are_preserved(fixture_path):
    assert read(fixture_path).active.row_heights[5] == 38.18


def test_gaps_are_none_not_missing(fixture_path):
    sheet = read(fixture_path).active

    assert sheet.cell("B2").value is None
    assert sheet.cell("G5").value is None
    assert sheet.cell("A4").value is None


def test_epoch_is_exposed(fixture_path):
    assert read(fixture_path).epoch == datetime(1899, 12, 30)


def test_1904_date_system_uses_its_epoch(date1904_workbook):
    workbook = read(date1904_workbook)

    assert workbook.epoch == datetime(1904, 1, 1)
    assert workbook.active.cell("C1").value == datetime(2004, 1, 2)


def test_reads_bytes(fixture_path):
    workbook = read(fixture_path.read_bytes())

    assert workbook.active.cell("A1").value == "a"


def test_formula_without_cached_value(formula_workbook):
    cell = read(formula_workbook).sheet("S").cell("A2")

    assert cell.formula == "=A1*2"
    assert cell.value is None
    assert cell.has_cached_value is False


def test_formula_with_cached_value(formula_workbook_cached):
    cell = read(formula_workbook_cached).sheet("S").cell("A2")

    assert cell.formula == "=A1*2"
    assert cell.value == 2
    assert cell.has_cached_value is True

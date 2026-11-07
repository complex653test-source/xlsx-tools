from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "test.xlsx"


@pytest.fixture(scope="session")
def fixture_path() -> Path:
    assert FIXTURE.exists(), f"missing fixture {FIXTURE}"
    return FIXTURE


def _add_cached_value(path: Path, coordinate: str, value: int) -> None:
    """Inject a cached formula result, which openpyxl itself never writes."""
    patched = path.with_suffix(".patched.xlsx")
    pattern = re.compile(
        r'(<c r="%s"[^>]*>)\s*(<f>[^<]*</f>)\s*(?:<v\s*/>|<v>[^<]*</v>)?\s*</c>'
        % re.escape(coordinate)
    )
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        patched, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                text, count = pattern.subn(r"\1\2<v>%d</v></c>" % value, text)
                assert count == 1, f"could not patch formula at {coordinate}:\n{text}"
                data = text.encode()
            target.writestr(item, data)
    shutil.move(str(patched), str(path))


@pytest.fixture
def formula_workbook(tmp_path: Path) -> Path:
    """A workbook whose formula cell has *no* cached result (the common case)."""
    path = tmp_path / "formula.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "S"
    sheet["A1"] = 1
    sheet["A2"] = "=A1*2"
    workbook.save(path)
    return path


@pytest.fixture
def formula_workbook_cached(tmp_path: Path) -> Path:
    """Same workbook, with a cached result injected into the XML."""
    path = tmp_path / "formula_cached.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "S"
    sheet["A1"] = 1
    sheet["A2"] = "=A1*2"
    workbook.save(path)
    _add_cached_value(path, "A2", 2)
    return path


@pytest.fixture
def date1904_workbook(tmp_path: Path, fixture_path: Path) -> Path:
    """The fixture with the 1904 date system enabled, shifting every serial date."""
    path = tmp_path / "date1904.xlsx"
    with zipfile.ZipFile(fixture_path) as source, zipfile.ZipFile(
        path, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.decode().replace(
                    "<workbookPr", '<workbookPr date1904="1"'
                ).encode()
            target.writestr(item, data)
    return path


@pytest.fixture
def ragged_workbook(tmp_path: Path) -> Path:
    """Two rows where the last row is missing a declared field (truncation case)."""
    path = tmp_path / "ragged.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "S"
    sheet["A1"], sheet["B1"] = "first", 1
    sheet["A2"] = "second"  # B2 deliberately empty
    workbook.save(path)
    return path

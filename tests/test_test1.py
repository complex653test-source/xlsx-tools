from __future__ import annotations

from datetime import datetime

from xlsx_tools.workbook import XlsxWorkbook
from pytest import raises

def test_string_extractions(test1_xlsx_path):
    wb = XlsxWorkbook.load(test1_xlsx_path)
    sheet=wb.active
    print(sheet.get_cell("A1").value)    #sheet = read(test1_xlsx_path).active
    print(sheet.get_cell("AG4").value)    #sheet = read(test1_xlsx_path).active
    assert isinstance(sheet.get_cell("AG4").value, datetime)
    with raises(ValueError):
        sheet.get_cell("AG4").str()  # AG4 contains a datetime

    
    with raises(ValueError):
        sheet.get_cell("AA4").str()
        

def test_date_extractions(test1_xlsx_path):
    wb = XlsxWorkbook.load(test1_xlsx_path)
    sheet=wb.active
    print(sheet.get_cell("A1").value)    #sheet = read(test1_xlsx_path).active
    print(sheet.get_cell("AG4").value)    #sheet = read(test1_xlsx_path).active
    assert isinstance(sheet.get_cell("AG4").value, datetime)
    assert sheet.get_cell("AG4").datetime().day==19
    assert sheet.get_cell("AG4").datetime().year==2026
    assert sheet.get_cell("AG4").datetime().month==11

    
    with raises(ValueError):
        sheet.get_cell("AA4").datetime()
        
    with raises(ValueError):
        sheet.get_cell("AA4").datetime(default="AAAA")

    with raises(IndexError):
        sheet.get_cell("AG5").datetime()





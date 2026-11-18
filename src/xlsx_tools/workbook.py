from abc import ABC
from abc import abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Dict
from typing import List

from openpyxl import load_workbook
from openpyxl import workbook
from openpyxl import worksheet

from xlsx_tools.dict_tools import DictParser
from xlsx_tools.worksheet import XlsxWorksheet


@dataclass(frozen=True)
class XlsxWorkbook:
    workbook:workbook
    user_data:Dict
    @staticmethod
    def load(source:Path, user_data:dict={})->XlsxWorkbook:
        if isinstance(source, Path):
            return XlsxWorkbook(workbook=load_workbook(source), user_data=user_data)
        assert False

    @property
    def active(self)->worksheet:
        return XlsxWorksheet(self.workbook.active)
    

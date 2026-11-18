import json
import logging
from abc import ABC
from abc import abstractmethod
from dataclasses import dataclass
from datetime import datetime
from typing import Dict

from openpyxl.cell.cell import Cell

from xlsx_tools.utils import INVALID, col_to_int
from xlsx_tools.dict_tools import DictParser

from xlsx_tools.cell import XlsxCell

from xlsx_tools.utils import ChainAssert


from typing import Tuple
    
@dataclass(frozen=True)
class XlsxRow:
    row_number:int
    user_data:Dict
    cells:Tuple[XlsxCell] # tuple of XlsxCell values


    def __post_init__(self):
        for i in self.cells:
            print(type(i))
            assert isinstance(i, XlsxCell)
            # ChainAssert(value=i, name="").isinstance(XlsxCell)

    def __len__(self)->int:
        return len(self.cells)

    # def xlsx_cells(self)->list[XlsxCell]:
    #     for i in self.cells:
    #         assert isinstance(i, XlsxCell)
    #     # for n, i in 
    #     return self.cells
        
    def get_raw(self, col:str|int)->XlsxCell:
        if isinstance(col, int):
            return XlsxCell(raw= self.cells[col])
        if isinstance(col, str):
            col=col.strip()
            assert col != ""
            col=col.upper()
            col_int=0

            temp=col
            def _char_offset(c)->int:
                c=c.upper()
                r = ord(c)-ord("A")
                assert r >= 0
                assert r < 26
                return r

            while len(temp) > 1:
                col_int += _char_offset(temp[0])
                temp=temp[1:]
                col_int = 26 * col_int
            col_int += _char_offset(temp)
            return self.get_raw(col_int)
        
        assert False, f"Should not get here {col}"


    def strict(self)->StrictXlsxRowParser:
        return StrictXlsxRowParser(row_number=self.row_number, user_data=self.user_data, cells=self.cells)

    def int(self, col:str|int, default=INVALID, strict:bool=True)->int:
        return self.get_raw(col).int(strict=strict, default=default)

    def float(self, col:str|int, default=INVALID, strict:bool=True)->float:
        return self.get_raw(col).float(strict=strict, default=default)

    def str(self, col:str|int, default=INVALID, strict:bool=True)->str:
        return self.get_raw(col).str(strict=strict, default=default)

    def bool(self, col:str|int, strict=False, default=False)->bool:
        return self.get_raw(col).bool(strict=strict, default=default)

    def is_none(self, col:str|int)->bool:
        return self.get_raw(col).value is None

    def is_empty_block(self, start:str|int|XlsxCell, end:str|int|XlsxCell)->bool:


        start = col_to_int(start)
        end = col_to_int(end)
        for n, cell in enumerate(self.cells):
            if n >=start:
                if n <=end:
                    if cell.value is not None:
                        return False
        return True
    def is_empty(self)->bool:
        for cell in self.cells:
            if not cell.is_none:
                return False
        return True

    def datetime(self, col:str|int, default=INVALID, strict:bool=True)->datetime:
        return self.get_raw(col).datetime(strict=strict, default=default)

    def as_list(self):
        return [l.value for l in self.cells]


@dataclass(frozen=True)
class StrictXlsxRowParser:
    row:XlsxRow

    def raw_cell(self, col:str|int)->XlsxCell:
        return self.row.raw_cell(col)

    def __post_init__(self):
        assert isinstance(self.row, XlsxRow)

    def int(self, col:str|int, default=INVALID)->int:
        return self.raw_cell(col).int(strict=True, default=default)

    def float(self, col:str|int, default=INVALID)->float:
        return self.raw_cell(col).float(strict=True, default=default)

    def str(self, col:str|int, default=INVALID)->str:
        return self.raw_cell(col).str(strict=True, default=default)

    def bool(self, col:str|int, default=INVALID)->bool:
        return self.raw_cell(col).bool(strict=True, default=default)

    def is_none(self, col:str|int)->bool:
        return self.raw_cell(col).value is None

    def datetime(self, col:str|int, default=INVALID)->datetime:
        return self.raw_cell(col).datetime(strict=True, default=default)

@dataclass(frozen=True)
class StrictDefaultXlsxRowParser:
    row:XlsxRow

    def raw_cell(self, col:str|int)->XlsxCell:
        return StrictXlsxRowParser(self.row).raw_cell(col=col)

    def __post_init__(self):
        assert isinstance(self.row, XlsxRow)

    def int(self, col:str|int, default:int)->int:
        assert default != INVALID
        return self.raw_cell.int(default=default)

    def float(self, col:str|int, default:float)->float:
        assert default != INVALID
        return self.raw_cell(col).float(default=default)

    def str(self, col:str|int, default:str)->str:
        assert default != INVALID
        return self.raw_cell(col).str(default=default)

    def bool(self, col:str|int, default:bool)->bool:
        assert default != INVALID
        return self.raw_cell(col).bool(default=default)

    def is_none(self, col:str|int)->bool:
        return self.raw_cell(col).value is None

    def datetime(self, col:str|int, default:datetime)->datetime:
        assert default != INVALID
        return self.raw_cell(col).datetime(default=default)

@dataclass(frozen=True)
class CustomXlsxRowParser:
    row:XlsxRow

    def raw_cell(self, col:str|int)->XlsxCell:
        return StrictXlsxRowParser(self.row).raw_cell(col=col)

    def __post_init__(self):
        assert isinstance(self.row, XlsxRow)

    def int(self, col:str|int|None, parse_function:any)->int:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, int)

    def float(self, col:str|int|None, parse_function:any)->float:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, float)

    def str(self, col:str|int|None, parse_function:any)->str:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, str)

    def bool(self, col:str|int|None, parse_function:any)->bool:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, bool)

    def datetime(self, col:str|int|None, parse_function:any)->datetime:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, datetime)


@dataclass(frozen=True)
class ParsedRow(ABC):
    row:XlsxRow
    user_data:dict

    @staticmethod
    @abstractmethod
    def parse(xlsx_row:XlsxRow, user_data:dict={})->ParsedRow:
        raise NotImplementedError()

    @staticmethod
    @abstractmethod
    def is_match(xlsx_row:XlsxRow, user_data:dict={})->bool:
        raise NotImplementedError()

    @property
    @abstractmethod
    def as_dict(self)->dict:
        raise NotImplementedError()

    @property
    def as_json(self)->str:
        return json.dumps(self.as_dict, indent=2, sort_keys=True)

@dataclass(frozen=True)
class ErrorRow(ParsedRow):
    errors:list

    @staticmethod
    def parse(xlsx_row:XlsxRow, user_data:dict)->ParsedRow:
        errors=DictParser(raw=user_data).list("errors")
        # assert "errors" in user_data
        return ErrorRow(row=xlsx_row, user_data=user_data, errors=user_data['errors'])

    @staticmethod
    def is_match(xlsx_row:XlsxRow, user_data:dict)->bool:
        return False

    @property
    def as_dict(self)->dict:
        return dict(raw=[str(c.value) for c in self.row.cells],row_number=self.row.row_number,
                    errors=self.errors)


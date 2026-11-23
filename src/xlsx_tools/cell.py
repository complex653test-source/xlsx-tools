import json
import logging
from abc import ABC
from abc import abstractmethod
from dataclasses import dataclass
from datetime import datetime
from typing import Callable
from typing import Dict
from typing import Type

from openpyxl.cell.cell import Cell

from xlsx_tools.utils import INVALID


@dataclass(frozen=True)
class XlsxCell:
    raw:Cell

    def __post_init__(self):
        if isinstance(self.raw, XlsxCell):
            object.__setattr__(self, "raw", self.raw.raw)
        assert isinstance(self.raw, Cell)

    @property
    def column_index(self)->int:
        return self.raw.col_idx

    @property
    def column_letter(self)->str:
        return self.raw.column_letter

    @property
    def is_blank(self)->bool:
        if self.value is None:
            return True
        if "" == self.str(default="", strict=False):
            return True
        logging.debug(f"Not blank: {self.value}")
        return False

    @property
    def is_none(self)->bool:
        return self.value is None
    
    def bool(self, default=INVALID, null_ok:bool=False, transform:Callable=INVALID)->bool:
        return self.clean_value(default=default, transform=transform, required_type=bool, null_ok=null_ok)

    def datetime(self, default=INVALID, null_ok:bool=False, transform:Callable=INVALID)->datetime:
        return self.clean_value(default=default, transform=transform, required_type=datetime, null_ok=null_ok)

    def float(self, default=INVALID, null_ok:bool=False,  transform:Callable=INVALID)->float:
        return self.clean_value(default=default, transform=transform, required_type=float, null_ok=null_ok)

    def int(self, default=INVALID, null_ok:bool=False, transform:Callable=INVALID)->int:
        return self.clean_value(default=default, transform=transform, required_type=int, null_ok=null_ok)

    def str(self, default=INVALID, null_ok:bool=False, transform:Callable=INVALID)->str:
        return self.clean_value(default=default, transform=transform, required_type=str, null_ok=null_ok)

    @property
    def value(self):
        return self.raw.value

    def clean_value(self, default:any=INVALID, null_ok:bool=False, transform:Callable=INVALID, required_type:Type=INVALID)->any:
        def _type_check(value:any, required_type:Type=INVALID)->bool:

            if required_type is INVALID:
                return True
            if isinstance(value, required_type):
                return True
            if issubclass(type(value), required_type):
                return True
            raise ValueError(f"Value {value} must be required type {required_type} not {type(value)}")
        def _apply_default(value:any, default:any)->any:
            if value is not None:
                return value
            if default is not INVALID:
                return default
            return value

        ret_val=self.value
        if transform is not INVALID:
            ret_val = transform(cell=self, default=default)
        ret_val = _apply_default(value=ret_val, default=default)
        if ret_val is None:
            if null_ok:
                return None
            raise ValueError("Call cannot be Null")

        _type_check(value=ret_val, required_type=required_type)

        return ret_val  # self.raw_value_or_default(default=default)

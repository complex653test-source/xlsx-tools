
class INVALID:
    pass


from dataclasses import dataclass
from typing import Callable
#(Pdb) isinstance(datatype, Callable)
True

@dataclass(frozen=True)
class ChainAssert:
    value:any
    name:str="UNDEFINED"

    def issubclass(self, datatype:Callable)->ChainAssert:
        assert isinstance(datatype, Callable)
        # print(type(self.value))
        # print(type(datatype))
        if not issubclass(self.value.__class__, datatype):
            raise TypeError(f"{self.name} value must be a cubclass of {datatype}: ({type(self.value)})")
        return self


    def isinstance(self, datatype:Callable)->ChainAssert:
        assert isinstance(datatype, Callable)
        if not isinstance(self.value, datatype):
            raise TypeError(f"{self.name} value must be a {datatype}: ({type(self.value)})")
        return self

    def is_none(self)->ChainAssert:
        if self.value is not None:
            raise ValueError(f"{self.name} ({self.value}) value must be None")

    def compare(self, lt:any|None, le:any|None, gt:any|None, ge:any|None, eq:any|None)->ChainAssert:
        if eq is not None:
            if not self.value == eq:
                raise ValueError(f"{self.name} ({self.value}) value must be == {eq}")
        if lt is not None:
            if not self.value < lt:
                raise ValueError(f"{self.name} ({self.value}) value must be < {lt}")
        if le is not None:
            if not self.value <= le:
                raise ValueError(f"{self.name} ({self.value}) value must be <= {le}")
        if gt is not None:
            if not self.value > gt:
                raise ValueError(f"{self.name} ({self.value}) value must be > {gt}")
        if ge is not None:
            if not self.value >= ge:
                raise ValueError(f"{self.name} ({self.value}) value must be >= {ge}")
        return self

    def __len__(self)->int:
        return len(self.value)

    def len_compare(self, lt:any|None, le:any|None, gt:any|None, ge:any|None, eq:any|None)->ChainAssert:
        ChainAssert(value=len(self), name=f"len({self.name})").compare(lt=lt, le=le, gt=gt, ge=ge, eq=eq)  # raises exception
        return self






def col_to_int(col:str|int)->int:
    if isinstance(col, int):
        return col
    if isinstance(col, str):
        return excel_col_to_index(col)
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
            print(temp)
            print(col_int)
            col_offset=0
            col_int += _char_offset(temp[0])+1
            temp=temp[1:]
            col_int = 26 * col_int
            print(col_int)
            print(temp)
        print(temp)
        col_int += _char_offset(temp)
        print(col_int)
        return col_int

    assert False, f"Should not get here {col}"



def excel_col_to_index(col: str) -> int:
  index = 0
  for char in col.upper():
    # ord('A') is 65, so ord(char) - 64 gives A=1, B=2, ..., Z=26
    index = index * 26 + (ord(char) - 64)
  return index - 1  # Subtract 1 for zero-indexing


assert col_to_int("A")==0
assert col_to_int("Z")==25
assert col_to_int("AA")==26
assert col_to_int("AB")==27
assert col_to_int("AC")==28
assert col_to_int("AD")==29
assert col_to_int("AE")==30
assert col_to_int("AF")==31
assert col_to_int("AG")==32
import pytest
from local_packages import zipxl
from logging import INFO, ERROR, getLogger

logger = getLogger('test')

def test_001():
    xl: zipxl.ImageBook = zipxl.ImageBook("./downloads/09390-JGr-Y含む-エクセル数値-210114.xlsx")
    sheetlist: list[zipxl.Sheet] = xl.__Sheets()
    assert sheetlist[0].name == 'sheet4'

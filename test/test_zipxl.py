import pytest
from local_packages import zipxl
from logging import INFO, ERROR, getLogger

logger = getLogger('test')
testfile = r'./downloads/09390-JGr-Y含む-エクセル数値-210114.xlsx'

def test_getSheetName():
    xl: zipxl.ImageBook = zipxl.ImageBook()
    xl.open(testfile)
    assert xl.Sheets[0].name == 'sheet4'

def test_getImageName():
    xl: zipxl.ImageBook = zipxl.ImageBook()
    xl.open(testfile)
    assert xl.Sheets[0].Images[0].name == 'image23.emf'

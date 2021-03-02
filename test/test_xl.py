import pytest
from local_packages import xl
from logging import INFO, ERROR, getLogger

logger = getLogger('test')
testfile = r'./downloads/09390-JGr-Y含む-エクセル数値-210114.xlsx'

def test_001():
    myxl: xl.xlimg = xl.xlimg(testfile)
    res = myxl.getxlimages()
    myxl.close()
    assert "SCRIBE" in res

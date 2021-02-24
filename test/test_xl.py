import pytest
from local_packages import xl
from logging import INFO, ERROR, getLogger

logger = getLogger('test')

def test_001():
    myxl: xl.xlimg = xl.xlimg(r"09390-JGr-Y含む-エクセル数値-210114.xls")
    res = myxl.getxlimages()
    assert "SCRIBE" in res

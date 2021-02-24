import pytest
from local_packages import fcon
from logging import INFO, ERROR, getLogger

logger = getLogger('test')

def test_001():
    mydata = fcon.openXlFile("./")
    assert ".xls" in mydata

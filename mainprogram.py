import xlwings as xw
from typing import List
from local_packages import fcon
from local_packages import xl

defaultFileDirctory: str = r"C:\Users\take\Documents\GitHub\xlocr_for_image_on_sheets"

if __name__ == "__main__":
    targetXlFileFullPath = fcon.openXlFile(defaultFileDirctory)
    print(targetXlFileFullPath)
    mydata: xl.xlimg = xl.xlimg(targetXlFileFullPath)
    RecgnitionList: List[str] = []
    item: xw.Sheet
    for item in mydata.wb.sheets:
        RecgnitionList = mydata.getStringsFromSheet(item)
        print(item.name + ":")
        if RecgnitionList != None:
            for RecgnitionString in RecgnitionList:
                print("    - " + RecgnitionString)

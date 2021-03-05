import pylightxl
from local_packages import zipxl
from local_packages import fcon
from local_packages import xl
from enum import Enum

defaultFileDirctory: str = r"./Downloads/"
remarkDict: dict = {
    0: "HOLE CTR",
    1: "SCRIBE LINE",
    2: "MEASURING POINT",
    3: "SHIM",
    4: "CP",
}


class remark(Enum):
    HoleCtr = "HOLE CTR"
    ScribeLine = "SCRIBE LINE"
    MeasuringPoint = "MEASURING POINT"
    Shim = "SHIM"
    Cp = "CP"


if __name__ == "__main__":
    targetXlFileFullPath = fcon.openXlFile(defaultFileDirctory)
    print(targetXlFileFullPath)
    mydata: pylightxl.Database = pylightxl.readxl(targetXlFileFullPath)
    zipdata: zipxl.ImageBook = zipxl.ImageBook()
    zipdata.open(targetXlFileFullPath)
    for targetSheetname in mydata.ws_names:
        col = mydata.ws(targetSheetname).col(37)
        RemarkCounter = 0
        for cell in col:
            if remark.ScribeLine.value == cell:
                RemarkCounter += 1
        PictureCount = 0
        for targetSheet in zipdata.Sheets:
            if targetSheet.displayName == targetSheetname:
                PictureCount = len(targetSheet.Pictures)
                break
        print(targetSheetname, RemarkCounter, PictureCount)

    print()

# HOLE CTR
# SCRIBE LINE
# MEASURING POINT
# SHIM
# CP
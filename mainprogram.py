import pylightxl
from local_packages import zipxl
from local_packages import fcon
from local_packages import xl
from enum import Enum
import pytesseract
from PIL import Image
import os

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


def getLinesFromImage(inputImage: Image.Image) -> list[str]:
    custom_oem_psm_config = r'--psm 6 -l eng -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789"'
    try:
        # nolook so small image
        if inputImage.size[0] * inputImage.size[1] <= 1600:
            result: str = ""
            return result
        result: str = pytesseract.image_to_string(
            inputImage, config=custom_oem_psm_config
        )
    except TypeError:
        result: str = ""
    return result.split("\n")


def hasRemark(inputString: str) -> bool:
    for target in remark:
        if target.value in inputString:
            return True
    return False


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
                for item in targetSheet.Pictures:
                    buf: list[str] = getLinesFromImage(item.Image())
                    for line in buf:
                        if hasRemark(line):
                            print(line)
            break # displayName is appear only once
        print(targetSheetname, RemarkCounter, PictureCount)

# HOLE CTR
# SCRIBE LINE
# MEASURING POINT
# SHIM
# CP
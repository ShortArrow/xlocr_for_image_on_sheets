import pylightxl
from local_packages import zipxl
from local_packages import fcon
from local_packages import xl
from enum import Enum
import pytesseract
from PIL import Image
import os
import csv

defaultFileDirctory: str = r"./Downloads/"
remarkColumn: int = 37
OutputFileDirectory: str = r"./Downloads/output.csv"


class remark(Enum):
    HoleCtr = "HOLE CTR"
    ScribeLine = "SCRIBE LINE"
    MeasuringPoint = "MEASURING POINT"
    Shim = "SHIM"
    Cp = "CP"


class checkDictionary:
    def __init__(self, mark: remark) -> None:
        self.__mark = mark
        self.__has = True

    def setThisMark(self, has: bool) -> None:
        self.__has = has

    def has(self) -> bool:
        return self.__has

    def mark(self) -> remark:
        return self.__mark

    def __str__(self) -> str:
        return "{" + self.__mark.value + ":" + self.__has + "}"


def getLinesFromImage(inputImage: Image.Image) -> list[str]:
    custom_oem_psm_config = (
        r'-c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789" --psm 6'
    )
    try:
        # nolook so small image
        if inputImage.size[0] * inputImage.size[1] <= 1600:
            result: str = ""
            return result
        result: str = pytesseract.image_to_string(
            inputImage, lang="eng", config=custom_oem_psm_config
        )
    except TypeError:
        result: str = ""
    return result.split("\n")


def ocr(inputImage: Image.Image) -> list[str]:
    pass


def hasRemark(inputString: str) -> bool:
    for target in remark:
        if target.value in inputString:
            return (True, target)
    return (False, None)


if __name__ == "__main__":
    targetXlFileFullPath = fcon.openXlFile(defaultFileDirctory)
    print(targetXlFileFullPath)
    mydata: pylightxl.Database = pylightxl.readxl(targetXlFileFullPath)
    myBook: zipxl.ImageBook = zipxl.ImageBook()
    myBook.open(targetXlFileFullPath)
    resData: list[list] = []
    for targetSheetname in mydata.ws_names:
        buflist: list = []
        buflist.append(targetSheetname)
        col = mydata.ws(targetSheetname).col(remarkColumn)
        RemarkCounter = 0
        for cell in col:
            if remark.ScribeLine.value == cell:
                RemarkCounter += 1
        buflist.append(RemarkCounter)
        PictureCount = 0
        for targetSheet in myBook.Sheets:
            if targetSheet.displayName == targetSheetname:
                buflist.append(len(targetSheet.Pictures))
                for index, item in enumerate(targetSheet.Pictures):
                    buf: list[str] = getLinesFromImage(item.Image().convert("LA"))
                    for line in buf:
                        containRemark, hitname = hasRemark(line)
                        if containRemark:
                            buflist.append(checkDictionary(hitname))
                break  # displayName is appear only once
        resData.append(buflist)
    with open(OutputFileDirectory, "w") as f:
        writer = csv.writer(f)
        for row in resData:
            writer.writerow(row[0:3] + row[0:3])

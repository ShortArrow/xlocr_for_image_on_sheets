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


class ColumnTitle(Enum):
    SheetName = "SheetName"
    Remark = "Remark"
    Pictures = "Pictures"
    Ocr = "OCR"


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
    if targetXlFileFullPath == "":
        exit()
    print(targetXlFileFullPath)
    mydata: pylightxl.Database = pylightxl.readxl(targetXlFileFullPath)
    myBook: zipxl.ImageBook = zipxl.ImageBook()
    myBook.open(targetXlFileFullPath)
    resData: list[dict] = []
    for targetSheetname in mydata.ws_names:
        buflist: dict = {}
        buflist[ColumnTitle.SheetName] = targetSheetname
        col = mydata.ws(targetSheetname).col(remarkColumn)
        bufdic: dict = {}
        for cell in col:
            for targetRemark in remark:
                if cell == targetRemark.value:
                    if bufdic.__contains__(targetRemark):
                        bufdic[targetRemark] += 1
                    else:
                        bufdic[targetRemark] = 1
        buflist[ColumnTitle.Remark] = bufdic
        bufdic = {}
        PictureCount = 0
        for targetSheet in myBook.Sheets:
            if targetSheet.displayName == targetSheetname:
                buflist[ColumnTitle.Pictures] = len(targetSheet.Pictures)
                for index, item in enumerate(targetSheet.Pictures):
                    buf: list[str] = getLinesFromImage(item.Image().convert("LA"))
                    for line in buf:
                        containRemark, hitname = hasRemark(line)
                        if containRemark:
                            if bufdic.keys().__contains__(hitname):
                                bufdic[hitname] += 1
                            else:
                                bufdic[hitname] = 1
                buflist[ColumnTitle.Ocr] = bufdic
                break  # displayName is appear only once
        resData.append(buflist)
    with open(OutputFileDirectory, "w", newline="\n") as f:
        writer = csv.writer(f)
        columnHeaders = []
        columnHeaders.append(ColumnTitle.SheetName.value)
        for targetRemark in remark:
            columnHeaders.append(targetRemark.value + "_" + ColumnTitle.Remark.value)
        columnHeaders.append(ColumnTitle.Pictures.value)
        for targetRemark in remark:
            columnHeaders.append(targetRemark.value + "_" + ColumnTitle.Ocr.value)
        writer.writerow(columnHeaders)
        for row in resData:
            makingline = []
            makingline.append(row[ColumnTitle.SheetName])
            for targetRemark in remark:
                try:
                    makingline.append(row[ColumnTitle.Remark][targetRemark])
                except KeyError:
                    makingline.append(0)
            makingline.append(row[ColumnTitle.Pictures])
            for targetRemark in remark:
                try:
                    makingline.append(row[ColumnTitle.Ocr][targetRemark])
                except KeyError:
                    makingline.append(0)
            writer.writerow(makingline)

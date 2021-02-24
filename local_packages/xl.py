from typing import Any, List
import xlwings as xw
from PIL import Image
import pytesseract
from xlwings.main import Pictures
from PIL import ImageGrab
import time
import win32clipboard
from ctypes import WinDLL, windll

# import pywin
# import pyocr


class xlimg:
    wb: xw.Book

    def __init__(self, xlpath: str) -> None:
        self.wb = xw.Book(xlpath)

    def getxlimages(self):
        # tools = pyocr.get_available_tools()
        # tool = tools[0]
        # builder = pyocr.builders.TextBuilder(tesseract_layout=6)
        sht: xw.Sheet = self.wb.sheets[3]
        pic: xw.Picture = sht.pictures[0]
        pic.api.Copy()
        myimage: Image.Image = ImageGrab.grabclipboard()
        return self.getStringFromImage(myimage)
    
    def ClearClip(self):
        if win32clipboard.OpenClipboard(None):
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()

    def getClip(self):
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            return win32clipboard.GetClipboardData(win32clipboard.CF_DIB)

    def getImagesFromSheet(self, inputSheet: xw.Sheet) -> List[Image.Image]:
        item: xw.Picture = None
        resultImageList: List[Image.Image] = []
        for item in inputSheet.pictures:
            inputSheet.activate()
            inputSheet.select()
            item.api.Copy()
            time.sleep(0.1)
            myimage: Any = ImageGrab.grabclipboard()
            # myimage: Any = self.getClip()
            if isinstance(myimage,Image.Image):
                resultImageList.append(myimage)
        return resultImageList

    def getStringFromImage(self, inputImage: Image.Image) -> str:
        custom_oem_psm_config = r'--psm 6 -l eng -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789"'
        result: str = pytesseract.image_to_string(
            inputImage, config=custom_oem_psm_config
        )
        return result

    def getStringsFromSheet(self, inputSheet: xw.Sheet) -> List[str]:
        imageList: List[Image.Image] = self.getImagesFromSheet(inputSheet)
        stringList: List[str] = []
        for item in imageList:
            stringList.append(self.getStringFromImage(item))
        return stringList

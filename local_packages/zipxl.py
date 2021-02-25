import zipfile


class ImageBook:

    def __init__(self, fileName: str) -> None:
        self.__zf: zipfile.ZipFile
        self.__sheetFolder: str = "xl/worksheets/"
        self.__relayFolder: str = "xl/worksheets/_rels/"
        self.__mediaFolder: str = "xl/media/"
        self.__zf = zipfile.ZipFile(fileName)
        # self.zf.extract(name, path='C:\\temp\\images')

    def __del__(self) -> None:
        self.zf.close()

    def getImageList(self) -> list[str]:
        res: list[str] = []
        for name in self.zf.namelist():
            if name.startswith(self.__mediaFolder):
                res.append(name)
        return res

    def getImagePathsFromSheet(self):
        pass

    def getSheetPaths(self) -> list[str]:
        res: list[str] = []
        for name in self.zf.namelist():
            if name.startswith(self.__sheetFolder):
                if not name.startswith(self.__relayFolder):
                    res.append(name)
        return res

    def getSheetNameFromSheetPath(self, sheet: str) -> str:
        return sheet.replace(self.__sheetFolder, "").replace(".xml", "")

    def getSheetNames(self) -> list[str]:
        source: list[str] = self.getSheetPaths()
        response: list[str] = []
        for item in source:
            response.append(self.getSheetNameFromSheetPath(item))
        return response

    def getRelayPaths(self) -> list[str]:
        source:list[str] = self.getSheetNames
        responcse:list[str] = []
        for item in source:
            responcse.append(self.getRelayPathFromSheetName(item))

    def getRelayPathFromSheetName(self, sheetname: str) -> str:
        return self.__relayFolder + sheetname

    def getImagePathsFromRelay():
        pass

    # image/sheet/bookでclassを作って連携させる

if __name__ == "__main__":
    xl: ImageBook = ImageBook(
        r"C:\Users\take\Documents\GitHub\xlocr_for_image_on_sheets\09390-JGr-Y含む-エクセル数値-210114.xlsx"
    )
    sheetlist: list[str] = xl.getSheetNames()
    for item in sheetlist:
        print(item)
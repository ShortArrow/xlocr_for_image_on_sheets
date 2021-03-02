import zipfile


class zipxlelement:
    def __init__(self) -> None:
        pass

    @property
    def sheetFolder(self) -> str:
        return "xl/worksheets/"

    @property
    def relayFolder(self) -> str:
        return "xl/worksheets/_rels/"

    @property
    def mediaFolder(self) -> str:
        return "xl/media/"

    @property
    def sheetfileExtension(self) -> str:
        return ".xml"


class Sheet(zipxlelement):
    def __init__(self, name: str) -> None:
        self.name = name
        self.parent = None
        self.root = None


class Image(zipxlelement):
    def __init__(self, name: str) -> None:
        self.name = name
        self.parent = None
        self.root = None


class ImageBook(zipxlelement):
    def __init__(self) -> None:
        self.Sheets: list[Sheet] = []

    def open(self, fileName: str) -> None:
        self.__zf: zipfile.ZipFile = zipfile.ZipFile(fileName)
        self.Sheets = self.__Sheets()

    def __del__(self) -> None:
        self.__zf.close()

    def __Sheets(self) -> list[Sheet]:
        res: list[Sheet] = []
        for item in self.getSheetNames():
            res.append(Sheet(item))
            res[len(res) - 1].parent = self
            res[len(res) - 1].root = self
        return res

    def getImageList(self) -> list[str]:
        res: list[str] = []
        for name in self.__zf.namelist():
            if name.startswith(self.mediaFolder):
                res.append(name)
        return res

    def getImagePathsFromSheet(self):
        pass

    def getSheetPaths(self) -> list[str]:
        res: list[str] = []
        for name in self.__zf.namelist():
            if name.startswith(self.sheetFolder):
                if not name.startswith(self.relayFolder):
                    res.append(name)
        return res

    def getSheetNameFromSheetPath(self, sheet: str) -> str:
        return sheet.replace(self.sheetFolder, "").replace(
            self.sheetfileExtension, ""
        )

    def getSheetNames(self) -> list[str]:
        source: list[str] = self.getSheetPaths()
        response: list[str] = []
        for item in source:
            response.append(self.getSheetNameFromSheetPath(item))
        return response

    def getRelayPaths(self) -> list[str]:
        source: list[str] = self.getSheetNames
        responcse: list[str] = []
        for item in source:
            responcse.append(self.getRelayPathFromSheetName(item))

    def getRelayPathFromSheetName(self, sheetname: str) -> str:
        return self.__relayFolder + sheetname

    def getImagePathsFromRelay():
        pass

    # image/sheet/bookでclassを作って連携させる


if __name__ == "__main__":
    xl: ImageBook = ImageBook()
    xl.open("./downloads/09390-JGr-Y含む-エクセル数値-210114.xlsx")
    for item in xl.Sheets:
        print(item.name)
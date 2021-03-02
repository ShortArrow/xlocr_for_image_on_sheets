import zipfile


class Element:
    def __init__(self, name: str, parent: "Element", root: "Element") -> None:
        self.name = name
        self.parent = parent
        self.root = root

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

    @property
    def relayfileExtension(self) -> str:
        return ".xml.rels"


class Image(Element):
    pass


class Sheet(Element):
    def __init__(self, name: str, parent: Element, root: Element) -> None:
        super().__init__(name=name, parent=parent, root=root)
        self.Images = self.__Images()

    def __Images(self) -> list[Image]:
        res: list[Image] = []
        buf: Image = None
        while True:
            buf = self.__Image(len(res))
            if buf.name == "":
                break
            res.append(buf)
        return res

    def __Image(self, Index: int) -> Image:
        try:
            f = self.root.zf.open(
                self.relayFolder + self.name + self.relayfileExtension, "r"
            )
            relsdata: str = f.read(-1)
            f.close()
        except KeyError:
            relsdata: str = ""
        startindex: int = 1
        for time in range(Index + 1):
            startindex = str(relsdata).find('/image"', startindex + 1)
            if startindex == -1:
                return Image(name="", parent=None, root=None)
        startindex = str(relsdata).find('Target="', startindex)
        finalindex = str(relsdata).find('"/>', startindex)
        startindex = str(relsdata).rfind("/", startindex, finalindex) + 1
        return Image(str(relsdata)[startindex:finalindex], parent=None, root=None)


class ImageBook(Element):
    def __init__(self) -> None:
        super().__init__(name="", parent=None, root=None)
        self.Sheets: list[Sheet] = []

    def open(self, fileName: str) -> None:
        self.zf: zipfile.ZipFile = zipfile.ZipFile(fileName)
        self.name = fileName
        self.Sheets = self.__Sheets()

    def __del__(self) -> None:
        self.zf.close()

    def __Sheets(self) -> list[Sheet]:
        res: list[Sheet] = []
        for item in self.getSheetNames():
            res.append(Sheet(name=item, parent=self, root=self))
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
        for name in self.zf.namelist():
            if name.startswith(self.sheetFolder):
                if not name.startswith(self.relayFolder):
                    res.append(name)
        return res

    def getSheetNameFromSheetPath(self, sheet: str) -> str:
        return sheet.replace(self.sheetFolder, "").replace(self.sheetfileExtension, "")

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
        return self.relayFolder + sheetname

    def getImagePathsFromRelay():
        pass


if __name__ == "__main__":
    xl: ImageBook = ImageBook()
    xl.open("./downloads/09390-JGr-Y含む-エクセル数値-210114.xlsx")
    for item in xl.Sheets:
        if len(item.Images):
            print(item.Images[0].name,item.Images[1].name)
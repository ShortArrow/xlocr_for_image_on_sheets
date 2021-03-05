import zipfile
from PIL import Image
import io


class Element:
    def __init__(
        self, name: str, parent: "Element" = None, root: "Element" = None
    ) -> None:
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


class Picture(Element):
    def Canvas(self) -> Image.Image:
        z = zipfile.ZipFile(self.root.name)
        img:Image.Image = Image.open(io.BytesIO(z.read(self.mediaFolder +  self.name)))
        z.close()
        # img.show()
        return img


class Sheet(Element):
    def __init__(self, name: str, parent: Element, root: Element) -> None:
        super().__init__(name=name, parent=parent, root=root)
        self.Pictures = self.__Pictures()

    def __Pictures(self) -> list[Picture]:
        res: list[Picture] = []
        buf: Picture = None
        while True:
            buf = self.__Pciture(len(res))
            if buf.name == "":
                break
            res.append(buf)
        return res

    def __Pciture(self, Index: int) -> Picture:
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
                return Picture(name="", parent=self, root=self.parent)
        startindex = str(relsdata).find("Target=", startindex)
        finalindex = str(relsdata).find('"/>', startindex)
        startindex = str(relsdata).rfind("/", startindex, finalindex) + 1
        return Picture(
            str(relsdata)[startindex:finalindex], parent=self, root=self.parent
        )


class ImageBook(Element):
    def __init__(self) -> None:
        super().__init__(name="")
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
    img=xl.Sheets[0].Pictures[0].Canvas()
    img.show()

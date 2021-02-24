import zipfile


class XlImageGetter:
    zf: zipfile.ZipFile
    mediaFolder: str = "xl/media/"
    sheetFolder: str = "xl/worksheets/"
    relayFolder: str = "xl/worksheets/_rels/"

    def __init__(self, fileName: str) -> None:
        self.zf = zipfile.ZipFile(fileName)
        # self.zf.extract(name, path='C:\\temp\\images')

    def __del__(self) -> None:
        self.zf.close()

    def getImageList(self) -> list[str]:
        res: list[str] = []
        for name in self.zf.namelist():
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
        return sheet.replace(self.sheetFolder, "").replace(".xml", "")

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
        return self.relayFolder + sheetname

    def getImagePathsFromRelay():
        pass

    # image/sheet/bookでclassを作って連携させる

if __name__ == "__main__":
    xl: XlImageGetter = XlImageGetter(
        r"C:\Users\take\Documents\GitHub\xlocr_for_image_on_sheets\09390-JGr-Y含む-エクセル数値-210114.xlsx"
    )
    sheetlist: list[str] = xl.getSheetNames()
    for item in sheetlist:
        print(item)
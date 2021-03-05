import pylightxl

def getcol(targetXlFileFullPath: int) -> list[str]:
    mydata: pylightxl.Database = pylightxl.readxl(targetXlFileFullPath)
    return mydata.ws(mydata.ws_names[0]).col(37)
    
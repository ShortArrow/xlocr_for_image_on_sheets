import pylightxl
from local_packages import fcon
from local_packages import xl

defaultFileDirctory: str = r"./Downloads/"

if __name__ == "__main__":
    targetXlFileFullPath = fcon.openXlFile(defaultFileDirctory)
    print(targetXlFileFullPath)
    mydata: pylightxl.Database = pylightxl.readxl(targetXlFileFullPath)
    for item in mydata.ws_names:
        col = mydata.ws(item).col(37)
        pass
        RemarkCounter = 0
        for cell in col:
            if "SCRIBE LINE" == cell:
                RemarkCounter += 1
        print(item,RemarkCounter)

# HOLE CTR
# SCRIBE LINE
# MEASURING POINT
# SHIM
# CP
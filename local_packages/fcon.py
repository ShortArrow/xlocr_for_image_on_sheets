from tkinter import filedialog, mainloop
import os


def openXlFile(defaultpath:str) -> str:
    typ = [('数値表Excelファイル', '*.xls*')]
    if defaultpath :
        dir = defaultpath
    else :
        dir = os.environ['userprofile'] + '/Desktop'
    # print(dir)
    fle = filedialog.askopenfilename(filetypes=typ, initialdir=dir)
    return fle


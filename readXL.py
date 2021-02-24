import xlwings as xw
wb:xw.Book = xw.Book('09390-JGr-Y含む-エクセル数値-210114.xls')  # connect to an existing file in the current working directory
sht:xw.Sheet = wb.sheets[2]
print(sht.cells(14,37).value)
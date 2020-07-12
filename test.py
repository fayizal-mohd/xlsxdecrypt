import os
import sys
import xlrd
str='java -jar xlsxunpass/xlsxunpass.jar '+sys.argv[1]+' decrypt.xlsx '+sys.argv[2]
os.system(str)
book=xlrd.open_workbook('decrypt.xlsx')
sheet=book.sheet_by_index(0)
cell = sheet.cell(0,0)
print cell.value

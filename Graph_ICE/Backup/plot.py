"""Plot Graph by xlrd module"""
import xlrd
file = xlrd.open_workbook('Data.xls')
##file.cell(0,0).value
print(file.sheet_names())
file.cell(0,0).value

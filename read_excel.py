# Writeten by Chun-Hsiang Chao
# Date:20250521
import openpyxl #apt install python3-openpyxl
workbook=openpyxl.load_workbook('my_financial_3.xlsx')
sheet=workbook.worksheets[0]
print(sheet.title,sheet['F1'].value)
print(sheet.max_row,sheet.max_column)

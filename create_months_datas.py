# Writeten by Chun-Hsiang Chao
# Date:20250612
import datetime
import openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, PatternFill, NamedStyle, GradientFill
from openpyxl.utils import FORMULAE
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference

table_name=["202507","202507預算","202508","202508預算"]

workbook=openpyxl.load_workbook('my_financial_3.xlsx')

sheet=workbook.worksheets[0]
sheet.page_setup_orientation=sheet.ORIENTATION_LANDSCAPE
sheet.page_setup.paperSize=sheet.PAPERSIZE_A4
sheet.page_setup.fitToHeight=0
sheet.page_setup.fitToWidth=1


source_sheet=workbook["yearMonth"]
workbook.copy_worksheet(source_sheet)
workbook["yearMonth Copy"].title=table_name[0]

source_sheet=workbook["yearMonth預算"]
workbook.copy_worksheet(source_sheet)
workbook["yearMonth預算 Copy"].title=table_name[1]
sheet=workbook[table_name[1]]
sheet['B2']=table_name[1][4:6]
string_to_replace = "=yearMonth"
replacement_string ="$'"+table_name[0]+"'.D98"



for row in sheet.iter_rows():
	for cell in row:
		if cell.value is not None and string_to_replace in str(cell.value):
			cell.value = str(cell.value).replace(string_to_replace, replacement_string)






workbook.save('test_my_financial_3.xlsx')
workbook.close()

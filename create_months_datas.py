# Writeten by Chun-Hsiang Chao
# Date:20250618
import datetime
import openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, PatternFill, NamedStyle, GradientFill
from openpyxl.utils import FORMULAE
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import Cell
#from xlcalculator import Evaluator
#pip install xlcalculator --break-system-packages

table_name=["202507","202507預算","202508","202508預算"]
print(len(table_name))

workbook=openpyxl.load_workbook('my_financial_3.xlsx',data_only=False)

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

string_to_replace = "yearMonth"
replacement_string ="'"+table_name[0]+"'"

amountOfRows = sheet.max_row
amountOfColumns = sheet.max_column

i = 0
for r in range(1,sheet.max_row+1):
    for c in range(1,sheet.max_column+1):
        s = sheet.cell(r,c).value
        if s != None and string_to_replace in str(s):
            sheet.cell(r,c).value = s.replace(string_to_replace,replacement_string)
						
            sheet.cell(r,c).data_type='f'

            print("row {}	col {}	: {}".format(r,c,s))
            i += 1


print(i)




workbook.save('test_my_financial_3.xlsx')
workbook.close()

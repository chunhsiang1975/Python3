# Writeten by Chun-Hsiang Chao
# Date:20250522
import openpyxl 
from openpyxl.styles import Font, Color, Alignment, Border, Side, PatternFill
#from openpyxl.styles.differential import DifferentialStyle
#from openpyxl.formatting.rule import Rule

from openpyxl.drawing.fill import PatternFillProperties


table_name=["Test1","Test2","Test3"]
interger_i=0
workbook=openpyxl.Workbook() #apt install python3-openpyxl
workbook.template=True
sheet=workbook.worksheets[0]
sheet['A1']='Hello'
sheet['B1']='World'
sheet.title='Test1'
headings = ["NAME", "ID", "EMPLOYEE NAME", "NUMBER", "START TIME", "STOP TIME"]
sheet.append(headings)
for x in range(3,10):
	for y in range(1,10):
		interger_i+=1
		sheet.cell(row=x,column=y,value=interger_i)
sheet['C1']='=sum(A3:I3)'

sheet=workbook.create_sheet(table_name[1],1)
sheet['A1']='Hello'

sheet=workbook.create_sheet('202505',2)
sheet['A1']='Hello'
sheet.cell(row=2,column=1).value=10

sheet=workbook.worksheets[2]
sheet['B1']='World'

# Create a few styles
bold_font = Font(bold=True)
big_red_text = Font(color="00FF0000", size=20)
center_aligned_text = Alignment(horizontal="center")
double_border_side = Side(border_style="double")
square_border = Border(top=double_border_side,right=double_border_side,bottom=double_border_side,left=double_border_side)

sheet=workbook.worksheets[0]
sheet["A1"].font = bold_font
#sheet["A1"].patternfill =PatternFillProperties(fgClr="00FF0000")
sheet["B2"].font = big_red_text
sheet["B1"].alignment = center_aligned_text
sheet["A1"].border = square_border


#red_background = PatternFill(fgColor="00FF0000")
#diff_style = DifferentialStyle(fill=red_background)
#rule = Rule(type="expression", dxf=diff_style)
#rule.formula = ["$H1<3"]
#sheet.conditional_formatting.add("A1:O100", rule)
#sheet.conditional_formatting.add("A1", rule)


workbook.save('test.xlsx')
workbook.close()


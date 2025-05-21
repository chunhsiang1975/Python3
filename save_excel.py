# Writeten by Chun-Hsiang Chao
# Date:20250521
import openpyxl 
workbook=openpyxl.Workbook() #apt install python3-openpyxl
workbook.create_chartsheet('202505',1)
sheet=workbook.worksheets[0]
sheet['A1']='Hello'
listtitle=["姓名","電話"]
sheet.append(listtitle)
workbook.save('test.xlsx')

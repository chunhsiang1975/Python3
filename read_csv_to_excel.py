# Writeten by Chun-Hsiang Chao
# Date:20250625
import csv
import openpyxl
import os
from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter


def convertDate(date):
	str1=str(date)
	yearstr=str1[:3]
	realyear=str(int(yearstr)+1911)
	realdate=realyear+str1[4:6]+str1[7:9]
	return realdate



file_path = 'STOCK_DAY_1101_202506_utf-8.csv'



wb = Workbook()
sheet=wb.active
sheet.title = "CSV Data"  # Optional: Rename the sheet

with open(file_path, 'r', newline='\n', encoding='utf-8') as f:
	reader = csv.reader(f,delimiter=',')
	for row in reader:
		sheet.append(row)


for x in range(3,20):
  for y in range(4,9):
    sheet.cell(row=x,column=y,value=float(sheet.cell(row=x,column=y).value))




i=get_column_letter(1)
sheet.column_dimensions[i].width=25


#for row in range(3,19):
#    sheet["{}{}".format("E", row)].number_format = 'General'
#    sheet["{}{}".format("F", row)].number_format = 'General'
#    sheet["{}{}".format("G", row)].number_format = 'General'


#print(sheet.cell(row=3,column=5).value)



chart=LineChart()
data = Reference(sheet,min_row=2,max_row=19,min_col=4,max_col=7)
#data1=Reference(sheet,(3,5),(19,5))

chart.add_data(data, titles_from_data=True)
#chart.y_axis.scaling.min = 20  # Set the minimum value for the y-axis
#chart.y_axis.scaling.max = 30 # Set the maximum value for the y-axis



sheet.add_chart(chart, "K2")



wb.save('output_excel_file.xlsx')
wb.close()

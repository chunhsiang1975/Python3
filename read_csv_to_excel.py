# Writeten by Chun-Hsiang Chao
# Date:20250624
import csv
import openpyxl
import os

def convertDate(date):
	str1=str(date)
	yearstr=str1[:3]
	realyear=str(int(yearstr)+1911)
	realdate=realyear+str1[4:6]+str1[7:9]
	return realdate



file_path = 'STOCK_DAY_1101_202506_utf-8.csv'



wb = openpyxl.Workbook()
ws = wb.active
ws.title = "CSV Data"  # Optional: Rename the sheet

with open(file_path, 'r', newline='', encoding='utf-8') as f:
	reader = csv.reader(f)
	for row in reader:
		ws.append(row)

wb.save('output_excel_file.xlsx')


# Writeten by Chun-Hsiang Chao
# Date:20250731
import pandas as pd
import csv
import openpyxl
from openpyxl import Workbook

def convertDate(date):
  str1=str(date)
  yearstr=str1[:3]
  realyear=str(int(yearstr)+1911)
  realdate=realyear+str1[4:6]+str1[7:9]
  return realdate


file_path = 'STOCK_DAY_1101_202506_utf-8.csv'
temp_file_path='test.xlsx'

wb = Workbook()
sheet=wb.active

with open(file_path, 'r', newline='\n', encoding='utf-8') as f:
  reader = csv.reader(f,delimiter=',')
  for row in reader:
    sheet.append(row)

title=sheet.cell(row=1,column=1).value
s_list=list(title)
new_s_list=[]
for i in range(0,17):
  new_s_list.append(s_list[i])
new_s=''.join(new_s_list)
sheet.title = new_s

rows=sheet.max_row
records_number=rows-7



for x in range(3,records_number+3):
  sheet.cell(row=x,column=1,value=(convertDate(sheet.cell(row=x,column=1).value)))

int_columns=[2,3,9]
for x in range(3,records_number+3):
  for y in int_columns:
    s_list=list(sheet.cell(row=x,column=y).value)
    new_s_list=[]
    for z in range(0,len(s_list)):
      if (s_list[z]!=","):
          new_s_list.append(s_list[z])
    new_s=''.join(new_s_list)
    sheet.cell(row=x,column=y,value=int(new_s))


for x in range(3,records_number+3):
  for y in range(4,8):
    if(sheet.cell(row=x,column=y).value[0]=="X"):
      sheet.cell(x,y).value="0"
    sheet.cell(row=x,column=y,value=float(sheet.cell(row=x,column=y).value))


sheet.delete_rows(idx=records_number+3,amount=5)
sheet.delete_rows(idx=1)

wb.save(temp_file_path)
wb.close()



#df=pd.read_csv('STOCK_DAY_1101_202506_data.csv')
#df=pd.read_excel('test.xlsx',sheet_name='CSV_data')
df=pd.read_excel(temp_file_path)
print(len(df.columns))
print(df.loc[[0,11]])
print(df.columns)


#pd.options.display.max_rows=9999
#print(df.to_string()) 
#print(df.loc[[0,11]])
#print(df.head())
#print(df.tail())
#print(df.info())
#print(df)
#print(len(df.columns))
#print(df.columns.tolist())

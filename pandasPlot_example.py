# Writeten by Chun-Hsiang Chao
# Date:20250723
import pandas as pd
import csv
import openpyxl
from openpyxl import Workbook

file_path = 'STOCK_DAY_1101_202506_utf-8.csv'


wb = Workbook()
sheet=wb.active
sheet.title = "CSV Data"  # Optional: Rename the sheet

with open(file_path, 'r', newline='\n', encoding='utf-8') as f:
  reader = csv.reader(f,delimiter=',')
  for row in reader:
    sheet.append(row)

title=sheet.cell(row=1,column=1).value

rows=sheet.max_row
records_number=rows-7

row_number=2
column_name = [cell.value for cell in sheet[row_number]] 
#print(list(column_name)[1])


row_number=2
df=pd.DataFrame(list([cell.value for cell in sheet[row_number]]))
#df.columns=list(column_name)
print(len(df.columns))
for i in range(3,records_number+3):
  row_data = [cell.value for cell in sheet[i]] 
  #print(row_data)
  new_df=pd.DataFrame(row_data)
  df = pd.concat([df, new_df], ignore_index=True)

print(df.loc[[0,11]])
print(df.columns.tolist())


#df=pd.DataFrame(data,columns=column_name)
#df = pd.read_csv(file_path,sep=',')
#df = pd.read_csv('data.csv')

full_df = pd.read_csv(file_path,sep=',')
new_row_series=full_df.iloc[1]
print(new_row_series)



#pd.options.display.max_rows=9999
#print(df.to_string()) 
#print(df.loc[[0,11]])
#print(df.head())
#print(df.tail())
#print(df.info())
#print(df)
#print(len(df.columns))
#print(df.columns.tolist())

#a=[1,7,2]
#var=pd.Series(a,index=["x","y","z"])
#print(var)

#calories = {"day1": 420, "day2": 380, "day3": 390}
#var = pd.Series(calories)
#print(var)

#data = {
#  "calories": [420, 380, 390],
#  "duration": [50, 40, 45]
#}
#var = pd.DataFrame(data)
#print(var)

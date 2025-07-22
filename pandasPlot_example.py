# Writeten by Chun-Hsiang Chao
# Date:20250722
import pandas as pd

df = pd.read_csv('BWIBBU_d_ALL_20250701_utf-8.csv')
#df = pd.read_csv('data.csv')

#pd.options.display.max_rows=9999
#print(df.to_string()) 
#print(df.loc[[0,11]])
#print(df.head())
print(df.tail())
#print(df.info())

a=[1,7,2]
var=pd.Series(a,index=["x","y","z"])
print(var)

calories = {"day1": 420, "day2": 380, "day3": 390}
var = pd.Series(calories)
print(var)

data = {
  "calories": [420, 380, 390],
  "duration": [50, 40, 45]
}
var = pd.DataFrame(data)
print(var)

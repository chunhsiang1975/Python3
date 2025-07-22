# Writeten by Chun-Hsiang Chao
# Date:20250722
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
matplotlib.rc('font',family='Noto Serif JP')

datas = [[65,92,78,83,70], [90,72,76,93,56], [81,85,91,89,77], [79,53,47,94,80]]
indexs = ["賣場A", "賣場B", "賣場C", "賣場D"]
columns = ["雞肉", "水果", "蔬菜", "牛奶", "麵包"]
df = pd.DataFrame(datas, columns=columns,  index=indexs)
df.plot(kind='bar', title='賣場銷售額', fontsize=12)
plt.xticks(rotation=45)
plt.savefig("dataframe.png")
plt.show()

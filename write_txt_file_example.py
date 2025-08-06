# Writeten by Chun-Hsiang Chao
# Date:20250806
content='''Hello Python
中文測試
'''
f=open('file.txt','w',encoding='UTF-8')
f.write(content)
f.close()

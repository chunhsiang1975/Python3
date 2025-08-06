# Writeten by Chun-Hsiang Chao
# Date:20250806
f=open('file.txt','r',encoding='UTF-8')
#with open('file.txt','r',encoding='UTF-8') as f:
content=f.readlines()
print(content)
f.seek(0)
print(f.read(2))
for line in f:
    print(line,end="")

f.seek(0)
str1=f.read(3)
print(str1)
f.seek(0)
print(f.readline())
f.close()

# Writeten by Chun-Hsiang Chao
# Date:20250619
import re
from bs4 import BeautifulSoup

m=re.search(r'[0-9]+','temp123abcd') #搜尋數字，r表示這是正規表示式,不寫也可以
print(m)
m=re.match('[a-z]+','abcd1234gggg')
print(m)
if not m==None:
    print(m.group())  
    print(m.start()) 
    print(m.end())    
    print(m.span())   

m=re.findall('[a-z]+','abcd1234gggg')
print(m)

reobj=re.compile('[a-z]+')
m=reobj.findall('abcd1234gggg')
print(m)




html = """
<div class="content">
    E-Mail：<a href="mailto:mail@test.com.tw">mail</a><br>
    E-Mail2：<a href="mailto:mail2@test.com.tw">mail2</a><br>
    <ul class="price">定價：360元 </ul>
    <img src="http://test.com.tw/p1.jpg">
    <img src="http://test.com.tw/p2.jpg">
    <img src="http://test.com.tw/p3.png">
    <ul class="price">定價：420元 </ul>
    <img src="http://test.com.tw/p1.jpg">
    <img src="http://test.com.tw/p2.jpg">
    <img src="http://test.com.tw/p3.png">
</div>
"""

sp = BeautifulSoup(html, 'html.parser')

emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+',html)
for email in emails:
    print(email)

price=re.findall(r"[0-9]+",sp.select('.price')[0].text)[0] #價格
#price=re.findall(r"[\d]+",sp.select('.price')[1].text)[0] #價格
print(price)

regex=re.compile('.*\.jpg')
imglist=sp.find_all("img",{"src":regex})
for img in imglist:
    print(img["src"])

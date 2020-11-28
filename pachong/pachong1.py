import requests        #导入requests包
from bs4 import BeautifulSoup
url = 'http://www.cntour.cn/'

strhtml = requests.get(url)        #Get方式获取网页数据
soup=BeautifulSoup(strhtml.text,'html.parser')
print(soup.select("a[href]"))
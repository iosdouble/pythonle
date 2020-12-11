# 爬取图片

import time
import requests
from bs4 import BeautifulSoup

class Aaa():
    headers = {
        "Cookie": "__cfduid=db706111980f98a948035ea8ddd8b79c11589173916",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36"
    }

    def get_cookies(self):
        url = "http://www.netbian.com/"
        response = requests.get(url=url)
        self.headers ={
            "Cookie":"__cfduid=" + response.cookies["__cfduid"],
            "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36"
        }

    # 获取图片列表
    def get_image_list(self,url):

        try:
            response = requests.get(url=url,headers=self.headers)
            response.encoding = 'gbk'
            soup = BeautifulSoup(response.text,'lxml')
            li_list = soup.select("#main > div.list > ul > li")
            for li in li_list:
                href = "http://www.netbian.com" + li.select_one("a").attrs["href"]
                self.get_image(href)
        except:
            self.get_cookies()


    def get_image(self,href):
        try:
            response = requests.get(url=href,headers=self.headers)
            response.encoding = 'gbk'
            soup = BeautifulSoup(response.text, 'lxml')
            image_href = "http://www.netbian.com" + soup.select_one("#main > div.endpage > div > p > a").attrs["href"]
            self.get_image_src(image_href)
        except:
            self.get_cookies()


    def get_image_src(self,href):
        try:
            response = requests.get(url=href,headers=self.headers)
            response.encoding = 'gbk'
            soup = BeautifulSoup(response.text, 'lxml')
            src = soup.select("img")[1].attrs["src"]
            self.download_image(src)
        except:
            self.get_cookies()

    # 下载图片
    def download_image(self,image_src):
        try:
            title = str(time.time()).replace('.', '')
            image_path = "static/images/" + title + ".png",
            image_path = list(image_path)
            response = requests.get(image_src,headers=self.headers)
            # 获取的文本实际上是图片的二进制文本
            img = response.content
            # 将他拷贝到本地文件 w 写  b 二进制  wb代表写入二进制文本
            with open(image_path[0],'wb' ) as f:
                f.write(img)
        except:
            self.get_cookies()


if __name__ == '__main__':
    aaa = Aaa()
    aaa.get_cookies()
    for i in range(2,3):
        url = "http://www.netbian.com/meinv/index_{}.htm".format(i)
        aaa.get_image_list(url)

import requests
from lxml import etree
import time

def get_article_list():
    base_page_link = 'https://blog.csdn.net/nihui123/article/list/'
    list_url = [];
    for i in range(0,6):
        real_page_link = base_page_link + str(i)+'?orderby=ViewCount'

        ##https://blog.csdn.net/nihui123/article/list/4?orderby=ViewCount
        req_headers = {
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.108 Safari/537.36'
        }
        # 提取本页所有文章链接
        resp = requests.get(real_page_link, headers=req_headers)
        if resp.status_code == requests.codes.ok:
            html = etree.HTML(resp.text)
            article_links = html.xpath("//h4/a/@href")
            print(article_links)
            # list_url = list_url+ article_links;

    return list_url

if __name__ == '__main__':
    print(get_article_list())
    print(get_article_list().__len__())

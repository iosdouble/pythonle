# 导入requests模块
import requests
import time

# 接口地址
# url = 'http://62.234.28.183:8081/jars'
url = 'http://123.206.55.168:8888/jsq'
# 请求的参数数据
# 发送请求
while 1:
    r = requests.get(url)
    print(r.text)
    time.sleep(4)


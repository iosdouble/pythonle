import time, datetime

# 字符类型的时间
tss1 = '2020-1-10 23:40:00'
# 转为时间数组
timeArray = time.strptime(tss1, "%Y-%m-%d %H:%M:%S")
print(timeArray)
# timeArray可以调用tm_year等
print(timeArray.tm_yday)
# 转为时间戳
timeStamp = int(time.mktime(timeArray))
print(timeStamp)

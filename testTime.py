#coding:utf-8
import datetime
import requests
import time
import http.client

now = datetime.datetime.now()
d1 = datetime.datetime(2017, 8, 21,14,8,0)
print now,d1
print (now-d1).min
print (now-d1).seconds

# response = requests.get("http://route.showapi.com/630-1?showapi_appid=myappid&showapi_sign=mysecret", verify=False)
# response.raise_for_status()
# print response.text




def get_webservertime(host):
    conn = http.client.HTTPConnection(host)
    conn.request("GET", "/")
    r = conn.getresponse()
    # r.getheaders() #获取所有的http头
    ts = r.getheader('date')  # 获取http头date部分
    # print(ts)

    # 将GMT时间转换成北京时间
    ltime = time.strptime(ts[5:25], "%d %b %Y %H:%M:%S")
    # print(ltime)
    ttime = time.localtime(time.mktime(ltime) + 8 * 60 * 60)
    # print(ttime)
    dat = "%u-%02u-%02u" % (ttime.tm_year, ttime.tm_mon, ttime.tm_mday)
    tm = "%02u:%02u:%02u" % (ttime.tm_hour, ttime.tm_min, ttime.tm_sec)
    print (dat, tm)


get_webservertime("www.baidu.com")
get_webservertime("www.taobao.com")
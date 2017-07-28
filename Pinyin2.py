#coding:utf-8


import urllib2
import json
import getpass
appid = "42965"
sign='2146bebb858743a0863618e59005342c'
url = "http://route.showapi.com/99-38?showapi_appid=${appid}&content=${content}&showapi_sign=${sign}"
url= url.replace("${appid}",appid).replace("${sign}",sign).replace("${content}","你好:周创:张三:李四:王五")
print url
req = urllib2.Request(url)
print req
res_data = urllib2.urlopen(req)
res = res_data.read()
print res
json_res = json.loads(res)
print (json_res["showapi_res_body"]["data"]).replace(" ","")
print getpass.getuser()
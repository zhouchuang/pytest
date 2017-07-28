#coding:utf-8
import json
import request
import parse

print('send data....')
showapi_appid = "42965"  # 替换此值
showapi_sign = "2146bebb858743a0863618e59005342c"  # 替换此值
url = "http://route.showapi.com/99-38"
send_data = parse.urlencode([
    ('showapi_appid', showapi_appid)
    , ('showapi_sign', showapi_sign)
    , ('content', "你好:是你吗")

])

req = request.Request(url)
with request.urlopen(req, data=send_data.encode('utf-8')) as f:
    print('Status:', f.status, f.reason)
    str_res = f.read().decode('utf-8')
    print('str_res:', str_res)
    json_res = json.loads(str_res)
    print ('json_res data is:', json_res)



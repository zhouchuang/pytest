#coding:utf-8
import requests
import json
import demjson
import sys


reload(sys)
sys.setdefaultencoding('utf-8')
url="http://www.wdzj.com/wdzj/html/json/dangan_search.json"
try:
    kv={'pageNum':'2','pageSize':'10'}
    r=requests.get(url,params=kv)
    print(r.request.url)
    r.raise_for_status()
    #print(len(r.text))
    print(r.text)
    compressdata = json.loads(r.text)
    datas = demjson.decode(r.text)
    # print json.dumps(datas, indent=4, sort_keys=False, ensure_ascii=False)
    # for i in range(0,100):
    #     item = compressdata[i]
    #     print item["platIconUrl"];
    #     html = requests.get(item["platIconUrl"])
    #     with open("E:\spiderSource\wdzj\logo\\"+item["allPlatPin"]+'.'+item["platIconUrl"].split('.')[-1], 'wb') as file:
    #         file.write(html.content)
    jsonstr = json.dumps(datas, indent=4, sort_keys=False, ensure_ascii=False)
    print jsonstr
    with open('E:\\spiderSource\\wdzj\\top.txt','w') as file:
        #file.write(json.dumps(datas, indent=4, sort_keys=False, ensure_ascii=False).decode("utf-8"))
        #file.write(json.dumps(datas, indent=4, sort_keys=False, ensure_ascii=True))
        file.write(jsonstr)
except Exception ,e:
    print("爬取失败")
    print e
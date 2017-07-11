"""
题目：暂停一秒输出，并格式化当前时间。
程序分析：无。
"""

import  time
print(time.strftime("%y-%m-%d %H:%M:%S",time.localtime(time.time())))
time.sleep(1)
print(time.strftime("%y-%m-%d %H:%M:%S",time.localtime(time.time())))



myD   = {1:'A',2:'B'}
for key ,value in dict.items(myD):
    print(key,value)
    time.sleep(1)

"""
题目：请输入星期几的第一个字母来判断一下是星期几，如果第一个字母一样，则继续判断第二个字母。
程序分析：用情况语句比较好，如果第一个字母一样，则判断用情况语句或if语句判断第二个字母。。
"""
import re

def autoInput(pre):
    list = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    day = pre+input("请输入下一个字母：")
    day = day.lower()
    mlist = []
    for d in list:
        if re.match(day, d.lower()):
            mlist.append(d)

    if len(mlist) == 1:
        print(mlist[0])
        return
    elif len(mlist)>1:
        print("当前输入为：{}".format(day))
        autoInput(day)
    else:
        print("你是个大傻逼")
        return

autoInput("")
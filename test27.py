"""
题目：利用递归函数调用方式，将所输入的5个字符，以相反顺序打印出来。
程序分析：无。
"""


def reprint(s):

    print(s[len(s)-1])
    if(len(s)==1):
        return
    else:
        reprint(s[0:len(s)-1])

s=input("请输入一个单词：")
reprint(s)
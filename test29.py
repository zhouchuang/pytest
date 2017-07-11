"""
题目：给一个不多于5位的正整数，要求：一、求它是几位数，二、逆序打印出各位数字。
程序分析：学会分解出每一位数。
"""

n = 0
def parse(num):
    print(num%10,end=",")
    if num<10:
        return
    else:
        parse((int)(num/10))

num = 12345
parse(num)
print("长度为：%d" % len(str(num)))
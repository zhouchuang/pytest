"""
题目：一个5位数，判断它是不是回文数。即12321是回文数，个位与万位相同，十位与千位相同。
程序分析：无。
"""
def huiwen(num):
    return num==int(str(num)[::-1])
print( huiwen(123321))

"""
题目：将一个列表的数据复制到另一个列表中。
程序分析：使用列表[:]。
"""

def copylist(a):
    b = a[:]
    return b

print(copylist([1,2,3,4]))

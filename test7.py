"""
题目：斐波那契数列。
程序分析：斐波那契数列（Fibonacci sequence），又称黄金分割数列，指的是这样一个数列：0、1、1、2、3、5、8、13、21、34、……。
"""

def fib(n):
    if n==1 :
        return [1]
    if n==2:
        return [1,1]
    fibs=[1,1]
    for i in range(2,n):
        fibs.append(fibs[-1]+fibs[-2])
    return fibs


def fibdigui(n):
    if n==1 or n==2:
        return 1
    return  fibdigui(n-1)+fibdigui(n-2)
print(fib(10))
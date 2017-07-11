"""
题目：练习函数调用。
"""

def test():
    print("hello")

def test1():
    for  i in range(0,3):
        test()

if __name__ == '__main__':
    test1()
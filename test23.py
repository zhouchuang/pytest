"""
题目：打印出如下图案（菱形）:
   *
  ***
 *****
*******
 *****
  ***
   *
程序分析：先把图形分成两部分来看待，前四行一个规律，后三行一个规律，利用双重for循环，第一层控制行，第二层控制列。
"""

for i in range(0,7):
    for k in range(0,abs(3-i)):
        print("",end=" ")
    for j in range(0,7):
        if  (2*((4-abs(i-3))-1)+1) >  j  :
            print("*",end="")


    print()
print("====================2=====================")

temp = "*******"
space = "       "
for i in range(0,7):
    a = abs(i-3)
    b = 7-abs(i-3)
    print(space[0:a],end="")
    print(temp[a:b],end="")
    print()
print("===================3======================")


def strabset(sStr1,ch,a,b):
    n = b-a
    sStr1 = sStr1[0:a] + n * ch + sStr1[b:]
    return sStr1


n = int(input("请输入n"))
for i in range(0,n):
    a = abs(i - int(n/2))
    #b = n - abs(i - int(n/2))
    #print(" "*a+"*"*(b-a))
    print(" "*a+"*"*(n-2*a))
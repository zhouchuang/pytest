"""
题目：有一分数序列：2/1，3/2，5/3，8/5，13/8，21/13...求出这个数列的前20项之和。
程序分析：请抓住分子与分母的变化规律
"""
fz=2
fm=1
sum = 0
for i in range(0,20):
    sum = sum + fz/fm
    tempfz = fz
    fz = fz+fm
    fm = tempfz

print(sum)

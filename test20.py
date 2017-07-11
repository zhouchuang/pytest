"""
题目：一球从100米高度自由落下，每次落地后反跳回原高度的一半；再落下，求它在第10次落地时，共经过多少米？第10次反弹多高？
"""


sum = 0.0
h = 100.0
for i in range(10):
    sum = sum + h
    h = h/2
    sum = sum + h
sum = sum-h
print("第10次落下共进过{0}".format(sum))
print("第10次落下高度{0}".format(h))
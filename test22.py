"""
题目：两个乒乓球队进行比赛，各出三人。甲队为a,b,c三人，乙队为x,y,z三人。已抽签决定比赛名单。有人向队员打听比赛的名单。a说他不和x比，c说他不和x,z比，请编程序找出三队赛手的名单。
"""

ta = ['a','b','c']
tb = ['x','y','z']
d = {}
for i  in  range(0,3):
    ua = ta[i]
    for j in range(0,3):
        ub = tb[j]
        if (ua == 'a' and ub == 'x') or (ua=='c' and (ub=='x' or ub=='z')):
            continue
        print( ua ,ub)
        d[ua]  = ub


print(d)
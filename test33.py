# -*- coding: UTF-8 -*-
"""
题目：按逗号分隔列表。
"""

L = [1,2,3,4,5]
sl = ''
sl = ','.join(str(n) for n in L)
print(sl)
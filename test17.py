"""
题目：输入一行字符，分别统计出其中英文字母、空格、数字和其它字符的个数。
程序分析：利用while语句,条件为输入的字符不为'\n'。
"""
def tongji():
    s  =  input("请随意输入字符")
    let=0
    spa=0
    dig=0
    oth=0
    for c in s :
        if c.isalpha():
            let +=1
        elif c.isspace():
            spa +=1
        elif c.isdigit():
            dig +=1
        else:
            oth +=1
    print("char = %d,space = %d,digit = %d,others = %d" % (let,spa,dig,oth))

tongji()
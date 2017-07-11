#题目：输入某年某月某日，判断这一天是这一年的第几天？
#程序分析：以3月5日为例，应该先把前两个月的加起来，然后再加上5天即本年的第几天，特殊情况，闰年且输入月份大于2时需考虑多加一天：

def addleap(year,month):
    if year % 4 ==0 and year%100!=0 and month > 2:
        return 1
    else:
        return 0
def getdayth():
    year  = int(input('请输入年份'))
    month = int(input("请输入月份"))
    day = int(input("请输入日期"))
    sum=0
    days = [31,30,28,30,31,30,31,31,30,31,30,31]
    for d in range(0,month-1):
        sum += days[d]

    sum += day
    sum += addleap(year,month)
    return sum

print(getdayth())
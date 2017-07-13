#coding:utf-8
import xlrd
import tkinter as tk
import win32ui
import win32api
import win32con
import xlwt
import os
import sys
import re
from tkinter import ttk

borders = xlwt.Borders()
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1
borders.bottom_colour=0x3A

alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中

fntdata = xlwt.Font()
fntdata.height=220
styledata=xlwt.XFStyle()
styledata.font = fntdata
styledata.alignment=alignment
styledata.borders = borders

fnt = xlwt.Font()
fnt.height=240
style = xlwt.XFStyle()
style.font = fnt
style.alignment = alignment
style.borders = borders


redfnt = xlwt.Font()
redfnt.height = 220
redfnt.colour_index=10
redstyle=xlwt.XFStyle()
redstyle.font = redfnt
redstyle.alignment = alignment
redstyle.borders = borders

badBG = xlwt.Pattern()
badBG.pattern = badBG.SOLID_PATTERN
badBG.pattern_fore_colour = 5
gbfnt = xlwt.Font()
gbfnt.height = 240
gbfnt.bold='true'
gbfnt.colour_index=0
styleTotalBg=xlwt.XFStyle()
styleTotalBg.font = gbfnt
styleTotalBg.alignment = alignment
styleTotalBg.borders = borders
styleTotalBg.pattern = badBG


badAllBG = xlwt.Pattern()
badAllBG.pattern = badBG.SOLID_PATTERN
badAllBG.pattern_fore_colour = 22
styleTotalAllBg=xlwt.XFStyle()
styleTotalAllBg.font = gbfnt
styleTotalAllBg.alignment = alignment
styleTotalAllBg.borders = borders
styleTotalAllBg.pattern = badAllBG

chineseM = ['一','二','三','四','五','六','七','八','九','十','十一','十二']


def importFile():
    dlg = win32ui.CreateFileDialog(1, None, None, 0, "Excel文件(*.xls)|*.xls||")  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('C:/Users/Public/Desktop')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    if filename:
        del datalist[:]
        pathtext.set(filename.decode('gbk'))

        data = xlrd.open_workbook(filename)
        table = data.sheet_by_index(0)
        nrows = table.nrows  # 行数
        colnames = table.row_values(0)  # 某一行数据
        names = set([])

        for rownum in range(1, nrows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                    if colnames[i]=='姓名':
                        names.add(row[i])
                datalist.append(app)
        for name in names:
            namelist.append(name)
        numberChosen['values'] = namelist;

def exportFile(event):
    if len(datalist) == 0:
        win32api.MessageBox(0, '请导入正确的考勤文件'.decode('utf-8'), '警告'.decode('utf-8'), win32con.MB_OK)
        return;
    event.widget.config(state="disabled")
    selectname = numberChosen.get()
    if selectname == '请选择':
        win32api.MessageBox(0, '请选择员工名称'.decode('utf-8'), '警告'.decode('utf-8'), win32con.MB_OK)
    else:
        del datatime[:]
        del datauser[:]
        # 过滤员工名称
        for entity in datalist:
            if entity['姓名'.decode('utf-8')] == selectname:
                datauser.append(entity)

        total = {'部门'.decode('utf-8'): '', '姓名'.decode('utf-8'): '', '日期'.decode('utf-8'): '', '周期'.decode('utf-8'): '', '上班时间'.decode('utf-8'): '', '上班情况'.decode('utf-8'): '', '下班时间'.decode('utf-8'): '', '下班情况'.decode('utf-8'): '合计',
                 '是否加班'.decode('utf-8'): 0, '费用'.decode('utf-8'): 0, '备注'.decode('utf-8'): ''}
        for entity in datauser:
            time = entity['下班时间'.decode('utf-8')].strip()
            if time and int(time[0:2]) >= 19:
                datatime.append(entity)

        total['是否加班'.decode('utf-8')] = len(datatime)
        total['费用'.decode('utf-8')] = len(datatime) * 20
        datatime.append(total)

        pathstr = os.path.realpath(sys.argv[0]).split('\\')
        filename = '加班餐费统计（' + selectname + '）.xls'
        # if os.path.exists(filename):
        #     os.remove(filename)
        wbk = xlwt.Workbook(encoding='utf-8')
        sheet = wbk.add_sheet('加班餐费统计（' + selectname + "）", cell_overwrite_ok=True)
        for i in range(0, len(titlelist)):
            col = sheet.col(i)
            col.width = 2500 * titleWidthList[i]
            sheet.write(0, i, titlelist[i], style)

        for i in range(0, len(datatime)):
            for j in range(0, len(titlelist)):
                col = sheet.col(j)
                col.width = 2500 * titleWidthList[j]
                if i == len(datatime) - 1:
                    value = datatime[i][titlelist[j]]
                    if value:
                        sheet.write((i + 1), j, value, styleTotalBg)
                else:
                    colstyle = titlelist[j] in dataDict.keys() and redstyle or styledata
                    value = titlelist[j] in dataDict.keys() and dataDict[titlelist[j]] or datatime[i][titlelist[j]]
                    if titlelist[j] == '下班情况':
                        if value.strip() == '':
                            value = '正常'
                    sheet.write((i + 1), j, value, colstyle)

        wbk.save(filename)
        msg.delete(0.0, 100.0)
        index = 1.0
        msg.insert(index, '----未打卡----\n\r')
        for entity in datauser:
            if ((entity['上班情况'.decode('utf-8')] == '未打卡' or entity['下班情况'.decode('utf-8')] == '未打卡') and entity['上班情况'.decode('utf-8')] != entity['下班情况'.decode('utf-8')]):
                index += 1
                str = entity['上班情况'.decode('utf-8')] == '未打卡' and '\t\t上班未打卡' or '\t\t下班未打卡'
                msg.insert(index, entity['日期'.decode('utf-8')] + '\t\t' + entity['周期'.decode('utf-8')] + str + '\t\t忘记打卡\n\r')

        index += 1
        msg.insert(index, '----迟到----\n\r')
        for entity in datauser:
            if entity['上班情况'.decode('utf-8')] == '迟到'.decode('utf-8'):
                index += 1
                msg.insert(index,
                           entity['日期'.decode('utf-8')] + '\t\t' + entity['周期'.decode('utf-8')] + '\t\t' + entity['上班时间'.decode('utf-8')] + '\t\t上班迟到\n\r')

        index += 1
        msg.insert(index, '----旷工----\n\r')
        for entity in datauser:
            if entity['上班情况'.decode('utf-8')] == '未打卡' and entity['下班情况'.decode('utf-8')] == '未打卡':
                index += 1
                msg.insert(index, entity['日期'.decode('utf-8')] + '\t\t' + entity['周期'.decode('utf-8')] + '\t\t旷工\n\r')

        index += 1
        msg.insert(index, '----节假日加班----\n\r')
        for entity in datauser:
            if not entity['上班情况'.decode('utf-8')] and not entity['下班情况'.decode('utf-8')] and (
                not entity['下班时间'.decode('utf-8')].strip() == '' or not entity['上班时间'.decode('utf-8')].strip() == ''):
                index += 1
                msg.insert(index,
                           entity['日期'.decode('utf-8')] + '\t\t' + entity['周期'.decode('utf-8')] + '\t\t' + entity['上班时间'.decode('utf-8')] + '\t\t' + entity[
                               '下班时间'.decode('utf-8')] + '\n\r')
    event.widget.config(state="normal")

def exportTotalFile(event):
    event.widget.config(state="disabled")
    currusername = ''
    del totallist[:]
    templist = []
    total = {}
    totalNum = 0
    for entity in datalist:
        time = entity['下班时间'.decode('utf-8')].strip()
        total = {'部门'.decode('utf-8'): '', '姓名'.decode('utf-8'): '', '日期'.decode('utf-8'): '', '周期'.decode('utf-8'): '', '上班时间'.decode('utf-8'): '', '上班情况'.decode('utf-8'): '', '下班时间'.decode('utf-8'): '',
                 '下班情况'.decode('utf-8'): '合计'.decode('utf-8'), '是否加班'.decode('utf-8'): 0, '费用'.decode('utf-8'): 0, '备注'.decode('utf-8'): ''}
        if time and int(time[0:2]) >= 19:
            totalNum += 1
            if currusername != entity['姓名'.decode('utf-8')]:
                if len(templist) > 0:
                    total['是否加班'.decode('utf-8')] = len(templist)
                    total['费用'.decode('utf-8')] = len(templist) * 20
                    templist.append(total)
                templist = []
                totallist.append(templist)
            templist.append(entity)
            currusername = entity['姓名'.decode('utf-8')]
    total['是否加班'.decode('utf-8')] = len(templist)
    total['费用'.decode('utf-8')] = len(templist) * 20
    templist.append(total)

    total = {'部门'.decode('utf-8'): '', '姓名'.decode('utf-8'): '', '日期'.decode('utf-8'): '', '周期'.decode('utf-8'): '', '上班时间'.decode('utf-8'): '', '上班情况'.decode('utf-8'): '', '下班时间'.decode('utf-8'): '',
             '下班情况'.decode('utf-8'): '总计', '是否加班'.decode('utf-8'): 0, '费用'.decode('utf-8'): 0, '备注'.decode('utf-8'): ''}
    total['是否加班'.decode('utf-8')] = totalNum
    total['费用'.decode('utf-8')] = totalNum * 20
    templist = []
    totallist.append(templist)
    templist.append(total)
    date= totallist[0][0]['日期'.decode('utf-8')]
    filename = chineseM[int(date[5:7])-1]+ '月份加班餐费统计汇总.xls'.decode('utf-8')
    # if os.path.exists(filename):
    #     os.remove(filename.encode('utf-8'))
    wbk = xlwt.Workbook(encoding='utf-8')
    sheet = wbk.add_sheet(chineseM[int(date[5:7])-1]+'月份加班餐费统计汇总'.decode('utf-8'), cell_overwrite_ok=True)
    index = 0
    for i in range(0, len(titlelist)):
        col = sheet.col(i)
        col.width = 2500 * titleWidthList[i]
        sheet.write(0, i, titlelist[i], style)
    for k in range(0, len(totallist)):
        list = totallist[k]
        for i in range(0, len(list)):
            index += 1
            for j in range(0, len(titlelist)):
                col = sheet.col(j)
                col.width = 2500 * titleWidthList[j]
                if i == len(list) - 1:
                    value = list[i][titlelist[j]]
                    if value:
                        if k == len(totallist) - 1:
                            sheet.write(index, j, value, styleTotalAllBg)
                        else:
                            sheet.write(index, j, value, styleTotalBg)
                else:
                    colstyle = titlelist[j] in dataDict.keys() and redstyle or styledata
                    value = titlelist[j] in dataDict.keys() and dataDict[titlelist[j]] or list[i][titlelist[j]]
                    if titlelist[j] == '下班情况':
                        if value.strip() == '':
                            value = '正常'
                    sheet.write(index, j, value, colstyle)

    sheet = wbk.add_sheet(chineseM[int(date[5:7]) - 1] + '月份人员加班餐补费'.decode('utf-8'), cell_overwrite_ok=True)
    index = 0
    for i in range(0, len(simpletitleWidthList)):
        col = sheet.col(i)
        col.width = 2500 * simpletitleWidthList[i]
        sheet.write(0, i, simpletitlelist[i], style)
    for k in range(0, len(totallist)-1):
        index += 1
        list = totallist[k]
        entity = list[0]
        listlen = len(list)-1
        entity['是否加班'.decode('utf-8')] = listlen
        entity['费用'.decode('utf-8')] = listlen * 20
        entity['日期'.decode('utf-8')] = date[0:7]
        for j in range(0, len(simpletitlelist)):
            col = sheet.col(j)
            col.width = 2500 * simpletitleWidthList[j]
            colstyle = simpletitlelist[j] in dataDict.keys() and redstyle or styledata
            value = entity[simpletitlelist[j]]
            sheet.write(index, j, value, colstyle)
    wbk.save(filename)
    msg.delete(1.0,100.0)
    msg.insert(1.0, '----生成'+chineseM[int(date[5:7])-1]+'月份人员加班餐补费----\n\r')
    totalmsg = "人数：{0}\t\t加班次数：{1}\t\t餐补费用：{2}".format(len(totallist),totalNum,totalNum*20);
    msg.insert(2.0, totalmsg)
    event.widget.config(state="normal")


def setCommon(event):
    choise.pack_forget()
    ttk.Label(win, text="文件路径").grid(column=0, row=0)
    ttk.Entry(win, width=50, textvariable=pathtext, state='readonly').grid(column=1, row=0)
    ttk.Button(win, text="导入考勤文件", command=importFile).grid(column=2, row=0)
    ttk.Label(win, text="员工姓名").grid(column=0, row=1)
    numberChosen['values'] = ('请选择')  # 设置下拉列表的值
    numberChosen.grid(column=1, row=1, columnspan=2)  # 设置其在界面中出现的位置  column代表列   row 代表行
    numberChosen.current(0)  # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
    msg.insert(1.0, "1:导入考勤文件\r\n2:选择员工姓名\r\n3:点击生成按钮\r\n\r\n加班补助文件会生成在程序当前目录下，控制台会打印相应考勤异常信息");
    msg.grid(column=0, row=3, columnspan=3)  # columnspan 个人理解是将3列合并成一列   也可以通过 sticky=tk.W  来控制该文本框的对齐方式
    action.grid(column=1, row=4)  # 设置其在界面中出现的位置  column代表列   row 代表行
    action.bind("<Button-1>", exportFile)


def setAdmin(event):
    choise.pack_forget()
    ttk.Label(win, text="文件路径").grid(column=0, row=0)
    ttk.Entry(win, width=50, textvariable=pathtext, state='readonly').grid(column=1, row=0)
    ttk.Button(win, text="导入考勤文件", command=importFile).grid(column=2, row=0)
    msg.insert(1.0, "1:导入考勤文件\r\n2:点击生成汇总按钮\r\n\r\n加班补助汇总文件会生成在程序当前目录下");
    msg.grid(column=0, row=2, columnspan=3)  # columnspan 个人理解是将3列合并成一列   也可以通过 sticky=tk.W  来控制该文本框的对齐方式
    action.grid(column=1, row=3)  # 设置其在界面中出现的位置  column代表列   row 代表行
    action.bind("<Button-1>", exportTotalFile)

def closeHandler():
    pathstr =  os.path.realpath(sys.argv[0]).split('\\')
    filename = pathstr[-1]
    #print filename
    cmd = "tasklist /fi \""+"imagename eq "+filename+"\" >> msg.txt";
    r = os.system(cmd)
    #print cmd


    # text = r.read()
    # r.close()

    f = open('msg.txt')
    text = f.read()
    f.close()
    text =  text.decode('gbk')
    if os.path.exists("msg.txt"):
        os.remove("msg.txt")
    result = re.findall('.*\\s+(\\d+)\\s+Console.*', text)
    for pid in result:
        try:
            os.popen("taskkill /f /pid:"+ str(pid))
        except Exception, e:
            print e.message
    win.destroy()
    sys.exit()

if __name__ == '__main__':

    reload(sys)
    sys.setdefaultencoding('utf-8')
    datalist = []
    datauser = []
    datatime = []
    namelist = []
    totallist = []
    titlelist = ['部门'.decode('utf-8'), '姓名'.decode('utf-8'), '日期'.decode('utf-8'), '周期'.decode('utf-8'), '上班时间'.decode('utf-8'), '上班情况'.decode('utf-8'), '下班时间'.decode('utf-8'), '下班情况'.decode('utf-8'), '是否加班'.decode('utf-8'), '费用'.decode('utf-8'), '备注'.decode('utf-8')]
    dataDict = {'是否加班'.decode('utf-8'): 1, '费用'.decode('utf-8'): 20, '备注'.decode('utf-8'): ''.decode('utf-8')}
    titleWidthList = [3, 1, 2, 1, 2, 1, 2, 1, 1, 1, 1]
    simpletitlelist= ['部门'.decode('utf-8'), '姓名'.decode('utf-8'), '日期'.decode('utf-8'),  '是否加班'.decode('utf-8'), '费用'.decode('utf-8'), '备注'.decode('utf-8')]
    simpletitleWidthList= [3, 1, 2,1, 1, 1]
    win = tk.Tk()
    win.protocol("WM_DELETE_WINDOW", closeHandler)
    w = 565
    h = 365
    win.title("餐补生成工具")  # 添加标题
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    win.geometry('%dx%d+%d+%d' % (w, h, x, y))

    pathtext = tk.StringVar()
    numberChosen = ttk.Combobox(win, width=63, textvariable=tk.StringVar(), state='readonly')
    action = ttk.Button(win, text="生成")  # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
    msg = tk.Text(win, height=20)


    choise = tk.Frame(win)
    choise.pack(fill="x")
    common = tk.Button(choise, text="普通人员", width=30, height=3, bg='orange')
    common.grid(rows=1, column=1, padx=165, pady=60)
    admin = tk.Button(choise, text="统计人员", width=30, height=3, bg='orange')
    admin.grid(rows=2, column=1, padx=165, pady=0)
    common.bind("<Button-1>", setCommon);
    admin.bind("<Button-1>", setAdmin);
    message = tk.Label(choise,text="Copyright ©深圳深信金融服务有限公司\r\n佳兆业金服 研发 周创\r\n版本号：v1.0",fg='orange').grid(rows=3,column=1,pady=20)

    win.mainloop()



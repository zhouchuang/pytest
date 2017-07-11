#!/usr/bin/python
# -*- coding: UTF-8 -*-


import xlrd
import tkinter as tk
import win32ui
import win32api
import win32con
import xlwt
import os
from tkinter import ttk




def showPanel(win):
    pathtext = tk.StringVar()
    numberChosen = ttk.Combobox(win, width=48, textvariable=tk.StringVar(),state='readonly' )

    msgtext = tk.StringVar()



    scr=tk.Text(win)
    scr.insert(1.0,"1:导入考勤文件\r\n2:选择员工姓名\r\n3:点击生成按钮\r\n\r\n加班补助文件会默认生成在桌面位置，控制台会打印相应考勤异常信息");


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

    #xlwt.easyxf('font: height 240, name Arial, colour_index black, bold off, italic on; align: wrap on, vert centre, horiz left;')
    #styleTotalBg = xlwt.easyxf('pattern: pattern solid, fore_colour gold; font: height 240, name Arial,colour_index black, bold on, italic off; align: wrap on, vert centre, horiz left;');
    #styleTotalBg=xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: bold on;align: wrap on, vert centre, horiz center;borders:'); # 80% like



    datalist = []
    datauser=[]
    datatime=[]
    namelist = []
    titlelist = ['部门','姓名','日期','周期','上班时间','上班情况','下班时间','下班情况','是否加班','费用','备注']
    dataDict={'是否加班':1,'费用':20,'备注':''}
    titleWidthList=[3,1,2,1,2,1,2,1,1,1,1]
    selectname = '请选择'
    def exportFile(event):
        event.widget.config(state="disabled")
        selectname = numberChosen.get()
        if selectname == '请选择':
            win32api.MessageBox(0, '请选择员工名称', '警告', win32con.MB_OK)
        else:
            datatime.clear()
            datauser.clear()
            # 过滤员工名称
            for entity in datalist:
                if entity['姓名'] == selectname:
                    datauser.append(entity)

            total={'部门':'','姓名':'','日期':'','周期':'','上班时间':'','上班情况':'','下班时间':'','下班情况':'合计','是否加班':0,'费用':0,'备注':''}
            for entity in datauser:
                time = entity['下班时间'].strip()
                if time and int(time[0:2]) >= 19:
                    datatime.append(entity)


            total['是否加班'] = len(datatime)
            total['费用'] = len(datatime)*20
            datatime.append(total)

            filename = '加班餐费统计（' + selectname + '）.xls'
            if os.path.exists(filename):
                os.remove(filename)
            wbk = xlwt.Workbook(encoding='utf-8')
            sheet = wbk.add_sheet('加班餐费统计（' + selectname + "）", cell_overwrite_ok=True)
            for i in range(0, len(titlelist)):
                col = sheet.col(i)
                col.width = 2500 * titleWidthList[i]
                sheet.write(0, i, titlelist[i],style)

            for i in range(0,len(datatime)):
                for j in range(0,len(titlelist)):
                    col = sheet.col(j)
                    col.width = 2500 * titleWidthList[j]
                    if i==len(datatime)-1:
                        value = datatime[i][titlelist[j]]
                        if value:
                            sheet.write((i + 1), j, value, styleTotalBg)
                    else:
                        colstyle = titlelist[j] in dataDict.keys() and redstyle or styledata
                        value = titlelist[j] in dataDict.keys()  and dataDict[titlelist[j]] or datatime[i][titlelist[j]]
                        if titlelist[j]=='下班情况':
                            if value.strip()=='':
                                value  = '正常'
                        sheet.write((i+1), j,  value  ,colstyle)

            wbk.save(filename)
            scr.delete(0.0, 100.0)
            index  =  1.0
            scr.insert(  index,'----未打卡----\n\r')
            for entity in datauser:
                if  ((entity['上班情况'] == '未打卡' or entity['下班情况'] == '未打卡') and entity['上班情况']!=entity['下班情况']) :
                    index += 1
                    str = entity['上班情况'] == '未打卡' and '\t\t上班未打卡' or  '\t\t下班未打卡'
                    scr.insert(index,entity['日期'] +'\t\t'+entity['周期']+str+'\t\t忘记打卡\n\r')

            index += 1
            scr.insert(index, '----迟到----\n\r')
            for entity in datauser:
                if entity['上班情况'] == '迟到':
                    index += 1
                    scr.insert(index,entity['日期'] +'\t\t'+entity['周期']+'\t\t' +entity['上班时间']+'\t\t上班迟到\n\r')

            index += 1
            scr.insert(index, '----旷工----\n\r')
            for entity in datauser:
                if entity['上班情况'] == '未打卡' and  entity['下班情况'] == '未打卡':
                    index += 1
                    scr.insert(index, entity['日期'] + '\t\t' + entity['周期'] + '\t\t旷工\n\r')

            index += 1
            scr.insert(index,'----节假日加班----\n\r')
            for entity in datauser:
                if not entity['上班情况'] and not entity['下班情况'] and ( not entity['下班时间'].strip()==''or   not entity['上班时间'].strip()=='' ):
                    index += 1
                    scr.insert(index,entity['日期'] + '\t\t' + entity['周期']+ '\t\t'  +entity['上班时间']+'\t\t'+entity['下班时间']+ '\n\r')


        event.widget.config(state="normal")


    def importFile():
        dlg = win32ui.CreateFileDialog(1, None, None,0, "Excel文件(*.xls)|*.xls||")  # 1表示打开文件对话框
        dlg.SetOFNInitialDir('C:/Users/Public/Desktop')  # 设置打开文件对话框中的初始显示目录
        dlg.DoModal()
        filename = dlg.GetPathName()  # 获取选择的文件名称

        if filename:
            datalist.clear()
            pathtext.set(filename)

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
                        if colnames[i] == '姓名':
                            names.add(row[i])
                    datalist.append(app)
            for name in names:
                namelist.append(name)
            numberChosen['values'] = namelist;

    ttk.Label(win, text="文件路径").grid(column=0, row=0)
    ttk.Entry(win,width=50, textvariable=pathtext,state='readonly').grid(column=1,row=0)
    ttk.Button(win,text="导入考勤文件",command=importFile).grid(column=2,row=0)

    ttk.Label(win, text="员工姓名").grid(column=0, row=1)

    numberChosen['values'] = ('请选择')   # 设置下拉列表的值
    numberChosen.grid(column=1, row=1)      # 设置其在界面中出现的位置  column代表列   row 代表行
    numberChosen.current(0)    # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值

    scr.grid(column=0, columnspan=3)        # columnspan 个人理解是将3列合并成一列   也可以通过 sticky=tk.W  来控制该文本框的对齐方式

    action = ttk.Button(win, text="生成")     # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
    action.grid(column=2, row=1)    # 设置其在界面中出现的位置  column代表列   row 代表行
    action.bind("<Button-1>",exportFile)

# menubar = tk.Menu(win)
# filemenu = tk.Menu(menubar, tearoff=0)
# filemenu.add_command(label="员工", command=changeCommonRole)
# filemenu.add_command(label="管理员", command=changeAdminRole)
# menubar.add_cascade(label='角色', menu=filemenu)
# win['menu'] = menubar

#win.mainloop()      # 当调用mainloop()时,窗口才会显示出来
if __name__ == '__main__':
    w = 565
    h = 371
    win = tk.Tk()
    win.iconbitmap('tool.ico')  # 加图标
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    win.geometry('%dx%d+%d+%d' % (w, h, x, y))
    win.title("餐补生成工具")    # 添加标题

    common = tk.Button(win,text="员工",width=30,height=3)
    common.grid(rows=1,column=1,padx=165, pady=60)
    admin = tk.Button(win,text="统计人员",width=30,height=3)
    admin.grid(rows=2, column=1, padx=165, pady=0)
    def setCommon(event):
        common.place_forget()
        showPanel(win)

    def setAdmin(event):
        admin.place_forget()
        showPanel(win)

    #启动有两按钮可以选择
    # common.grid(row=0,column=6,columnspan=6,rowspan=5,sticky="se")
    common.bind("<Button-1>",setCommon);
    # admin.grid(row=7,column=6,columnspan=6,rowspan=5,sticky="wn")
    admin.bind("<Button-1>",setAdmin);

    # common.pack( side='left')
    # admin.pack( side='right')
    win.mainloop()  # 当调用mainloop()时,窗口才会显示出来
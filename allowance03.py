# -*- coding: UTF-8 -*-
import xlrd
import tkinter as tk
import win32ui
import win32api
import win32con
import xlwt
import os
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

# 主窗
class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('餐补生成工具说')
        #self.iconbitmap('tool.ico')  # 加图标
        self.pathtext = pathtext = tk.StringVar()
        self.numberChosen = ttk.Combobox(self, width=63, textvariable=tk.StringVar(), state='readonly')
        self.action = ttk.Button(self, text="生成")  # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
        self.msg = tk.Text(self, height=20)
        self.datalist = []
        self.datauser = []
        self.datatime = []
        self.namelist = []
        self.totallist = []
        self.titlelist = ['部门', '姓名', '日期', '周期', '上班时间', '上班情况', '下班时间', '下班情况', '是否加班', '费用', '备注']
        self.dataDict = {'是否加班': 1, '费用': 20, '备注': ''}
        self.titleWidthList = [3, 1, 2, 1, 2, 1, 2, 1, 1, 1, 1]
        self.role = 'common'

        w = 565
        h = 350
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        self.choise = tk.Frame(self)
        self.choise.pack(fill="x")
        common = tk.Button(self.choise, text="普通人员", width=30, height=3, bg='orange')
        common.grid(rows=1, column=1, padx=165, pady=60)
        admin = tk.Button(self.choise, text="统计人员", width=30, height=3, bg='orange')
        admin.grid(rows=2, column=1, padx=165, pady=0)
        common.bind("<Button-1>", self.setCommon);
        admin.bind("<Button-1>", self.setAdmin);

    def setCommon(self,event):
        self.role = 'common'
        self.choise.pack_forget()
        ttk.Label(self, text="文件路径").grid(column=0, row=0)
        ttk.Entry(self, width=50, textvariable=self.pathtext, state='readonly').grid(column=1, row=0)
        ttk.Button(self, text="导入考勤文件", command=self.importFile).grid(column=2, row=0)
        ttk.Label(self, text="员工姓名").grid(column=0, row=1)
        self.numberChosen['values'] = ('请选择')  # 设置下拉列表的值
        self.numberChosen.grid(column=1, row=1, columnspan=2)  # 设置其在界面中出现的位置  column代表列   row 代表行
        self.numberChosen.current(0)  # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
        self.msg.insert(1.0, "1:导入考勤文件\r\n2:选择员工姓名\r\n3:点击生成按钮\r\n\r\n加班补助文件会默认生成在桌面位置，控制台会打印相应考勤异常信息");
        self.msg.grid(column=0, row=3, columnspan=3)  # columnspan 个人理解是将3列合并成一列   也可以通过 sticky=tk.W  来控制该文本框的对齐方式
        self.action.grid(column=1, row=4)  # 设置其在界面中出现的位置  column代表列   row 代表行
        self.action.bind("<Button-1>", self.exportFile)
    def setAdmin(self,event):
        self.role = 'admin'
        self.choise.pack_forget()
        ttk.Label(self, text="文件路径").grid(column=0, row=0)
        ttk.Entry(self, width=50, textvariable=self.pathtext, state='readonly').grid(column=1, row=0)
        ttk.Button(self, text="导入考勤文件", command=self.importFile).grid(column=2, row=0)
        self.msg.insert(1.0, "1:导入考勤文件\r\n2:点击生成汇总按钮\r\n\r\n加班补助文件会默认生成在桌面位置");
        self.msg.grid(column=0, row=2, columnspan=3)  # columnspan 个人理解是将3列合并成一列   也可以通过 sticky=tk.W  来控制该文本框的对齐方式
        self.action.grid(column=1, row=3)  # 设置其在界面中出现的位置  column代表列   row 代表行
        self.action.bind("<Button-1>", self.exportFile)

    def exportFile(self,event):
        if len(self.datalist) == 0:
            win32api.MessageBox(0, '请导入正确的考勤文件', '警告', win32con.MB_OK)
            return;
        event.widget.config(state="disabled")
        if  self.role=='common':
            selectname = self.numberChosen.get()
            if selectname == '请选择':
                win32api.MessageBox(0, '请选择员工名称', '警告', win32con.MB_OK)
            else:
                self.datatime.clear()
                self.datauser.clear()
                # 过滤员工名称
                for entity in self.datalist:
                    if entity['姓名'] == selectname:
                        self.datauser.append(entity)

                total = {'部门': '', '姓名': '', '日期': '', '周期': '', '上班时间': '', '上班情况': '', '下班时间': '', '下班情况': '合计',
                         '是否加班': 0, '费用': 0, '备注': ''}
                for entity in self.datauser:
                    time = entity['下班时间'].strip()
                    if time and int(time[0:2]) >= 19:
                        self.datatime.append(entity)

                total['是否加班'] = len(self.datatime)
                total['费用'] = len(self.datatime) * 20
                self.datatime.append(total)

                filename = '加班餐费统计（' + selectname + '）.xls'
                if os.path.exists(filename):
                    os.remove(filename)
                wbk = xlwt.Workbook(encoding='utf-8')
                sheet = wbk.add_sheet('加班餐费统计（' + selectname + "）", cell_overwrite_ok=True)
                for i in range(0, len(self.titlelist)):
                    col = sheet.col(i)
                    col.width = 2500 * self.titleWidthList[i]
                    sheet.write(0, i, self.titlelist[i], style)

                for i in range(0, len(self.datatime)):
                    for j in range(0, len(self.titlelist)):
                        col = sheet.col(j)
                        col.width = 2500 * self.titleWidthList[j]
                        if i == len(self.datatime) - 1:
                            value = self.datatime[i][self.titlelist[j]]
                            if value:
                                sheet.write((i + 1), j, value, styleTotalBg)
                        else:
                            colstyle = self.titlelist[j] in self.dataDict.keys() and redstyle or styledata
                            value = self.titlelist[j] in self.dataDict.keys() and self.dataDict[self.titlelist[j]] or self.datatime[i][self.titlelist[j]]
                            if self.titlelist[j] == '下班情况':
                                if value.strip() == '':
                                    value = '正常'
                            sheet.write((i + 1), j, value, colstyle)

                wbk.save(filename)
                self.msg.delete(0.0, 100.0)
                index = 1.0
                self.msg.insert(index, '----未打卡----\n\r')
                for entity in self.datauser:
                    if ((entity['上班情况'] == '未打卡' or entity['下班情况'] == '未打卡') and entity['上班情况'] != entity['下班情况']):
                        index += 1
                        str = entity['上班情况'] == '未打卡' and '\t\t上班未打卡' or '\t\t下班未打卡'
                        self.msg.insert(index, entity['日期'] + '\t\t' + entity['周期'] + str + '\t\t忘记打卡\n\r')

                index += 1
                self.msg.insert(index, '----迟到----\n\r')
                for entity in self.datauser:
                    if entity['上班情况'] == '迟到':
                        index += 1
                        self.msg.insert(index,
                                   entity['日期'] + '\t\t' + entity['周期'] + '\t\t' + entity['上班时间'] + '\t\t上班迟到\n\r')

                index += 1
                self.msg.insert(index, '----旷工----\n\r')
                for entity in self.datauser:
                    if entity['上班情况'] == '未打卡' and entity['下班情况'] == '未打卡':
                        index += 1
                        self.msg.insert(index, entity['日期'] + '\t\t' + entity['周期'] + '\t\t旷工\n\r')

                index += 1
                self.msg.insert(index, '----节假日加班----\n\r')
                for entity in self.datauser:
                    if not entity['上班情况'] and not entity['下班情况'] and (
                        not entity['下班时间'].strip() == '' or not entity['上班时间'].strip() == ''):
                        index += 1
                        self.msg.insert(index,
                                   entity['日期'] + '\t\t' + entity['周期'] + '\t\t' + entity['上班时间'] + '\t\t' + entity[
                                       '下班时间'] + '\n\r')
        else:
            currusername = ''
            self.totallist.clear()
            templist = []
            total = {}
            totalNum = 0
            for entity in self.datalist:
                time = entity['下班时间'].strip()
                total = {'部门': '', '姓名': '', '日期': '', '周期': '', '上班时间': '', '上班情况': '', '下班时间': '',
                         '下班情况': '合计', '是否加班': 0, '费用': 0, '备注': ''}
                if time and int(time[0:2]) >= 19:
                    totalNum += 1
                    if currusername != entity['姓名']:
                        if len(templist) > 0:
                            total['是否加班'] = len(templist)
                            total['费用'] = len(templist) * 20
                            templist.append(total)
                        templist = []
                        self.totallist.append(templist)
                    templist.append(entity)
                    currusername = entity['姓名']
            total['是否加班'] = len(templist)
            total['费用'] = len(templist) * 20
            templist.append(total)

            total = {'部门': '', '姓名': '', '日期': '', '周期': '', '上班时间': '', '上班情况': '', '下班时间': '',
                     '下班情况': '总计', '是否加班': 0, '费用': 0, '备注': ''}
            total['是否加班'] = totalNum
            total['费用'] = totalNum * 20
            templist = []
            self.totallist.append(templist)
            templist.append(total)
            date= self.totallist[0][0]['日期']
            filename = chineseM[int(date[5:7])-1]+ '月份加班餐费统计汇总.xls'
            if os.path.exists(filename):
                os.remove(filename)
            wbk = xlwt.Workbook(encoding='utf-8')
            sheet = wbk.add_sheet(chineseM[int(date[5:7])-1]+'月份加班餐费统计汇总', cell_overwrite_ok=True)
            index = 0
            for i in range(0, len(self.titlelist)):
                col = sheet.col(i)
                col.width = 2500 * self.titleWidthList[i]
                sheet.write(0, i, self.titlelist[i], style)
            for k in range(0, len(self.totallist)):
                list = self.totallist[k]
                for i in range(0, len(list)):
                    index += 1
                    for j in range(0, len(self.titlelist)):
                        col = sheet.col(j)
                        col.width = 2500 * self.titleWidthList[j]
                        if i == len(list) - 1:
                            value = list[i][self.titlelist[j]]
                            if value:
                                if k == len(self.totallist) - 1:
                                    sheet.write(index, j, value, styleTotalAllBg)
                                else:
                                    sheet.write(index, j, value, styleTotalBg)
                        else:
                            colstyle = self.titlelist[j] in self.dataDict.keys() and redstyle or styledata
                            value = self.titlelist[j] in self.dataDict.keys() and self.dataDict[self.titlelist[j]] or list[i][self.titlelist[j]]
                            if self.titlelist[j] == '下班情况':
                                if value.strip() == '':
                                    value = '正常'
                            sheet.write(index, j, value, colstyle)
            wbk.save(filename)
            self.msg.delete(0.0, 100.0)
            index = 1.0
            self.msg.insert(index, '----生成'+chineseM[int(date[5:7])-1]+'月份加班餐补汇总完成----\n\r')
            index +=1
            totalmsg = "人数：{0}\t\t加班次数：{1}\t\t餐补费用：{2}".format(len(self.totallist),totalNum,totalNum*20);
            self.msg.insert(index, totalmsg)
        event.widget.config(state="normal")

    def importFile(self):
        dlg = win32ui.CreateFileDialog(1, None, None, 0, "Excel文件(*.xls)|*.xls||")  # 1表示打开文件对话框
        dlg.SetOFNInitialDir('C:/Users/Public/Desktop')  # 设置打开文件对话框中的初始显示目录
        dlg.DoModal()
        filename = dlg.GetPathName()  # 获取选择的文件名称

        if filename:
            self.datalist.clear()
            self.pathtext.set(filename)

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
                    self.datalist.append(app)
            for name in names:
                self.namelist.append(name)
            self.numberChosen['values'] = self.namelist;





if __name__ == '__main__':
    app = MyApp()
    app.mainloop()

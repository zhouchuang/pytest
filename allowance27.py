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
import kaisaMailDriver
import urllib2
import json
import getpass
import kaisaIEDriver
import threading
import time as timer
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

config = {"username":"","password":"","url":"http://192.168.1.82/Portal/index.aspx","receiver":"","autoMail":'False',"autoOA":'False',"name":"","driverPath":"","auditor":"陈小龙"}

outLook = kaisaMailDriver.OutLook()


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
            namelist.append(name.strip())
        if len(namelist):
            numberChosen['values'] = namelist;
            if not config["name"]:
                currpinyinname = getCurrentChineseName(namelist)
                config["name"]  = currpinyinname
                updateConfig()
            else:
                currpinyinname = config["name"]
            numberChosen.set(currpinyinname or namelist[0])
        else:
            win32api.MessageBox(0, '请导入正确的考勤文件'.decode('utf-8'), '警告'.decode('utf-8'), win32con.MB_OK)




def exportFile(event):
    if len(datalist) == 0:
        win32api.MessageBox(0, '请导入正确的考勤文件'.decode('utf-8'), '警告'.decode('utf-8'), win32con.MB_OK)
        return;
    event.widget.config(state="disabled")
    selectname = numberChosen.get()
    if selectname == '请选择':
        win32api.MessageBox(0, '请选择员工名称'.decode('utf-8'), '警告'.decode('utf-8'), win32con.MB_OK)
        event.widget.config(state="normal")
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
        exportfilename = '加班餐费统计（' + selectname + '）.xls'
        # if os.path.exists(exportfilename):
        #     os.remove(exportfilename)
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

        wbk.save(exportfilename)
        msg.delete(0.0, 100.0)
        index = 1.0
        msg.insert(index, '----未打卡----\n\r')
        forgetSign = []
        for entity in datauser:
            if ((entity['上班情况'.decode('utf-8')] == '未打卡' or entity['下班情况'.decode('utf-8')] == '未打卡') and entity['上班情况'.decode('utf-8')] != entity['下班情况'.decode('utf-8')]):
                index += 1
                str = entity['上班情况'.decode('utf-8')] == '未打卡' and '上班未打卡' or '下班未打卡'
                msg.insert(index, entity['日期'.decode('utf-8')] + '\t\t' + entity['周期'.decode('utf-8')] + '\t\t'+str + '\t\t忘记打卡\n\r')
                forgetSign.append({"date":entity['日期'.decode('utf-8')],"day":entity['周期'.decode('utf-8')],"status":str.decode("utf-8")})
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
        date = datauser[0]['日期'.decode('utf-8')]
        index += 1
        msg.insert(index, "----邮件发送----\n\r")
        #发送邮件
        if (config['receiver']) and  config["autoMail"]=='True':

            # index += 1
            # msg.insert(index, "启动Outlook\n\r")
            # os.startfile("C:\Users\zhouchuang\Desktop\Microsoft Outlook 2010.lnk")
            # timer.sleep(2)
            #filename = chineseM[int(date[5:7]) - 1]
            # try:
            #     os.startfile("C:\Users\zhouchuang\Desktop\Microsoft Outlook 2010.lnk")
            #     outLook.send(config['receiver'], "加班餐费统计",os.path.realpath(sys.argv[0])[0:os.path.realpath(sys.argv[0]).rfind("\\")+1]+exportfilename ,chineseM[int(date[5:7]) - 1]+"月份加班餐费统计")
            #     index += 1
            #     msg.insert(index, "邮件已经发送到用户‘" + config["receiver"] + "’，邮件发送成功\n\r")
            # except Exception ,e:
            #     index += 1
            #     msg.insert(index, "邮件发送失败，请先打开Outlook后重试下\n\r")
            #     pass

            threading.Thread(target=startMail,args=(index, date,exportfilename,)).start()
        else:
            index += 1
            msg.insert(index, "没有开启自动发送邮件功能或者没有填写统计人员邮箱地址，请在 “设置>邮件设置” 中配置\n\r")
        #if(config['autoOA'])=='True'
        index += 1
        msg.insert(index, "----启动IE填写考勤异常信息----\n\r")
        if config["autoOA"] == 'True':
            threading.Thread(target=startOA,args=(index,int(date[0:4]),int(date[5:7]),forgetSign,config["auditor"],)).start()
        else:
            index += 1
            msg.insert(index, "没有开启自动处理考勤异常功能，请在 “设置>OA设置” 中配置\n\r")

    # win32api.MessageBox(0, '处理成功'.decode('utf-8'), '提示'.decode('utf-8'), win32con.MB_OK)
    # event.widget.config(state="normal")
def startMail(index,date,exportfilename):
    try:
        os.startfile("C:\Users\zhouchuang\Desktop\Microsoft Outlook 2010.lnk")
        outLook.send(config['receiver'], "加班餐费统计",
                     os.path.realpath(sys.argv[0])[0:os.path.realpath(sys.argv[0]).rfind("\\") + 1] + exportfilename,
                     chineseM[int(date[5:7]) - 1] + "月份加班餐费统计")
        index += 1
        msg.insert(index, "邮件已经发送到用户‘" + config["receiver"] + "’，邮件发送成功\n\r")
    except Exception, e:
        index += 1
        msg.insert(index, "邮件发送失败，请先打开Outlook后重试下\n\r")
        pass
def startOA(index,year,month,forgetSign,auditor):
    kaisaIEDriver.step_00_closeHandler("IEDriverServer.exe")
    index += 1
    status = kaisaIEDriver.step_01_loadIEDriver(config["driverPath"])
    msg.insert(index,status + "\n\r")
    if "失败" in status: return
    index += 1
    status = kaisaIEDriver.step_02_login(config["username"], config["password"], config["url"])
    msg.insert(index, status+ "\n\r")
    if "失败" in status: return
    index += 1
    status = kaisaIEDriver.step_03_locationFlowPage()
    msg.insert(index,status+ "\n\r")
    if "失败" in status: return
    index += 1
    status = kaisaIEDriver.step_04_getSessionID()
    msg.insert(index, status + "\n\r")
    if "失败" in status: return
    index += 1
    status = kaisaIEDriver.step_05_locationAllowancePage()
    msg.insert(index,  status+ "\n\r")
    if "失败" in status: return
    index += 1
    status = kaisaIEDriver.step_06_createAllowanceException()
    msg.insert(index,status+"\n\r")
    if "失败" in status: return
    index += 1
    status  = kaisaIEDriver.step_07_completeTable(year,month,forgetSign)
    msg.insert(index,status+"\n\r")
    if "失败" in status:return
    index += 1
    status = kaisaIEDriver.step_08_selectAuditor(auditor)
    msg.insert(index,status+"\n\r")


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
def resetButton(event):
    action.config(state="normal")
def closeMailHandler():
    mail.withdraw()
def closeApiHandler():
    showapi.withdraw()
def closeAccountHandler():
    account.withdraw()
def setMail():
    mail.update()
    mail.deiconify()
# def setApi():
#     showapi.update()
#     showapi.deiconify()
def setAccount():
    account.update()
    account.deiconify()
def about():
    win32api.MessageBox(0, '----------------------------------------V2.0----------------------------------------\n\r新增功能\n\r1:自动选择用户姓名\n\r程序会根据当前系统登录人作为用户名称（因为登陆名为姓名的拼音，所以需要在联网情况下才能正常访问汉字转拼音接口），省去了手动选择用户姓名操作\n\r\n\r'
                           '2:自动发送邮件\n\r如果你在 ‘设置>邮箱设置’ 中配置好了收件人邮箱（比如佳兆业金服统计人员邮箱为‘lvxm@kaisagroup.com’）并且启动了自动发送邮件，则程序会在生成餐补文件后自动发送给统计人员'.decode('utf-8'), '关于'.decode('utf-8'), win32con.MB_OK)
def quitSys():
    closeHandler()
def setCommon(event):
    choise.pack_forget()

    menubar = tk.Menu(win)
    filemenu = tk.Menu(menubar, tearoff=0)
    filemenu.add_command(label="邮箱设置", command=setMail)
    #filemenu.add_command(label="接口设置", command=setApi)
    filemenu.add_command(label="OA设置", command=setAccount)
    menubar.add_cascade(label="设置", menu=filemenu)

    helpmenu = tk.Menu(menubar, tearoff=0)
    helpmenu.add_command(label="关于", command=about)
    helpmenu.add_command(label="退出", command=quitSys)
    menubar.add_cascade(label="帮助", menu=helpmenu)

    win.config(menu=menubar)

    ttk.Label(win, text="文件路径").grid(column=0, row=0)
    ttk.Entry(win, width=50, textvariable=pathtext, state='readonly').grid(column=1, row=0)
    ttk.Button(win, text="导入考勤文件", command=importFile).grid(column=2, row=0)
    ttk.Label(win, text="员工姓名").grid(column=0, row=1)
    numberChosen['values'] = ('请选择')  # 设置下拉列表的值
    numberChosen.grid(column=1, row=1, columnspan=2)  # 设置其在界面中出现的位置  column代表列   row 代表行
    numberChosen.current(0)  # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
    numberChosen.bind("<<ComboboxSelected>>", resetButton)
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
    msg.config(height=23)
    action.grid(column=1, row=3)  # 设置其在界面中出现的位置  column代表列   row 代表行
    action.bind("<Button-1>", exportTotalFile)
def savemail():
    config["receiver"]=receiver.get()
    if autoMailValue.get()==0:
        config["autoMail"] = 'False'
    else:
        config["autoMail"] = 'True'
    mail.withdraw()
    updateConfig()
# def saveapi():
#     config["appid"]=appid.get()
#     config["secret"] = secret.get()
#     config["url"]= url.get()
#     showapi.withdraw()
#     updateConfig()
def saveOA():
    config["driverPath"] = driverPath.get()
    config["username"] = username.get()
    config["password"] = password.get()
    config["auditor"] = auditor.get()
    if autoException.get()==0:
        config["autoOA"] = 'False'
    else:
        config["autoOA"] = 'True'
    account.withdraw()
    updateConfig()
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
def initConfigFile():
    path = "C:\\Users\\" + getpass.getuser() + "\\.tools"
    if not os.path.exists(path):
        os.mkdir(path)
    #如果存在，则读取配置文件
    file = path+"\\allowance.properties"
    if  os.path.exists(file):
        pro = open(file)
        db = pro.read()
        pro.close()
        for pro in db.split("\n"):
            if '=' in pro:
                config[pro.split('=')[0].strip()] = pro.split('=')[1].strip()
    if not config["driverPath"]:
        config["driverPath"] = os.path.realpath(sys.argv[0])[0:os.path.realpath(sys.argv[0]).rfind("\\")+1]+"IEDriverServer.exe"
        updateConfig()

def updateConfig():
    profile = "C:\\Users\\" + getpass.getuser() + "\\.tools" + "\\allowance.properties"
    file_object = open(profile, 'w')
    list_of_text_strings  = ""
    for key in config:
        list_of_text_strings += (key+"="+str(config[key])+"\n")
    file_object.writelines(list_of_text_strings)
    file_object.close()


def getCurrentChineseName(namelist):
    url = "http://route.showapi.com/99-38"
    appid = "42965"
    secret = "2146bebb858743a0863618e59005342c"
    names = ""
    for name in namelist:
        names += name+":"
    names = names[0:len(names)-1]
    # appid = "42965"
    # sign = '2146bebb858743a0863618e59005342c'
    apiurl = "${url}?showapi_appid=${appid}&content=${content}&showapi_sign=${sign}"
    apiurl = apiurl.replace("${url}",url).replace("${appid}",appid).replace("${sign}",secret).replace("${content}", names)
    req = urllib2.Request(apiurl)
    res_data = urllib2.urlopen(req)
    res = res_data.read()
    json_res = json.loads(res)
    try:
        pinyins = (json_res["showapi_res_body"]["data"]).replace(" ", "").split(":")
    except:
        win32api.MessageBox(0, '拼音接口参数有误'.decode('utf-8'), '错误'.decode('utf-8'), win32con.MB_OK)
        return ""
    currentNamePinyin = getpass.getuser()
    for i in range(0,len(pinyins)):
        name = pinyins[i]
        if name  == currentNamePinyin:
            return namelist[i]
    return ""

    # p = Pinyin()
    # currentNamePinyin = getpass.getuser()
    # for name in namelist:
    #     #decodename = unicode(name,'utf-8')
    #     decodename = name.decode("utf-8")
    #     if currentNamePinyin == p.get_pinyin(decodename).replace("-",""):
    #         return name
    # return ""

    # currentNamePinyin = getpass.getuser()
    # for name in namelist:
    #     if currentNamePinyin == ''.join(lazy_pinyin(name.decode("utf-8"))):
    #         return name

if __name__ == '__main__':

    reload(sys)
    sys.setdefaultencoding('utf-8')
    initConfigFile()
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
    win.wm_attributes('-topmost', 1)
    w = 565
    h = 365
    win.title("餐补生成工具")
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    win.geometry('%dx%d+%d+%d' % (w, h, x, y))

    pathtext = tk.StringVar()
    numberChosen = ttk.Combobox(win, width=63, textvariable=tk.StringVar(), state='readonly')
    action = ttk.Button(win, text="生成")  # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
    msg = tk.Text(win, height=22)

    # 配置弹出框
    #邮箱设置
    mail = tk.Toplevel()
    mail.title("邮箱设置")
    mail.withdraw()
    mail.attributes('-toolwindow', True)
    mail.geometry('%dx%d+%d+%d' % (500, 200, (ws/2)-250,hs/2-100))
    tk.Label(mail, text='统计人员邮箱').grid(column=0, row=0)
    receiver = tk.StringVar()
    receiver.set(config["receiver"])
    ttk.Entry(mail, width=50,textvariable =receiver).grid(column=1, row=0)
    autoMailValue = tk.IntVar()
    autoMailValue.set(1 if str(config["autoMail"]) == "True" else 0)
    tk.Label(mail, text='自动发送邮件').grid(column=0, row=1)
    ttk.Radiobutton(mail, text="启动", variable=autoMailValue,  value=1).grid(column=1, row=1,sticky='w')
    ttk.Radiobutton(mail, text="关闭", variable=autoMailValue,  value=0).grid(column=1, row=2,sticky='w')
    ttk.Button(mail,text="保存",command=savemail).grid(column=1, row=3)
    mail.protocol("WM_DELETE_WINDOW", closeMailHandler)



    #账号设置
    account = tk.Toplevel()
    account.title("OA账号设置")
    account.withdraw()
    account.attributes('-toolwindow', True)
    account.geometry('%dx%d+%d+%d' % (500, 200, (ws / 2) - 250, hs / 2 - 100))
    tk.Label(account, text='驱动路径').grid(column=0, row=0)
    driverPath = tk.StringVar()
    driverPath.set(config["driverPath"])
    ttk.Entry(account, width=50, textvariable=driverPath).grid(column=1, row=0)

    tk.Label(account, text='登陆账号').grid(column=0, row=1)
    username = tk.StringVar()
    username.set(config["username"])
    ttk.Entry(account, width=50, textvariable=username).grid(column=1, row=1)

    tk.Label(account, text='登陆密码').grid(column=0, row=2)
    password = tk.StringVar()
    password.set(config["password"])
    ttk.Entry(account, width=50, textvariable=password).grid(column=1, row=2)

    tk.Label(account, text='审批人员').grid(column=0, row=3)
    auditor = tk.StringVar()
    auditor.set(config["auditor"])
    ttk.Entry(account, width=50, textvariable=auditor).grid(column=1, row=3)

    autoException = tk.IntVar()
    autoException.set(1 if str(config["autoOA"]) == "True" else 0)
    tk.Label(account, text='O A处理').grid(column=0, row=4)
    ttk.Radiobutton(account, text="启动", variable=autoException, value=1).grid(column=1, row=4, sticky='w')
    ttk.Radiobutton(account, text="关闭", variable=autoException, value=0).grid(column=1, row=5, sticky='w')

    ttk.Button(account, text="保存", command=saveOA).grid(column=1, row=6)
    account.protocol("WM_DELETE_WINDOW", closeAccountHandler)
    #接口设置
    # showapi = tk.Toplevel()
    # showapi.title("拼音接口设置")
    # showapi.withdraw()
    # showapi.attributes('-toolwindow', True)
    # showapi.geometry('%dx%d+%d+%d' % (500, 200, (ws / 2) - 250, hs / 2 - 100))
    #
    # tk.Label(showapi, text='url').grid(column=0, row=0)
    # url = tk.StringVar()
    # url.set(config["url"])
    # ttk.Entry(showapi, width=50, textvariable=url).grid(column=1, row=0)
    #
    # tk.Label(showapi, text='appid').grid(column=0, row=1)
    # appid = tk.StringVar()
    # appid.set(config["appid"])
    # ttk.Entry(showapi, width=50, textvariable=appid).grid(column=1, row=1)
    #
    # tk.Label(showapi, text='secret').grid(column=0, row=2)
    # secret = tk.StringVar()
    # secret.set(config["secret"])
    # ttk.Entry(showapi, width=50, textvariable=secret).grid(column=1, row=2)
    #
    # ttk.Label(showapi,text='该接口使用时间为1年，1年后需要重新订阅').grid(column=1,row=3)
    # ttk.Button(showapi, text="保存", command=saveapi).grid(column=1, row=4)
    # showapi.protocol("WM_DELETE_WINDOW", closeApiHandler)


    choise = tk.Frame(win)
    choise.pack(fill="x")
    common = tk.Button(choise, text="普通人员", width=30, height=3, bg='orange')
    common.grid(rows=1, column=1, padx=165, pady=60)
    admin = tk.Button(choise, text="统计人员", width=30, height=3, bg='orange')
    admin.grid(rows=2, column=1, padx=165, pady=0)
    common.bind("<Button-1>", setCommon);
    admin.bind("<Button-1>", setAdmin);
    message = tk.Label(choise,text="Copyright ©深圳深信金融服务有限公司\r\n佳兆业金服 研发 周创\r\n版本号：v2.0",fg='orange').grid(rows=3,column=1,pady=20)

    #outLook.sendSimple('635659050@qq.com','hello','hello world')
    win.mainloop()


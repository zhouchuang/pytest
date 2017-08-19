#coding:utf-8
import os
import time
import sys
import re
import getpass
from socket import socket, AF_INET, SOCK_DGRAM
import threading
import requests
import json
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import  random

w = 210
h = 60
btnW = 55
isSeach = False
handlerList = []
status="START"
count=0
browser=None

# property  = {"alipayAccount":"18607371493","alipayPassword":"1988204110zc","payOrder":"https://kyfw.12306.cn/otn//payOrder/init","loginurl":"https://kyfw.12306.cn/otn/login/init","username":"18607371493","password":"1988204110zc","ticket":"https://kyfw.12306.cn/otn/leftTicket/init","fromStationText":"深圳","fromStation":"","toStationText":"长沙","toStation":"","time":"2017-08-17"}
property  = {"payPassword":"","alipayAccount":"","alipayPassword":"","payOrder":"https://kyfw.12306.cn/otn//payOrder/init","loginurl":"https://kyfw.12306.cn/otn/login/init","username":"","password":"","ticket":"https://kyfw.12306.cn/otn/leftTicket/init","fromStationText":"深圳","fromStation":"","toStationText":"长沙","toStation":"","time":"2017-08-17"}
DesiredCapabilities.INTERNETEXPLORER['ignoreProtectedModeSettings'] = True

def openBrowser():
    global browser
    browser = webdriver.Ie()
    browser.maximize_window()
    printMsg("打开12306页面成功")

def login():
    global browser
    global handlerList
    browser.get(property["loginurl"])
    browser.find_element_by_id("username").send_keys(property["username"])
    browser.find_element_by_id("password").send_keys(property["password"])
    handlerList.append(browser.current_window_handle)
    printMsg("请选择正确验证码")

def checkout():
    global browser
    while True:
        time.sleep(1)
        if not ( browser.current_url == property["loginurl"] or browser.current_url==(property["loginurl"]+"#")) :
            break
    printMsg("登陆成功\n获取用户信息")
    while True:
        try:
            if not "登录" == browser.find_element_by_id("login_user").text:
                break
        except Exception,e:
            pass
        time.sleep(1)
    printMsg("获取成功\n打开购票页面")


def startScan():
    global isSeach
    global status
    global stopim
    global mainBtn
    global browser
    isSeach = True
    mainBtn.configure(image=stopim)
    status = "STOP"
    property["fromStationText"] = browser.find_element_by_id("fromStationText").get_attribute("value")
    property["toStationText"] = browser.find_element_by_id("toStationText").get_attribute("value")
    property["time"] = browser.find_element_by_id("train_date").get_attribute("value")
    printStation(property["fromStationText"]  + "-" +property["toStationText"])
    printTime(property["time"])
    #threading.Thread(target=useRequest).start()
    threading.Thread(target=autoSearch).start()

def useRequest():
    matchParam()
    scanTicket()

def scanTicket():
    global count
    global isSeach
    url = "https://kyfw.12306.cn/otn/leftTicket/query?leftTicketDTO.train_date=${time}&leftTicketDTO.from_station=${fromStation}&leftTicketDTO.to_station=${toStation}&purpose_codes=ADULT"
    url = url.replace("${time}",property["time"]).replace("${fromStation}",property["fromStation"]).replace("${toStation}",property["toStation"])
    while isSeach:
        response = requests.get(url, verify=False)
        response.raise_for_status()
        compressdata = json.loads(response.text)
        # for str in  compressdata["data"]["result"]:
            # print str
        count += 1
        print compressdata
        time.sleep(0.5)

def matchParam():
    response = requests.get("https://kyfw.12306.cn/otn/resources/js/framework/station_name.js?station_version=1.9024",verify=False)
    response.raise_for_status()
    stations = response.text
    property["fromStation"] = getCityCode(stations, property["fromStationText"])
    property["toStation"] = getCityCode(stations, property["toStationText"])

def getCityCode(stations,city):
    city = city.decode("utf-8")
    fromcitystart = stations.index("|" + city + "|") + len(city) + 2
    fromcityend = stations.index("|", fromcitystart + 1)
    return stations[fromcitystart:fromcityend]

def stopScan():
    global mainBtn
    global isSeach
    global status
    global scanim
    global count
    isSeach = False
    status = "SCAN"
    mainBtn.configure(image=scanim)
    printMsg("已经扫描"+str(count)+"次")


def toPayGatewayPage():
    global browser
    time.sleep(5)
    browser.find_element_by_id("payButton").click()
    browser.find_element_by_id("payButton").click()
    browser.find_element_by_id("payButton").click()
    toPayAlipayPage()

def toPayAlipayPage():
    global browser
    switchCurrentHandler()
    browser.find_elements_by_tag_name("img")[-1].click()
    toPayOrderByAlipay()

def toPayOrderByAlipay():
    global browser
    if property["alipayAccount"] and property["alipayPassword"]:
        browser.find_element_by_id("J_tLoginId").send_keys(property["alipayAccount"])
        browser.find_element_by_id("payPasswd_rsainput").send_keys(property["alipayPassword"])
        browser.find_element_by_id("J_submitBtn").click()
        toPayFinal()

def toPayFinal():
    global browser
    time.sleep(5)
    # browser.find_element_by_id("payPassword_rsainput").send_keys(property["payPassword"])
    elem = browser.find_element_by_css_selector('#payPassword_rsainput')
    browser.execute_script('''
        var elem = arguments[0];
        var value = arguments[1];
        elem.value = value;
    ''', elem, property["payPassword"])
    browser.find_element_by_id("J_authSubmit").click()

def autoSearch():
    global isSeach
    global browser
    global count
    while isSeach:
        # print browser.find_element_by_id("query_ticket").is_enabled()
        try:
            browser.find_element_by_id("query_ticket").click()
        except Exception,e:
            pass
        count += 1
        # if count%5==0 and (not (property["ticket"]) == browser.current_url):
        if count%5==0 and browser.current_url and str(property["payOrder"]) in str(browser.current_url):
            stopScan()
            threading.Thread(target=shake, args=("抢到票了啊",200,)).start()
            toPayGatewayPage()
            break


#浏览器控制切换到最新页面
def switchCurrentHandler():
    global browser
    for handler in browser.window_handles:
        if handler not in handlerList:
            browser.switch_to_window(handler)
            browser.maximize_window()
            handlerList.append(handler)
            break

def openTicket():
    global browser
    global scanim
    global status
    global mainBtn
    # browser.execute_script("window.open('"+property["ticket"]+"')")
    # switchCurrentHandler()
    browser.get(property["ticket"])
    printMsg("填写购票信息\n点击扫描按钮")
    mainBtn.configure(state="normal")
    mainBtn.configure(image=scanim)
    status = "SCAN"

def closeHandler():
    pathstr =  os.path.realpath(sys.argv[0]).split('\\')
    filename = pathstr[-1]
    cmd = "tasklist /fi \""+"imagename eq "+filename+"\" >> msg.txt";
    r = os.system(cmd)
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


def handler():
    global status

    if property["password"] and property["username"]:
        if "START"==status:
            flowHandler()
        elif "SCAN"==status:
            startScan()
        else:
            stopScan()
    else:
        openConfig()
        threading.Thread(target=shake, args=("请输入登录账号密码",40,)).start()


def flowHandler():
    global mainBtn
    mainBtn.configure(state="disabled")
    printMsg("已启动浏览器\n打开登陆页面")
    threading.Thread(target=flow).start()

def flow():
    openBrowser()
    login()
    checkout()
    # outstandingOrder()
    openTicket()
def outstandingOrder():
    global browser
    browser.get("https://kyfw.12306.cn/otn/queryOrder/initNoComplete")
    time.sleep(3)
    browser.find_element_by_id("continuePayNoMyComplete").click()
    toPayGatewayPage()

def printMsg(msg):
    canvas.delete("msg","station","time")
    canvas.create_text(btnW+2, h / 2,  # 使用create_text方法在坐标（302，77）处绘制文字
                       text=msg  # 所绘制文字的内容
                       , fill='#BB4444', tags="msg")  # 所绘制文字的颜色为灰色
def printStation(station):
    canvas.delete("station","msg")
    canvas.create_text(btnW+2, h / 3,
                       text=station
                       , fill='#BB4444', tags="station")
def printTime(time):
    canvas.delete("time","msg")
    canvas.create_text(btnW+2, h*2 / 3,
                       text=time
                       , fill='#BB4444', tags="time")


def mousedownHandler(event):
    global clickpx
    global clickpy
    clickpx = event.x
    clickpy = event.y
def releaseFrame(event):
    global clickpx
    global clickpy
    global lastW
    global lastH
    plusx = clickpx-event.x
    plusy = clickpy-event.y
    lastW -= plusx
    lastH -= plusy
    win.geometry('%dx%d+%d+%d' % (w, h, lastW, lastH))
def mouseMotionHandler(event):
    global clickpx
    global clickpy
    global lastW
    global lastH
    lastW -= clickpx - event.x
    lastH -= clickpy - event.y
    win.geometry('%dx%d+%d+%d' % (w, h, lastW, lastH))

def configureHandler(event):
    openConfig()

def openConfig():
    global config
    config.update()
    config.deiconify()

def saveConfig():
    global config
    global username
    global password
    global alipayAccount
    global alipayPassword
    property["username"] = username.get()
    property["password"] = password.get()
    property["alipayAccount"] = alipayAccount.get()
    property["alipayPassword"] = alipayPassword.get()
    property["loginurl"] = loginurl.get()
    property["ticket"] = ticket.get()
    property["payOrder"] = payOrder.get()
    config.withdraw()
    updateConfig()

def updateConfig():
    profile = "C:\\Users\\" + getpass.getuser() + "\\.tools" + "\\12306.properties"
    file_object = open(profile, 'w')
    list_of_text_strings  = ""
    for key in property:
        list_of_text_strings += (key+"="+str(property[key])+"\n")
    file_object.writelines(list_of_text_strings)
    file_object.close()

def initConfigFile():
    path = "C:\\Users\\" + getpass.getuser() + "\\.tools"
    if not os.path.exists(path):
        os.mkdir(path)
    #如果存在，则读取配置文件
    file = path+"\\12306.properties"
    if  os.path.exists(file):
        pro = open(file)
        db = pro.read()
        pro.close()
        for pro in db.split("\n"):
            if '=' in pro:
                property[pro.split('=')[0].strip()] = pro.split('=')[1].strip()

def shake(msg,delay):
    printMsg(msg)
    i = delay
    a = 8
    while i>0:
        time.sleep(0.03)
        win.geometry('%dx%d+%d+%d' % (w, h, lastW+random.randint(-a ,a ), lastH+random.randint(-a ,a )))
        i -=1
    win.geometry('%dx%d+%d+%d' % (w, h, lastW , lastH ))
def closeConfigHandler():
    config.withdraw()

if __name__=='__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    initConfigFile()
    global startim
    global stopim
    global scanim
    global mainBtn
    global ws
    global hs
    global lastW
    global lastH
    win = tk.Tk()
    win.overrideredirect(True)
    # win.attributes("-alpha", 0.8)  # 窗口透明度60 %
    win.protocol("WM_DELETE_WINDOW", closeHandler)
    win.wm_attributes('-topmost', 1)

    win.title("12306Click")
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    lastW = ws-w-10
    lastH = hs-h-50
    win.geometry('%dx%d+%d+%d' % (w, h,lastW ,lastH ))


    bg = Image.open("12306.bmp")
    bgim = ImageTk.PhotoImage(bg)

    start = Image.open("start.png")
    startim = ImageTk.PhotoImage(start)

    stop = Image.open("stop.png")
    stopim = ImageTk.PhotoImage(stop)

    scan = Image.open("scan.png")
    scanim = ImageTk.PhotoImage(scan)

    mainBtn = tk.Button(win,image=startim, width=btnW, height=btnW,bg="lightblue",command=handler)
    mainBtn.pack(side="left",expand="yes")


    canvas = tk.Canvas(win,
               width=w,
               height=h,
                bg="lightblue")
    canvas.configure(highlightthickness=0)
    canvas.pack(side="left")
    canvas.create_image((w-btnW)/2, (h)/2, image=bgim)  # 使用create_image将图片添加到Canvas组件中
    canvas.bind("<Button-1>", mousedownHandler)
    canvas.bind("<B1-Motion>", mouseMotionHandler)
    canvas.bind("<Double-Button-1>",configureHandler)

    printMsg("点击开始\n双击配置")
    # canvas.create_text(btnW, h/2,  # 使用create_text方法在坐标（302，77）处绘制文字
    #                    text='点击开始启动浏览器'  # 所绘制文字的内容
    #                    , fill='#BB4444',tags = "msg")  # 所绘制文字的颜色为灰色


    #canvas.create_oval(0+pad,0+pad,h-pad,h-pad, fill = "#BB4444",outline='lightblue')

    # ttk.Label(win,text="账号").grid(column=0, row=0, padx=10, pady=10)
    # username = tk.StringVar()
    # username.set("18607371493")
    # ttk.Entry(win,width=50, textvariable=username).grid(column=1,row=0)
    #
    # ttk.Label(win, text="密码").grid(column=0, row=1, padx=0, pady=0)
    # password = tk.StringVar()
    # password.set("1988204110zc")
    # ttk.Entry(win, width=50, textvariable=password,show = '*').grid(column=1, row=1)
    #
    # ttk.Button(win, text="打开12306",command=flow).grid(column=1,row=2, padx=10, pady=20)
    # searchText = tk.StringVar()
    # searchText.set("自动扫描（启动）")
    # searchBtn = ttk.Button(win, textvariable=searchText,command=searchHandler).grid(column=1, row=3)

    # 配置弹出框
    config = tk.Toplevel()
    config.title("账号设置")
    config.withdraw()
    config.attributes('-toolwindow', True)
    config.geometry('%dx%d+%d+%d' % (500, 300, (ws / 2) - 250, hs / 2 - 100))

    tk.Label(config, text='登录账号').grid(column=0, row=0)
    username = tk.StringVar()
    username.set(property["username"])
    ttk.Entry(config, width=50, textvariable=username).grid(column=1, row=0)

    tk.Label(config, text='登录密码').grid(column=0, row=1)
    password = tk.StringVar()
    password.set(property["password"])
    ttk.Entry(config, width=50, textvariable=password,show="*").grid(column=1, row=1)

    tk.Label(config, text='支付账号').grid(column=0, row=2)
    alipayAccount = tk.StringVar()
    alipayAccount.set(property["alipayAccount"])
    ttk.Entry(config, width=50, textvariable=alipayAccount).grid(column=1, row=2)

    tk.Label(config, text='支付密码').grid(column=0, row=3)
    alipayPassword = tk.StringVar()
    alipayPassword.set(property["alipayPassword"])
    ttk.Entry(config, width=50, textvariable=alipayPassword,show="*").grid(column=1, row=3)

    tk.Label(config, text='交易密码').grid(column=0, row=4)
    payPassword = tk.StringVar()
    payPassword.set(property["payPassword"])
    ttk.Entry(config, width=50, textvariable=payPassword,show="*").grid(column=1, row=4)


    tk.Label(config, text='登录地址').grid(column=0, row=5)
    loginurl = tk.StringVar()
    loginurl.set(property["loginurl"])
    ttk.Entry(config, width=50, textvariable=loginurl).grid(column=1, row=5)

    tk.Label(config, text='购票地址').grid(column=0, row=6)
    ticket = tk.StringVar()
    ticket.set(property["ticket"])
    ttk.Entry(config, width=50, textvariable=ticket).grid(column=1, row=6)

    tk.Label(config, text='付款地址').grid(column=0, row=7)
    payOrder = tk.StringVar()
    payOrder.set(property["payOrder"])
    ttk.Entry(config, width=50, textvariable=payOrder).grid(column=1, row=7)

    ttk.Button(config, text="保存", command=saveConfig).grid(column=1, row=8)
    config.protocol("WM_DELETE_WINDOW", closeConfigHandler)


    # autoMailValue = tk.IntVar()
    # autoMailValue.set(1 if str(config["autoMail"]) == "True" else 0)
    # tk.Label(mail, text='自动发送邮件').grid(column=0, row=1)
    # ttk.Radiobutton(mail, text="启动", variable=autoMailValue, value=1).grid(column=1, row=1, sticky='w')
    # ttk.Radiobutton(mail, text="关闭", variable=autoMailValue, value=0).grid(column=1, row=2, sticky='w')
    # ttk.Button(mail, text="保存", command=savemail).grid(column=1, row=3)



    win.mainloop()

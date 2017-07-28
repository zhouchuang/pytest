#coding:utf-8
import os
import time
import sys
import re
import getpass
from socket import socket, AF_INET, SOCK_DGRAM
import threading

from selenium import webdriver
# from selenium.common.exceptions import NoSuchElementException
# from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.wait import WebDriverWait

HOST = '127.0.0.1'
PORT = 11567
CLIENT_PORT=11568
BUFSIZE = 1024
ADDR = (HOST, PORT)
udpSerSock = socket(AF_INET, SOCK_DGRAM)
REMOTE_ADDR  = (HOST, CLIENT_PORT)
# chrome_driver = os.path.abspath(r"C:\Python27\chromedriver.exe")
# os.environ["webdriver.chrome.driver"] = chrome_driver
# driver = webdriver.Chrome(chrome_driver)
property  = {}

flushable = True

path = "C:\\Users\\" + getpass.getuser() + "\\.tools\\kaisa.properties"
print path
f = open(path)
db = f.read()
f.close()
print db
for pro in db.split("\n"):
    if '=' in pro:
        property[pro.split('=')[0].strip()] = pro.split('=')[1].strip()
print property


driverpath = str(property["driverPath"]).replace("\:", ":").replace("\\\\", "\\")
print driverpath
chrome_driver = os.path.abspath(r"C:\Python27\chromedriver.exe")
os.environ["webdriver.chrome.driver"] = chrome_driver
driver = webdriver.Chrome(chrome_driver)


def readProperties():
    path = "C:\\Users\\" + getpass.getuser() + "\\.tools\\kaisa.properties"
    print path
    f = open(path)
    db = f.read()
    f.close()
    print db
    for pro in db.split("\n"):
        if '=' in pro:
            property[pro.split('=')[0].strip()] = pro.split('=')[1].strip()
    print property
def openHome():
    driver.get("https://www.kaisafax.com/loan")

def investPage():
    js = 'window.open("https://www.kaisafax.com/loan/loanDetail?loanId=1579285122711142400");'
    driver.execute_script(js)

def userLogin():
    global flushable
    loginurl = property["host"] + property["loginUrl"]
    loginurl = loginurl.replace('\\', '')
    print  loginurl
    driver.get(loginurl)
    time.sleep(0.1)
    # 设置登陆账号密码跳转
    userinput = driver.find_element_by_id("userInput")
    userinput.send_keys(property["username"])
    pawdinput = driver.find_element_by_id("passwordInput")
    pawdinput.send_keys(property["password"])
    submitbt = driver.find_element_by_id("loginBtn");
    submitbt.click()
    time.sleep(0.5)
    driver.find_element_by_xpath("//p[@class='i_regbtn']/a").click()
    time.sleep(1)
    my_moneys = driver.find_element_by_xpath("//li[@id='user_available']/p[2]/em").text.replace('元','').replace(',','');
    print "mymoney:"+my_moneys
    my_money = float(str(my_moneys))
    property["my_money"] = my_money



    while True:
        time.sleep(600) #30分钟刷新一次
        # 切换登陆窗口
        if flushable == True:
            driver.get("https://www.kaisafax.com/account/")
            time.sleep(0.5)
            try:
                my_moneys = driver.find_element_by_xpath("//li[@id='user_available']/p[2]/em").text.replace('元','').replace(',','');
            except:
                # 设置登陆账号密码跳转
                driver.find_element_by_id("userInput").clear()
                driver.find_element_by_id("userInput").send_keys(property["username"])
                driver.find_element_by_id("passwordInput").clear()
                driver.find_element_by_id("passwordInput").send_keys(property["password"])
                driver.find_element_by_id("loginBtn").click()
                time.sleep(1)
                my_moneys = driver.find_element_by_xpath("//li[@id='user_available']/p[2]/em").text.replace('元','').replace( ',', '');
                pass
            print "mymoney:" + my_moneys
            my_money = float(str(my_moneys))
            property["my_money"] = my_money
            print  "refresh login page..."

def invest(loanId):
    # 跳转到投资页面
    invserurl = property["host"] + property["loanDetail"] + "=" +loanId
    invserurl = invserurl.replace('\\', '')
    print invserurl
    js = 'window.open("' + invserurl + '");'
    driver.execute_script(js)
    time.sleep(0.5)
    handles = driver.window_handles  # 获取当前窗口句柄集合（列表类型）
    driver.switch_to_window(handles[-1])
    # 查询还能投多少钱，填充钱
    try:
        inverstedmoney = driver.find_element_by_id("investAmountspan").get_attribute("val")
        investAmountInput = driver.find_element_by_id("investAmountInput")
        investmoney = "可投余额：{0}\t\t我的余额：{1}\t\t我的最大投资额度：{2}\t\t我的投资金额：{3}".decode("utf-8").format(float(inverstedmoney),  property["my_money"], int(property["maxInvestMoney"]),int(property["investMoney"]));
        print investmoney
        #如果我的投资金额为0或者没有具体数字，则不必须用固定金额投资int(4040.00/100)*100
        print int(property["investMoney"])
        print float(property["investMoney"])
        print property["my_money"]
        if int(property["investMoney"])==0:
            final_money = int(min(float(inverstedmoney), int(property["my_money"]/100)*100, int(property["maxInvestMoney"])))
            if final_money==0:
                print "投资金额不能为0".decode("utf-8")
                return
            else:
                investAmountInput.send_keys(str(final_money))
        else:
            if float(property["investMoney"]) > property["my_money"] or property["my_money"] <100:
                print "我的余额不足，请充值".decode("utf-8")
                return
            else:
                investAmountInput.send_keys(str(int(property["investMoney"])/100*100))

        driver.find_element_by_id("loanviewsbtn").click()
        time.sleep(2)
        #新交易方式，以后使用
        # driver.find_element_by_id("payPassword_rsainput").send_keys(property["payPassword"])
        # driver.find_element_by_id("transactionPwd_recharge").click()
        driver.find_element_by_class_name("form-unit").send_keys(property["payPassword"])
        driver.find_element_by_xpath("//div[@class='form-content']/div/a").click()
    except Exception,e:
        print  e
        print "满标了".decode("utf-8")
        time.sleep(5)
        launch("restart")
        pass



def userLoginAndInvsert(loanId):
    #登陆页面
    loginurl = property["host"] + property["loginUrl"]
    loginurl = loginurl.replace('\\', '')
    print  loginurl
    driver.get(loginurl)
    time.sleep(0.1)
    cookie = driver.get_cookies()
    print cookie
    #设置登陆账号密码跳转
    userinput = driver.find_element_by_id("userInput")
    userinput.send_keys(property["username"])
    pawdinput = driver.find_element_by_id("passwordInput")
    pawdinput.send_keys(property["password"])
    submitbt = driver.find_element_by_id("loginBtn");
    submitbt.click()
    my_moneys = driver.find_element_by_xpath("//div[@class='i_ac']/a").text.replace('元','').replace(',','');
    my_money = float(my_moneys+"")
    time.sleep(0.1)

    #跳转到投资页面
    invserurl = property["host"] + property["loanDetail"] + "=" + property["loanId"]
    invserurl = invserurl.replace('\\', '')
    print invserurl
    js = 'window.open("' + invserurl + '");'
    driver.execute_script(js)
    #切换窗口
    #print driver.current_window_handle  # 输出当前窗口句柄（百度）
    handles = driver.window_handles  # 获取当前窗口句柄集合（列表类型）
    #print handles  # 输出句柄集合

    for handle in handles:  # 切换窗口（切换到搜狗）
        if handle != driver.current_window_handle:
            #print 'switch to ', handle
            driver.switch_to_window(handle)
            #print driver.current_window_handle  # 输出当前窗口句柄（搜狗）
            break


    time.sleep(0.5)
    #查询还能投多少钱，填充钱
    inverstedmoney = driver.find_element_by_id("investAmountspan").get_attribute("val")
    investinput_tip = driver.find_element_by_id("investinput_tip").text
    investAmountInput = driver.find_element_by_id("investAmountInput")
    investmoney = "inverstedmoney：{0}\t\tmy_money：{1}\t\tmaxInvestMoney：{2}".format(float(inverstedmoney), my_money, int(property["maxInvestMoney"]));
    investAmountInput.send_keys( str(int(min(float(inverstedmoney),my_money,int(property["maxInvestMoney"])))))
    driver.find_element_by_id("loanviewsbtn").click()
    time.sleep(0.5)
    driver.find_element_by_id("payPassword_rsainput").send_keys(property["payPassword"])
    driver.find_element_by_id("transactionPwd_recharge").click()
    time.sleep(0.5)
    #投资完了后刷新登陆页面重新获取我的金额
    # driver.switch_to_window(handler['login'])
    # driver.refresh()
def launch(cmd):
    print  REMOTE_ADDR
    print  cmd
    udpSerSock.sendto(cmd,REMOTE_ADDR)
def receiver():
    global flushable
    print  ADDR
    while True:
        print 'wating for message...'
        data,addr = udpSerSock.recvfrom(BUFSIZE)
        print REMOTE_ADDR
        print data
        if data=="start":
            flushable = True
            readProperties()
        else:
            flushable = False
            invest(data)
        print "flushAble:{0}".decode("utf-8").format(flushable)
def closeUDP():
    udpSerSock.close()
def closeHandler():
    closeUDP()
    pathstr =  os.path.realpath(sys.argv[0]).split('\\')
    filename = pathstr[-1]
    cmd = "tasklist /fi \""+"imagename eq "+filename+"\"";
    r = os.popen(cmd)
    text = r.read()
    r.close()
    text =  text.decode('gbk')
    print text
    # if os.path.exists("msg.txt"):
    #     os.remove("msg.txt")
    result = re.findall('.*\\s+(\\d+)\\s+Console.*', text)
    for pid in result:
        try:
            print  "pid:{0}".decode("utf-8").format(pid)
            os.popen("taskkill /f /pid:"+ str(pid))
        except Exception, e:
            print e.message
    sys.exit()
def openBrowser():
    try:
        driverpath = str(property["driverPath"]).replace("\:",":").replace("\\\\","\\")
        print driverpath
        chrome_driver = os.path.abspath(r"C:\Python27\chromedriver.exe")
        os.environ["webdriver.chrome.driver"] = chrome_driver
        driver = webdriver.Chrome(chrome_driver)
        print driver
    except IOError, e:
        print  e
        print "驱动路径错误，请重新配置路径".decode("utf-8")
        pass
if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    udpSerSock.bind(ADDR)
    # readProperties()
    # openBrowser()
    #启动接收器
    threading.Thread(target=receiver).start()
    userLogin()




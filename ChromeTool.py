#coding:utf-8
import os
import time
import sys
import re
from socket import socket, AF_INET, SOCK_DGRAM
import threading

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait

HOST = '127.0.0.1'
PORT = 11567
BUFSIZE = 1024
ADDR = (HOST, PORT)




chrome_driver = os.path.abspath(r"C:\Python27\chromedriver.exe")
os.environ["webdriver.chrome.driver"] = chrome_driver
driver = webdriver.Chrome(chrome_driver)



property  = {}
handler = {}
f = open('C:\Users\zhouchuang\.tools\kaisa.properties')
db = f.read()
f.close()
print db
for pro in db.split("\n"):
    if '=' in pro:
        property[pro.split('=')[0].strip()] =pro.split('=')[1].strip()
print property


# username = "18607371493"
# password = "1988204110zc"
# username = "13597812114"
# password = "a123456"
# paypassword="519210"
# host = "http://localhost:8082/"
# login_url = "login/logout/"
# invest_url="loan/loanDetail?loanId="
# loan_id = "1565479121603878912"
# my_money=0
# max_money=10000
def openHome():
    driver.get("https://www.kaisafax.com/loan")

def investPage():
    js = 'window.open("https://www.kaisafax.com/loan/loanDetail?loanId=1579285122711142400");'
    driver.execute_script(js)

def userLogin():
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
    time.sleep(0.5)
    my_moneys = driver.find_element_by_xpath("//li[@id='user_available']/p[2]/em").text
    print "mymoney:"+my_moneys
    my_money = float(str(my_moneys))
    property["my_money"] = my_money
    loginHandler = driver.current_window_handle
    handler['login'] = loginHandler


    while True:
        time.sleep(600) #10分钟刷新一次
        # 切换登陆窗口
        driver.switch_to_window(handler['login'])
        driver.refresh()
        print  "fresh login page..."

def invest(loanId):
    #投资之前先查查自己有多少钱,切换到登陆页面
    driver.switch_to_window(handler['login'])
    #my_moneys = driver.find_element_by_xpath("//div[@class='i_ac']/a").text.replace('元', '').replace(',', '');
    my_moneys = driver.find_element_by_xpath("//li[@id='user_available']/p[2]/em").text
    my_money = float(str(my_moneys))
    property["my_money"] = my_money
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
    inverstedmoney = driver.find_element_by_id("investAmountspan").get_attribute("val")
    investinput_tip = driver.find_element_by_id("investinput_tip").text
    investAmountInput = driver.find_element_by_id("investAmountInput")
    investmoney = "可投余额：{0}\t\t我的余额：{1}\t\t我的最大投资额度：{2}".decode("utf-8").format(float(inverstedmoney),  property["my_money"],
                                                                                    int(property["maxInvestMoney"]));
    print investmoney
    investAmountInput.send_keys(str(int(min(float(inverstedmoney), property["my_money"], int(property["maxInvestMoney"])))))
    driver.find_element_by_id("loanviewsbtn").click()
    time.sleep(0.5)
    driver.find_element_by_id("payPassword_rsainput").send_keys(property["payPassword"])
    driver.find_element_by_id("transactionPwd_recharge").click()

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
    driver.switch_to_window(handler['login'])
    driver.refresh()

def receiver():
    udpSerSock = socket(AF_INET, SOCK_DGRAM)
    udpSerSock.bind(ADDR)
    print  "HOST：{0}\t\tPORT：{1}".decode("utf-8").format(HOST, PORT);
    while True:
        print 'wating for message...'
        data, addr = udpSerSock.recvfrom(BUFSIZE)
        #udpSerSock.sendto('[%s] %s' % (time.ctime(), data), addr)
        #print '...received from and retuned to:', addr
        print data
        # if data=='close':
        #     driver.close()
        #     udpSerSock.close()
        #     closeHandler()
        # else:
        invest(data)
    udpSerSock.close()

def closeHandler():
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

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    #启动接收器
    threading.Thread(target=receiver).start()
    userLogin()



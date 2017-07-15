#coding:utf-8
import os
import time
import sys
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait


chrome_driver = os.path.abspath(r"C:\Python27\chromedriver.exe")
os.environ["webdriver.chrome.driver"] = chrome_driver
driver = webdriver.Chrome(chrome_driver)



property  = {}
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

def userLoginAndInvsert():
    #登陆页面
    loginurl = property["host"] + property["loginUrl"]
    loginurl = loginurl.replace('\\', '')
    print  loginurl
    driver.get(loginurl)
    time.sleep(0.1)
    #cookie = driver.get_cookies()
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
    print investinput_tip
    investAmountInput = driver.find_element_by_id("investAmountInput")
    investmoney = "inverstedmoney：{0}\t\tmy_money：{1}\t\tmaxInvestMoney：{2}".format(float(inverstedmoney), my_money, int(property["maxInvestMoney"]));
    print  investmoney
    investAmountInput.send_keys( str(int(min(float(inverstedmoney),my_money,int(property["maxInvestMoney"])))))
    driver.find_element_by_id("loanviewsbtn").click()
    time.sleep(0.5)
    driver.find_element_by_id("payPassword_rsainput").send_keys(property["payPassword"])
    driver.find_element_by_id("transactionPwd_recharge").click()



if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    userLoginAndInvsert()
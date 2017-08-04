#coding:utf-8
import os
import time
import sys
import re
import getpass
from socket import socket, AF_INET, SOCK_DGRAM
import threading
import calendar
from selenium import webdriver
from selenium.common.exceptions import NoSuchWindowException

parentHander = []
IEdriver = "C:\Users\zhouchuang\Desktop\IEDriverServer.exe"


def closeHandler(filename):
    cmd = "tasklist /fi \""+"imagename eq "+filename+"\"";
    print cmd

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
    #sys.exit()

#关闭进程中的驱动
closeHandler("IEDriverServer.exe")

os.environ["webdriver.ie.driver"] = IEdriver
browser = webdriver.Ie(IEdriver)


def switchCurrentHandler():
    for handler in browser.window_handles:
        if handler  not in parentHander:
            browser.switch_to_window(handler)
            parentHander.append(handler)
            break


url = "http://192.168.1.82/Portal/index.aspx"
browser.get(url)
parentHander.append(browser.current_window_handle)

print  browser.current_url
browser.find_element_by_name("username").send_keys("zhouchuang")
browser.find_element_by_name("password").send_keys("1988204110zc@jzy.com")
browser.find_element_by_name("submit").click()


time.sleep(1)
switchCurrentHandler()
browser.find_element_by_link_text("流程申请").click()
#切换子窗口成功后关闭父窗口
#browser.close()

print  browser.current_url
# handles = browser.window_handles  # 获取当前窗口句柄集合（列表类型）
# for handle in handles:  # 切换窗口（切换到搜狗）
#     if handle != browser.current_window_handle:
#         browser.switch_to_window(handle)
#         break

switchCurrentHandler()
logined_url  =   browser.current_url
print logined_url
logined_url = logined_url.replace("?","&")
sessionidstr =[elem for elem in logined_url.split("&") if elem.startswith("Session")]
sessionID =  sessionidstr[0].split("=")[1]
print sessionID
time.sleep(1)

#切换并且点击人力类
# browser.switch_to_window(browser.window_handles[0])
browser.maximize_window()

browser.switch_to_frame("right")
browser.execute_script("selectTreeNode('373')")
time.sleep(1)

#点击考勤异常情况审批
browser.switch_to_frame("testFrm")
exceptionurl = "OpenDocList('${sessionID}','1999','02.考勤异常情况审批表',0,'')".decode("utf-8").replace("${sessionID}",sessionID)
browser.execute_script(exceptionurl)

#点击新建
time.sleep(1)
browser.switch_to_default_content()
browser.switch_to_frame("right")
#browser.find_element_by_id("Toolbar_ImgNew").click()
browser.execute_script("javascript:NewItem(1999);return false;")
# browser.close()
# browser.switch_to_window(browser.window_handles[0])
# browser.maximize_window()

#填充表单

# handles = browser.window_handles  # 获取当前窗口句柄集合（列表类型）
# for handle in handles:  # 切换窗口（切换到搜狗）
#     print handle
# browser.switch_to_window(handles[-1])
switchCurrentHandler()
y = 2017
m = 7
day_now  = calendar.month(y, m)
wday, monthRange = calendar.monthrange(y, m)
print wday,monthRange
time.sleep(1)
browser.find_element_by_id("ctl00_HfCalendar1_dateTextBox").send_keys( time.strftime("%Y/%m/%d"))
browser.find_element_by_id("ctl00_HfCalendar3_dateTextBox").send_keys(str(y)+"/"+str(m)+"/"+"1")
browser.find_element_by_id("ctl00_HfTextBox7").send_keys("9")
browser.find_element_by_id("ctl00_HfTextBox8").send_keys("00")
browser.find_element_by_id("ctl00_HfCalendar5_dateTextBox").send_keys(str(y)+"/"+str(m)+"/"+str(monthRange))
browser.find_element_by_id("ctl00_HfTextBox9").send_keys("18")
browser.find_element_by_id("ctl00_HfTextBox10").send_keys("00")
browser.find_element_by_id("ctl00_HfTextBox11").send_keys("0")
browser.find_element_by_id("ctl00_HfTextBox12").send_keys("忘记打卡".decode("utf-8"))
browser.execute_script("alert(_objID)")
#browser.find_element_by_id("Toolbar_ImgSend").click()

# time.sleep(5)
# audit = "http://192.168.1.18:9090/UI/FlowItem/ItemDetail.aspx?SessionID=${sessionID}&DataId=1999".replace("${sessionID}",sessionID);
# browser.execute_script("window.open('"+audit+"')")







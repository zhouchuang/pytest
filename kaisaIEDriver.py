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
from selenium.webdriver.common.keys import Keys

global browser
global sessionID
handlerList = []
def step_00_closeHandler(filename):
    cmd = "tasklist /fi \""+"imagename eq "+filename+"\"";
    print cmd
    r = os.popen(cmd)
    text = r.read()
    r.close()
    text =  text.decode('gbk')
    print text
    result = re.findall('.*\\s+(\\d+)\\s+Console.*', text)
    for pid in result:
        try:
            print  "pid:{0}".decode("utf-8").format(pid)
            os.popen("taskkill /f /pid:"+ str(pid))
        except Exception, e:
            print e.message

#浏览器控制切换到最新页面
def switchCurrentHandler():
    global browser
    for handler in browser.window_handles:
        if handler not in handlerList:
            browser.switch_to_window(handler)
            handlerList.append(handler)
            break
#加载驱动
def step_01_loadIEDriver(driverPath):
    global browser
    try:
        os.environ["webdriver.ie.driver"] = driverPath
        browser = webdriver.Ie(driverPath)
        return "加载IE驱动成功，尝试登陆OA系统"
    except Exception,e:
        print e
        return "IE驱动加载失败，找不到路径，请设置正确的IE驱动路径"

#登陆，并且加入到集合
def step_02_login(username,password,url):
    global browser
    try:
        browser.get(url)
        browser.find_element_by_name("username").send_keys(username)
        browser.find_element_by_name("password").send_keys(password)
        browser.find_element_by_name("submit").click()
        handlerList.append(browser.current_window_handle)
        return "登陆成功，打开流程申请页面"
    except Exception,e:
        print e
        return "登陆失败"

#定位到流程申请页面
def step_03_locationFlowPage():
    global browser
    try:
        browser.find_element_by_link_text("流程申请").click()
        switchCurrentHandler()
        browser.maximize_window()
        return "切换流程页面成功"
    except Exception,e:
        print e
        return "切换流程页面失败"

# 获取登陆后的sessionID
def step_04_getSessionID():
    global browser
    global sessionID
    try:
        time.sleep(0.5)
        logined_url = browser.current_url.replace("?", "&")
        sessionidstr = [elem for elem in logined_url.split("&") if elem.startswith("Session")]
        sessionID = sessionidstr[0].split("=")[1]
        return "获取sessionID成功"
    except Exception, e:
        print e
        return "获取sessionID失败"

#定位到考勤页面
def step_05_locationAllowancePage():
    global browser
    global sessionID
    try:
        time.sleep(0.5)
        browser.switch_to_frame("right")
        browser.execute_script("selectTreeNode('373')")
        browser.switch_to_frame("testFrm")
        time.sleep(0.5)
        exceptionurl = "OpenDocList('${sessionID}','1999','02.考勤异常情况审批表',0,'')".decode("utf-8").replace("${sessionID}",sessionID)
        browser.execute_script(exceptionurl)
        return "点击人力类>考勤异常情况审批成功"
    except Exception ,e :
        print  e
        return "点击人力类>考勤异常情况审批失败"

#新建考勤异常审批流
def step_06_createAllowanceException():
    global browser
    try:
        browser.switch_to_default_content()
        browser.switch_to_frame("right")
        time.sleep(1)
        #browser.find_element_by_id("Toolbar_ImgNew").click()
        browser.execute_script("javascript:NewItem(1999);return false;")
        return "打开新建考勤异常页面成功"
    except Exception, e:
        print  e
        return "打开新建考勤异常页面失败"

#填充异常信息
def step_07_completeTable(y,m,forgetSign):
    global browser
    try:
        switchCurrentHandler()
        time.sleep(0.5)
        wday, monthRange = calendar.monthrange(y, m)
        browser.find_element_by_id("ctl00_HfCalendar1_dateTextBox").send_keys(time.strftime("%Y/%m/%d"))
        browser.find_element_by_id("ctl00_HfCalendar3_dateTextBox").send_keys(str(y) + "/" + str(m) + "/" + "1")
        browser.find_element_by_id("ctl00_HfTextBox7").send_keys("9")
        browser.find_element_by_id("ctl00_HfTextBox8").send_keys("00")
        browser.find_element_by_id("ctl00_HfCalendar5_dateTextBox").send_keys(str(y) + "/" + str(m) + "/" + str(monthRange))
        browser.find_element_by_id("ctl00_HfTextBox9").send_keys("18")
        browser.find_element_by_id("ctl00_HfTextBox10").send_keys("00")
        browser.find_element_by_id("ctl00_HfTextBox11").send_keys("0")
        textValue = ""
        for d in forgetSign:
            textValue += (d["date"]+Keys.SPACE+Keys.SPACE+Keys.SPACE+d["day"]+Keys.SPACE+Keys.SPACE+Keys.SPACE+d["status"]+Keys.SPACE+Keys.SPACE+Keys.SPACE+"忘记打卡".decode("utf-8")+Keys.ENTER)
        browser.find_element_by_id("ctl00_HfTextBox12").send_keys(textValue)
        browser.find_element_by_id("Toolbar_ImgSend").click()
        #browser.execute_script("alert('只能帮你到这啦')")


        # objID = browser.execute_script("return _objID")
        # js = "window.open('UserQueryView2.aspx?SessionID=${sessionID}&Field=0&ObjID=${objID}&DataId=1999&ProcId=0&CurActiID=0".replace("${sessionID}",sessionID).replace("{objID}",objID)+"')"
        # print js
        # browser.execute_script(js)
        return "填充信息成功"
    except Exception ,e :
        print  e
        return "填充信息失败"

#选择审批人
def step_08_selectAuditor (auditor):
    global browser
    try:
        switchCurrentHandler()
        time.sleep(1)
        browser.switch_to_frame("userSelect")
        browser.find_element_by_id("TBSearchValue").send_keys(auditor.decode("utf-8"))
        browser.find_element_by_id("btnSearch").click()
        time.sleep(1)
        browser.find_element_by_xpath("//select[@id='listUsers']/option[1]").click()
        return "选择审批人成功"
    except Exception ,e :
        print  e
        return "选择审批人失败"

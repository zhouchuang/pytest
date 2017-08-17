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
property  = {"driverPath":"C:\Program Files (x86)\Mozilla Firefox\\firefox.exe","loginurl":"https://kyfw.12306.cn/otn/login/init","username":"18607371493","password":"1988204110zc","ticket":"https://kyfw.12306.cn/otn/leftTicket/init"}
profile_dir="C:\Users\zhouchuang\AppData\Roaming\Mozilla\Firefox\Profiles\\axqidngq.selenium"
profile = webdriver.FirefoxProfile(profile_dir)
browser = webdriver.Firefox(profile)


def login():
    browser.get(property["loginurl"])
    browser.find_element_by_id("username").send_keys(property["username"])
    browser.find_element_by_id("password").send_keys(property["password"])

def checkout():
    while True:
        time.sleep(1)
        if   not browser.current_url == property["loginurl"]:
            break
def alive():
    while True:
        time.sleep(600) #10分钟刷新一次
        browser.refresh()
        print "refresh..."
def openTicket():
    time.sleep(5)
    browser.get(property["ticket"])
    #browser.execute_script("alert('请填写相关订票需求后点击查询按钮')")
    while True:
        time.sleep(1)#一秒钟检查一次
        try:
            if browser.find_element_by_xpath(".//tbody[@id='queryLeftTable']/tr").size>0:
                break
        except Exception,e:
            print e
            break

    #browser.execute_script("alert('已保存相关需求，工具会自动扫描相关符合条件的车次')")
    print  browser.execute_script("return station_names")
    print  browser.find_element_by_id("fromStationText").text

    # browser.find_element_by_id("fromStationText").send_keys("深圳".decode("utf-8"))
    # browser.find_element_by_id("toStationText").send_keys("长沙".decode("utf-8"))
    # browser.find_element_by_id("train_date").send_keys("2017-08-11")
    # browser.find_element_by_id("back_train_date").send_keys("2017-08-11")
    # browser.find_element_by_id("query_ticket").click()

    # time.sleep(2)
    # fromStation = browser.find_element_by_id("fromStation")
    # print browser.execute_script('return arguments[0].value', fromStation)
    # toStation = browser.find_element_by_id("toStation")
    # print browser.execute_script('return arguments[0].value', toStation)
def scanTicket():
    url = "https://kyfw.12306.cn/otn/leftTicket/query?leftTicketDTO.train_date=2017-08-11&leftTicketDTO.from_station=SZQ&leftTicketDTO.to_station=CSQ&purpose_codes=ADULT"
    response = requests.get(url, verify=False)
    response.raise_for_status()
    print(response.text)
    compressdata = json.loads(response.text)
    print json.dumps(compressdata, indent=4, sort_keys=False, ensure_ascii=False)
if __name__=='__main__':
    login()
    checkout()
    openTicket()
    # threading.Thread(target=alive).start()
    # threading.Thread(target=scanTicket).start()
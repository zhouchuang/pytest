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

handlerList = []
property  = {"loginurl":"https://kyfw.12306.cn/otn/login/init","username":"18607371493","password":"1988204110zc","ticket":"https://kyfw.12306.cn/otn/leftTicket/init","fromStationText":"深圳","fromStation":"","toStationText":"长沙","toStation":"","time":"2017-08-17"}
browser = webdriver.Ie()
browser.maximize_window()


def login():
    browser.get(property["loginurl"])
    browser.find_element_by_id("username").send_keys(property["username"])
    browser.find_element_by_id("password").send_keys(property["password"])
    handlerList.append(browser.current_window_handle)

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
# def openTicket():
#     time.sleep(5)
#     browser.get(property["ticket"])
#
#     browser.find_element_by_id("fromStationText").click()
#     browser.find_element_by_id("fromStationText").send_keys("深圳".decode("utf-8"))
#     browser.find_element_by_id("fromStationText").send_keys(Keys.ENTER)
#
#     browser.find_element_by_id("toStationText").click()
#     browser.find_element_by_id("toStationText").send_keys("长沙".decode("utf-8"))
#     browser.find_element_by_id("toStationText").send_keys(Keys.ENTER)
#
#     browser.find_element_by_id("train_date").click()
#     browser.find_element_by_id("train_date").send_keys("2017-08-20")
#     browser.find_element_by_id("train_date").send_keys(Keys.ENTER)

def updateParam():
    response = requests.get("https://kyfw.12306.cn/otn/resources/js/framework/station_name.js?station_version=1.9024", verify=False)
    response.raise_for_status()
    stations  = response.text
    property["fromStation"] = getCityCode(stations,property["fromStationText"])
    property["toStation"] = getCityCode(stations, property["toStationText"])

def getCityCode(stations,city):
    city = city.decode("utf-8")
    fromcitystart = stations.index("|" + city + "|") + len(city) + 2
    fromcityend = stations.index("|", fromcitystart + 1)
    return stations[fromcitystart:fromcityend]

def scanTicket():
    url = "https://kyfw.12306.cn/otn/leftTicket/query?leftTicketDTO.train_date=${time}&leftTicketDTO.from_station=${fromStation}&leftTicketDTO.to_station=${toStation}&purpose_codes=ADULT"
    url = url.replace("${time}",property["time"]).replace("${fromStation}",property["fromStation"]).replace("${toStation}",property["toStation"])

    while True:
        response = requests.get(url, verify=False)
        response.raise_for_status()
        compressdata = json.loads(response.text)
        # print json.dumps(compressdata, indent=4, sort_keys=False, ensure_ascii=False)
        for str in  compressdata["data"]["result"]:
            print str
        time.sleep(1)

def openTicket():
    browser.execute_script("window.open('"+property["ticket"]+"')")
    browser.maximize_window()
    switchCurrentHandler()
    while True:
        time.sleep(1)  # 一秒钟检查一次
        try:
            if browser.find_element_by_xpath(".//tbody[@id='queryLeftTable']/tr").size > 0:
                break
        except Exception, e:
            print e
            pass

def autoSeach():
    while True:
        time.sleep(1)  # 一秒钟检查一次
        browser.find_element_by_id("query_ticket").click()


#浏览器控制切换到最新页面
def switchCurrentHandler():
    global browser
    for handler in browser.window_handles:
        if handler not in handlerList:
            browser.switch_to_window(handler)
            handlerList.append(handler)
            break

if __name__=='__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    login()
    checkout()
    openTicket()
    #autoSeach()
    #updateParam()
    #openTicket()
    # threading.Thread(target=alive).start()
    #threading.Thread(target=scanTicket).start()

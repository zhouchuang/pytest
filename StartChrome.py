#coding:utf-8
import os
import time
import sys
from selenium import webdriver

chrome_driver = os.path.abspath(r"C:\Python27\chromedriver.exe")
os.environ["webdriver.chrome.driver"] = chrome_driver
driver = webdriver.Chrome(chrome_driver)
driver.get("https://www.kaisafax.com/loan")
#coding:utf-8
import os
import time
import sys
import re
import getpass
import threading
import requests
import json
from selenium import webdriver
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import  random
import winsound
import traceback

def log(msg):
    profile = "C:\\Users\\" + getpass.getuser() + "\\.tools" + "\\12306.err"
    file_object = open(profile, 'w')
    file_object.writelines(msg)
    file_object.close()

try:
    browser = webdriver.Ie()
    browser.maximize_window()
except Exception, e:
    log(traceback.format_exc())
    pass


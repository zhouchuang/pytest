#coding:utf-8
import tkinter as tk
from hashlib import sha1

class Config():
    def __init__(self,dict):
        self.username = dict["username"]
        self.password = dict["password"]
        self.receiver=dict["receiver"]

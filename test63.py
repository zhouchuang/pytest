"""
题目：画椭圆ellipse。　
程序分析：无。
"""

# !/usr/bin/python
# -*- coding: UTF-8 -*-
# from tkinter import *
# from tkinter.ttk import Combobox
#
# xin  = Tk()
# Label(xin,text = "文件路径：").grid(row=0,sticky=W)
# Entry(xin).grid(row=0,column=1,sticky=E)
# Button(xin,text="导入考勤文件").grid(row=0,column=2,sticky=E)
# Label(xin,text = "员工姓名：").grid(row=1,sticky=W)
# Combobox(xin).grid(row=1,column=1,sticky=E)
# Button(xin,text="　导出　").grid(row=2,column=1,sticky=E)
# xin.mainloop()




# import tkinter as tk
# from tkinter import ttk
#
# win = tk.Tk()
# win.title("Python GUI")    # 添加标题
#
# ttk.Label(win, text="Chooes a number").grid(column=1, row=0)    # 添加一个标签，并将其列设置为1，行设置为0
# ttk.Label(win, text="Enter a name:").grid(column=0, row=0)      # 设置其在界面中出现的位置  column代表列   row 代表行
#
# # button被点击之后会被执行
# def clickMe():   # 当acction被点击时,该函数则生效
#   action.configure(text='Hello ' + name.get())     # 设置button显示的内容
#   action.configure(state='disabled')      # 将按钮设置为灰色状态，不可使用状态
#
# # 按钮
# action = ttk.Button(win, text="Click Me!", command=clickMe)     # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
# action.grid(column=2, row=1)    # 设置其在界面中出现的位置  column代表列   row 代表行
#
# # 文本框
# name = tk.StringVar()     # StringVar是Tk库内部定义的字符串变量类型，在这里用于管理部件上面的字符；不过一般用在按钮button上。改变StringVar，按钮上的文字也随之改变。
# nameEntered = ttk.Entry(win, width=12, textvariable=name)   # 创建一个文本框，定义长度为12个字符长度，并且将文本框中的内容绑定到上一句定义的name变量上，方便clickMe调用
# nameEntered.grid(column=0, row=1)       # 设置其在界面中出现的位置  column代表列   row 代表行
# nameEntered.focus()     # 当程序运行时,光标默认会出现在该文本框中
#
# # 创建一个下拉列表
# number = tk.StringVar()
# numberChosen = ttk.Combobox(win, width=12, textvariable=number)
# numberChosen['values'] = (1, 2, 4, 42, 100)     # 设置下拉列表的值
# numberChosen.grid(column=1, row=1)      # 设置其在界面中出现的位置  column代表列   row 代表行
# numberChosen.current(0)    # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
#
# win.mainloop()      # 当调用mainloop()时,窗口才会显示出来




import tkinter as tk
from tkinter import ttk

import win32ui


win = tk.Tk()
win.title("员工考勤工具")    # 添加标题
pathtext = tk.StringVar()
numberChosen = ttk.Combobox(win, width=48, textvariable=tk.StringVar())

def exportFile():
    print("export")
def importFile():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('C:/Users/Public/Desktop')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    pathtext.set(filename)
    numberChosen['values'] = ("张三","李四")




ttk.Label(win, text="文件路径").grid(column=0, row=0)
ttk.Entry(win,width=50, textvariable=pathtext,state='readonly').grid(column=1,row=0)
ttk.Button(win,text="导入考勤文件",command=importFile).grid(column=2,row=0)

ttk.Label(win, text="员工姓名").grid(column=0, row=1)

numberChosen['values'] = ('1')   # 设置下拉列表的值
numberChosen.grid(column=1, row=1)      # 设置其在界面中出现的位置  column代表列   row 代表行
numberChosen.current(0)    # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值

action = ttk.Button(win, text="导出", command=exportFile)     # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
action.grid(column=2, row=1)    # 设置其在界面中出现的位置  column代表列   row 代表行

win.mainloop()      # 当调用mainloop()时,窗口才会显示出来


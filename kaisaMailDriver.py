#coding:utf-8

import win32com.client as win32


class OutLook():
    def __init__(self):
        self.app = 'Outlook'

    def send(self,receiver,title,context):

        olook = win32.gencache.EnsureDispatch("%s.Application" % self.app)
        mail=olook.CreateItem(win32.constants.olMailItem)
        mail.Recipients.Add(receiver)
        subj = mail.Subject = title.decode("utf-8")
        body = context.decode("utf-8")
        mail.Body = body
        mail.Send()

class foxMail():
    def __init__(self):
        self.app = 'foxMail'

    def send(self,receiver,title,context):
        olook = win32.gencache.EnsureDispatch("%s.Application" % self.app)
        mail=olook.CreateItem(win32.constants.olMailItem)
        mail.Recipients.Add(receiver)
        subj = mail.Subject = title.decode("utf-8")
        body = context.decode("utf-8")
        mail.Body = body
        mail.Send()




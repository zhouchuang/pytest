# _*_coding:utf-8 _*_
import requests
from bs4 import BeautifulSoup
import re
from PIL import Image
import os

def loginin():
    global session
    session = requests.Session()
    url='https://www.douban.com/accounts/login'
    name='18607371493'
    psw='1988204110zc'
    headers={
    "Host":"www.douban.com",
    "User-Agent":"'Mozilla/5.0 (Windows NT 6.1; rv:53.0)Gecko/20100101 Firefox/53.0'",
    "Accept-Language":"zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
    "Accept-Encoding":"gzip,deflate",
    "Connection":"keep-alive"
    }
    data={'form_email':name,'form_password':psw,'source':'index_nav','remember':'on'}
    captcha=session.get(url,headers=headers,timeout=30)
    soup=BeautifulSoup(captcha.content,'lxml')
    img=soup.find_all('img',id='captcha_image')
    print img
    if img:
        captcha_url=re.findall('src="(.*?)"',str(img))[0]
        print u'验证码所在标签为：',captcha_url
        a=captcha_url.split('&')[0]
        capid=a.split('=')[1]
        print capid
        cap=session.get(captcha_url,headers=headers).content
        with open('captcha.jpg','wb') as f:
            f.write(cap)
            f.close()
            im = Image.open('captcha.jpg')
            im.show()
            capimg=raw_input('请输入验证码：')
            newdata={
            'captcha-solution':capimg,
            'captcha-id':capid}
            data.update(newdata)
            print data
            os.remove('captcha.jpg')
    else:
        print '不存在验证码，请直接登陆'
        r=session.post(url,data=data,headers=headers,timeout=30)
## print(r.content)
## print r.status_code
## #print r.cookies
## html=session.get('https://movie.douban.com/')
## print html.status_code
## print html.content
## print html.cookies
##
if __name__=='__main__':
    loginin()
# coding=utf-8
# 封装selenium常用操作

from seleniumUtil import seleniumUtil


def autoAuthenticate():
    se = seleniumUtil()
    se.openUrl("http://10.252.229.236:8080/portal/customPage/fsadmin/gglogin.jsp")
    se.input('#id_userName','yeyifeng')
    se.input('#id_userPwd','Iephong0@')
    se.clickElement('#login')

if __name__ == '__main__':
    autoAuthenticate()
    

#!/usr/bin/env python
# -*- coding: utf-8 -*-
from src.GlobalList import *
from src.User import *
from src.Select import *
from src.Update import *
from src.Create import *
from src.SQLAnalysis import *
from src.Help import *
from src.Delete import *
from src.Insert import *
from src.MutiSelect import *
from src.View import *

#currentdb = 'test'
#username = 'admin'
global currentdb
global username

def Welcome():
    print "Welcome to V0W_DBMS!!!"
    print "Type '-h' for help, and 'exit' to exit V0W_DBMS."

def passwd_input(msg=''):
    import sys, msvcrt
    if msg != '':
        sys.stdout.write(msg)
    chars = []
    while True:
        newChar = msvcrt.getch()
        if newChar in '\3\r\n':  # 如果是换行，Ctrl+C，则输入结束
            print ''
            if newChar in '\3':  # 如果是Ctrl+C，则将输入清空，返回空字符串
                chars = []
            break
        elif newChar == '\b':  # 如果是退格，则删除末尾一位
            if chars:
                del chars[-1]
                sys.stdout.write('\b \b')  # 左移一位，用空格抹掉星号，再退格
        else:
            chars.append(newChar)
            sys.stdout.write('*')  # 显示为星号
    return ''.join(chars)

def Checklogin():
    usrname = raw_input(u"请输入用户名：")
    print u"请输入密码：",
    passwd = passwd_input()     # 读取密码输入
    Login(usrname, passwd)
    return usrname

def handle(sql):
    sql = sql.lower()
    if 'select' in sql:
        executeSelect(sql)
    elif 'create' in sql:
        executeCreate(sql)
    elif 'update' in sql:
        excuteUpdate(sql)
    elif 'delete' in sql:
        excuteDelete(sql)
    elif 'help' in sql:
        executeHelp(sql)
    elif 'use' in sql:
        executeUse(sql.split(' ')[1])
    elif 'insert' in sql:
        executeInsert(sql)
    else:
        print u'sql 语法错误'

def main():
    username = Checklogin()
    Welcome()
    while True:
        if currentdb is not None:
            sql = raw_input(currentdb+'>')
        else:
            sql = raw_input('>')
        if sql.lower() == 'exit':
            print "Bye!"
            exit(0)
        else:
            handle(sql)

if __name__ == '__main__':
    main()

## 测试
## 1731 Lines
# Checklogin()
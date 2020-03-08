#!/usr/bin/env python
# -*- coding: utf-8 -*-

from User import *
import re
import GlobalList
from Select import *
from Delete import *
from Create import *
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.writer.excel import ExcelWriter
from MutiSelect import *

currentdb = 'test'
username = 'admin'
# 去除列表里面不匹配的括号
def kuohaoCheck(Item):
    li = []
    for istr in Item:
        MatchObj = re.match('(\(.*)', istr)
        if MatchObj:
            istr = istr.replace('(','')

        MatchObj = re.match('(.*)\)\)', istr)
        if MatchObj:
            istr = istr.replace('))',')')

        MatchObj = re.match('\w*\)', istr)
        if MatchObj:
            istr = istr.replace(')','')
        li.append(istr)
    return li

def executeUse(dbname):
    try:
        currentdb = dbname
        wb = load_workbook('data/'+dbname+'.xlsx')
        return wb
    except:
        return None

def executeSelect(sql):
    qx = Checkprivilege(username, currentdb)
    if qx[0] is True or ('select' in qx[2] and (currentdb in qx[1] or qx[1]=='all')):
        if '.' in sql or 'union' in sql or ('(' in sql and ')' in sql):
            MutiSelect(sql)
        else:
            Select(sql, currentdb)
    else:
        print 'Permission Denied.'
        return 1

def executeCreate(sqlstr):
    sqlstr = sqlstr.lower()
    qx = Checkprivilege(username, currentdb)
    if qx[0] is True or ('create' in qx[2] and (currentdb in qx[1] or qx[1] == 'all')):
        if len(sqlstr.split(' ')) < 3:
            print u'SQL 语法错误'
            return
        else:
            if 'database' in sqlstr:
                sqlItem = sqlstr.split(' ')
                dbname = sqlItem[sqlItem.index('database') + 1]
                CreateDatabase(dbname)
            elif 'table' in sqlstr:
                sqlItem = sqlstr.split(' ')
                tablename = sqlItem[sqlItem.index('table') + 1]
                tableItem = sqlItem[sqlItem.index('table') + 2:]
                tableItem = kuohaoCheck(tableItem)
                tablestr = ' '
                tablestr = tablestr.join(tableItem)
                tablenature = tablestr.split(', ')
                CreateTable(currentdb, tablename, tablenature)
            elif 'index' in sqlstr:
                sqlItem = sqlstr.split(' ')
                indexname = sqlItem[sqlItem.index('index') + 1]
                print indexname
            elif 'view' in sqlstr:
                sqlItem = sqlstr.split(' ')
                viewname = sqlItem[sqlItem.index('view') + 1]
                print viewname
    else:
        print 'Permission Denied.'
        return 1

def executeUpdate(sql):
    qx = Checkprivilege(username, currentdb)
    if qx[0] is True or ('update' in qx[2] and (currentdb in qx[1]) or qx[1] == 'all'):
        Update(sql)
    else:
        print 'Permission Denied.'

def executeInsert(sql):
    qx = Checkprivilege(username, currentdb)
    if qx[0] is True or ('insert' in qx[2] and (currentdb in qx[1]) or qx[1] == 'all'):
        Insert(sql)
    else:
        print 'Permission Denied.'

def executeDelete(sql):
    qx = Checkprivilege(username, currentdb)
    if qx[0] is True or ('delete' in qx[2] and (currentdb in qx[1]) or qx[1] == 'all'):
        Insert(sql)
    else:
        print 'Permission Denied.'



# CreateView('Id-Grade', 'test', 'Select id course grade From aaa')

# 测试
sql = 'SELECT * From aaa WHERE x=2 AND y=2'
sql2 = 'CREATE TABLE stu (id char(9) primary key, age int UNIQUE , x char(2) not null, y char(3) check(y>\'a\'))'
#sql1 = 'CREATE DATABASE xxs'
executeCreate(sql)
#executeUse('test')
#currentdb = 'test'
executeCreate(sql2)

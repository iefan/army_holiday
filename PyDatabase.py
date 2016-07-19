#!/usr/bin/python
# -*- coding: utf-8 -*-
# Filename: PyDatabase.py

import sqlite3
def DBSelect(strSelect, DTable, lstItem, flag):
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    if flag == 0:       # fuzzy select
        strSQL = "select * from " + DTable + \
            " where " + lstItem[0] + " like '%%" + strSelect + "%%'"
#        print 'flag 0', strSQL.encode('gbk')

    elif flag == 1:     # Generate Problem, Answer for word
        strSQL = "select "
        for item in lstItem:
            strSQL += item + ","        
        strSQL = strSQL[:-1]
        strSQL += " from " + DTable + " where " + strSelect
        
    elif flag == 2:
        strSQL = "select * from " + DTable + \
            " where " + lstItem[0] + " = '" + strSelect + "'"
                    
#    print strSQL.encode('gbk')
    cur.execute(strSQL)
    strResult = cur.fetchall()
    cur.close()
    conn.close()
    return strResult

def DBInsert(listInput, DTable):
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    strTemp = "("
    for i in range(len(listInput)):
        strTemp += "?,"
    strTemp = strTemp[:-1] + ')'
    strSQL = "insert into " + DTable + " values" + strTemp   

#    print strSQL.encode('gbk')
    cur.execute(strSQL, tuple(listInput))
    conn.commit()
    cur.close()
    conn.close()
    return True
    
def DBUpdate(oldlst, newlst, DTable):
    conn = sqlite3.connect('HolidayData')
    strSQL = "select * from " + DTable
    cur = conn.cursor()
    cur.execute(strSQL)
    columnName = [item[0] for item in cur.description]

    strSQL = "update " + DTable + " set "
    strtmp1 = ""
    strtmp2 = ""
    for item in zip(columnName[1:], newlst, oldlst):
        strtmp1 += item[0] + "='" + item[1] + "',"
        strtmp2 += item[0] + "='" + item[2] + "' and "
    strSQL += strtmp1[:-1] + " where " + strtmp2[:-4]

#    print strSQL.encode('gbk')    
    cur.execute(strSQL)
    conn.commit()
    cur.close()
    conn.close()
    
def DBDelete(lstInput, DTable):
    conn = sqlite3.connect('HolidayData')
    strSQL = "select * from " + DTable
    cur = conn.cursor()
    cur.execute(strSQL)
    columnName = [item[0] for item in cur.description]

    strSQL = "delete from " + DTable + " where "
    strtmp1 = ""
    for item in zip(columnName[1:], lstInput):
        strtmp1 += item[0] + "='" + item[1] + "' and "
    strSQL += strtmp1[:-4]

#    print strSQL.encode('gbk')    
    cur.execute(strSQL)
    conn.commit()
    cur.close()
    conn.close()    

def DBDeleteTab(tabname):
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    strSQL = "delete from " + tabname + " where HoliSn like '%%%%'"
    cur.execute(strSQL)
    conn.commit()
    conn.close()
    
def DBDeleteALL():
    conn = sqlite3.connect('HolidayData')
    tabname = ['PersonInfo', 'RankDays', 'RoadDays', 'HoliDays', 'UnitTab', 'PPHdRecord', 'UserLib']
    cur = conn.cursor()
    for item in tabname:
        strSQL = 'drop table '
        cur.execute(strSQL + item)
        conn.commit()
    conn.close()

def DBCTLib():
#1. PersonInfo: ID, Sn, Name, Sex, RankSn, UnitSn, AddrSn, AddrAll, Married, Tel;
#2. RankDays: ID, RankSn, LevelRank, MinDays, MaxDays;
#3. RoadDays: ID, AddrSn, Address, Days;
#4. HoliDays: ID, HoliSn, HolidayTime, HolidayName, Days;
#5. HolidayKind: ID, KindSn, HdKind; ->XXX
#6. UnitTab: ID, UnitSn, UnitName;
#7. PPHdRecord: ID, Sn, HdKind, Days, DateStart, DateEnd, ByMan, Demo;
#8. UserLib: (ID text, UserName text, Pwd text, Level text)
    conn = sqlite3.connect('HolidayData')
    DTable_PersonInfo = ''' create table PersonInfo 
            (ID Integer Primary key, Sn text, Name text, 
            Sex text, RankSn text, UnitSn text, \
            AddrSn text, AddrAll text, Married text, Tel text, RankTime text)'''
    DTable_BasiclDays = ''' create table RankDays (ID Integer Primary key, RankSn text, LevelRank text, Days text)'''
    DTable_RoadDays = ''' create table RoadDays (ID Integer Primary key, AddrSn text, Address text, MinDays text, MaxDays text)'''
    DTable_HoliDays = ''' create table HoliDays (ID Integer Primary key, HoliSn text, HolidayTime text, HolidayName text, Days text)'''
#    DTable_HolidayKind = ''' create table HolidayKind (ID Integer Primary key, KindSn text, HdKind text)'''
    DTable_UnitTab = ''' create table UnitTab (ID Integer Primary key, UnitSn text, UnitName text)'''
    DTable_PPHdRecord = ''' create table PPHdRecord 
            (ID Integer Primary key, Sn text, Reason text, Days text, 
            DateStart text, DateEnd text, DateInfactEnd text, ByMan text, ByLeader text, \
            DaysDelta text, Demo text, HoliKind text, Demo2 text)'''
    DTable_User = ''' create table UserLib (ID Integer Primary key, UserName text, Pwd text, Level text)'''
    
    tabList = [DTable_PersonInfo, DTable_BasiclDays, DTable_RoadDays, DTable_HoliDays, \
        DTable_UnitTab, DTable_PPHdRecord, DTable_User]
    
    cur = conn.cursor()
    for itab in tabList:
        cur.execute(itab)
        conn.commit()
    
    strSQL = "insert into UserLib values(?, ?, ?, ?)"
    item = [None, 'cc', 'cc', u"管理员"]
            
    cur.execute(strSQL, tuple(item))
    conn.commit()

#    cur.execute('drop table UserLib')
#    conn.commit()    
#    cur.execute("delete from UserLib where UserName == ''")
#    conn.commit()

#    cur.execute('select * from infoOutPerson') 
#    for item in cur.fetchall():
#        print item
    
#    print cur.fetchall()    
    
    cur.close()
    conn.close()

def DBInit():
    conn = sqlite3.connect('HolidayData')
    DBTableLst = ['PersonInfo', 'RoadDays', 'RankDays', 'UnitTab', 'PPHdRecord']
    keyLst = ['Sn', 'AddrSn', 'RankSn', 'UnitSn', 'Sn']
    
    cur = conn.cursor()
    for item in zip(DBTableLst, keyLst):
        strSQL = "delete from %s where %s like '%%%%'" % (item[0], item[1])
        cur.execute(strSQL)
        conn.commit()
    
    cur.close()
    conn.close()
    
def UpdateDBPHR():
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    DTable = "PPHdRecord"    
    strSQL = "alter table " + DTable + " add column Demo2 text"
    print strSQL
#    cur.execute(strSQL)
#    conn.commit()
    strSQL = "select * from " + DTable
    cur.execute(strSQL)
    columnName = [item[0] for item in cur.description]    
    
#    strSQL = "update " + DTable + " set Demo2 = '' where Sn like '%%%%'"
#    cur.execute(strSQL)
    
    print columnName
    conn.close()
    
def UpdateDBPP():
    import os
    if os.path.exists('HolidayData2'):
        os.system('del HolidayData2')
        
    DTable = "PersonInfo"
    DTable_PersonInfo = ''' create table PersonInfo 
            (ID Integer Primary key, Sn text, Name text, 
            Sex text, RankSn text, UnitSn text, \
            AddrSn text, AddrAll text, Married text, Tel text, RankTime text)'''
    DTable_BasiclDays = ''' create table RankDays (ID Integer Primary key, RankSn text, LevelRank text, Days text)'''
    DTable_RoadDays = ''' create table RoadDays (ID Integer Primary key, AddrSn text, Address text, MinDays text, MaxDays text)'''
    DTable_HoliDays = ''' create table HoliDays (ID Integer Primary key, HoliSn text, HolidayTime text, HolidayName text, Days text)'''
#    DTable_HolidayKind = ''' create table HolidayKind (ID Integer Primary key, KindSn text, HdKind text)'''
    DTable_UnitTab = ''' create table UnitTab (ID Integer Primary key, UnitSn text, UnitName text)'''
    DTable_PPHdRecord = ''' create table PPHdRecord 
            (ID Integer Primary key, Sn text, Reason text, Days text, 
            DateStart text, DateEnd text, DateInfactEnd text, ByMan text, ByLeader text, \
            DaysDelta text, Demo text, HoliKind text, Demo2 text)'''
    DTable_User = ''' create table UserLib (ID Integer Primary key, UserName text, Pwd text, Level text)'''
    
    tabList = [DTable_PersonInfo, DTable_BasiclDays, DTable_RoadDays, DTable_HoliDays, \
        DTable_UnitTab, DTable_PPHdRecord, DTable_User]
    
    #=====================================================
    # 1. Create Database and table include 'RankTime' in table 'PersonInfo'
    conn2 = sqlite3.connect('HolidayData2')
    cur2 = conn2.cursor()
    for itab in tabList:
        cur2.execute(itab)
        conn2.commit()
        
    #======================================================
    # 2. Update the table 'PersonInfo' 
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    strSQL = "select * from PersonInfo"
    cur.execute(strSQL)
    ppresult = cur.fetchall()

    strTemp = "("
    for i in range(len(ppresult[0])):
        strTemp += "?,"
    strTemp += "?)"
    
    strSQL = "insert into PersonInfo values" + strTemp
    for item in ppresult:
        tmpitem = list(item)
        tmpitem.append('')
        cur2.execute(strSQL, tuple(tmpitem))
        conn2.commit()
        
    #================================================
    # 3. update the other table
    tabList2 = ['RankDays', 'RoadDays', 'HoliDays', \
        'UnitTab', 'PPHdRecord', 'UserLib']
    for itable in tabList2:
        strSQL1 = "select * from " + itable
        cur.execute(strSQL1)
        strRes1 = cur.fetchall()
        strTemp = "("
        for i in range(len(strRes1[0])):
            strTemp += "?,"
        strTemp = strTemp[:-1] + ')'
        
        strSQL2 = "insert into " + itable + " values" + strTemp        
        for item in strRes1:
            cur2.execute(strSQL2, tuple(item))
            conn2.commit()
    
    cur.close()
    conn.close()
    cur2.close()
    conn2.close()
    

def UpdateDBPP2():
    import xlrd
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    
    strSQL = "select * from PersonInfo"
    cur.execute(strSQL)
    ppresult = cur.fetchall()    
    
    namelst = []
    for item in ppresult:
        namelst.append(item[2])
        
    bk = xlrd.open_workbook("hmc.xls")
    sh = bk.sheets()[0]
    
    xlsTrueHead = [sh.row(0)[i].value for i in range(sh.ncols)]
    for i in range(1, sh.nrows):
        tmpname = sh.row(i)[xlsTrueHead.index(u"姓名")].value
        tmpranktime = sh.row(i)[xlsTrueHead.index(u"军衔时间")].value
        tmpaddrAll = sh.row(i)[xlsTrueHead.index(u"通讯地址")].value
        tmptel = sh.row(i)[xlsTrueHead.index(u"联系电话")].value
        try:
            index = namelst.index(tmpname)
            oldlst = ppresult[index][1:]
            newlst = []
            newlst.extend(oldlst)
            newlst[6] = tmpaddrAll
            newlst[8] = tmptel
            newlst[-1] = tmpranktime
            DBUpdate(oldlst, newlst, 'PersonInfo')
        except:
            print tmpname.encode('gbk')
            print tmpranktime
        
    cur.close()
    conn.close()
    
def ExSelect():
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    
    strSQL = "select * from PersonInfo"
    cur.execute(strSQL)
    for item in cur.fetchall():
        count = 0
        for ich in item[1:]:
            count += 1
            print count,
            print ich.encode('gbk'),
        print
    
    cur.close()
    conn.close()
    
def Select1Tab(DTable):
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    strSQL = "select * from " + DTable
    try:
#        strSQL = 'drop table ' + DTable
        cur.execute(strSQL)
    except:
        cur.close()
        conn.close()
        return False    
    cur.close()
    conn.close()
    return True

def Insert1Tab(DTable):
    conn = sqlite3.connect('HolidayData')
    cur = conn.cursor()
    strSQL = "create table " + DTable + "(ID Integer Primary key)"
    cur.execute(strSQL)
    cur.close()
    conn.close()
    
if __name__ == '__main__':
#    DBCTLib()
#    UpdateDBPP2()
    print Select1Tab('PersonPhd')
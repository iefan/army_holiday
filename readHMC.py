#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import sqlite3
import xlrd
import string

conn = sqlite3.connect('HolidayData')
cur = conn.cursor()

rankDict = [u"列兵", u"上等兵", u"一级士官", u"二级士官", u"三级士官", \
    u"四级士官",u"五级士官", u"六级士官"]

strSQL = "insert into PersonInfo values(?,?,?,?,?,?,?,?,?,?)"
strSQL3 = "insert into RoadDays values(?,?,?,?,?)"
strRankSQL = "insert into RankDays values(?,?,?,?)"
strSQL5 = "insert into UnitTab values(?,?,?)"

for index in range(len(rankDict)):
    ranklst = [None]
    ranklst.append("RSN"+string.zfill(str(index+1),3))
    ranklst.append(rankDict[index])
    ranklst.append(20)
    cur.execute(strRankSQL, tuple(ranklst))
    conn.commit()

bk = xlrd.open_workbook('c:\\hmc.xls')
sh = bk.sheets()[0]

roadtmp = []
unittmp = []
addrCount = 1
count = 1
for i in range(sh.nrows):
    tmplist = [None, 'BZDD'+string.zfill(str(i+1),4)]
    tmplist3 = [None]
    tmplist5 = [None]
    
    if sh.row(i)[3].value not in unittmp:
        unittmp.append(sh.row(i)[3].value)
        tmplist5.append("USN"+string.zfill(str(count),3))
        tmplist5.append(sh.row(i)[3].value)
        count += 1
        cur.execute(strSQL5, tuple(tmplist5))
        conn.commit()   
    
    if sh.row(i)[4].value[:2] not in roadtmp:
        roadtmp.append(sh.row(i)[4].value[:2])
        tmplist3.append("ASN"+string.zfill(str(addrCount),3))
        tmplist3.append(sh.row(i)[4].value[:2])
        tmplist3.append('4')
        tmplist3.append('20')
        addrCount += 1
        cur.execute(strSQL3, tuple(tmplist3))
        conn.commit()
        
    for j in range(6):
        if j == 2:
            index = rankDict.index(sh.row(i)[j].value)
            tmplist.append("RSN"+string.zfill(str(index+1),3))
        elif j == 4:
            index = roadtmp.index(sh.row(i)[j].value[:2])
            tmplist.append("ASN"+string.zfill(str(index+1),3))
#            tmplist.append(sh.row(i)[j].value[:2])
            tmplist.append(sh.row(i)[j].value)
        elif j==3:
            index = unittmp.index(sh.row(i)[j].value)
            tmplist.append("USN"+string.zfill(str(index+1),3))
        else:
            tmplist.append(sh.row(i)[j].value)
    tmplist.append("")
    cur.execute(strSQL, tuple(tmplist))
    conn.commit()
#    print sh.row(i)[j].value.encode('gbk'),
#    print 
#    print strSQL.encode('gbk')

cur.close()

cur.execute('select * from PersonInfo') 
for item in cur.fetchall():
    print item[5].encode('gbk'), item[6].encode('gbk'),item[7].encode('gbk')
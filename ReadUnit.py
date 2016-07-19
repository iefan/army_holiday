#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import xlrd
import string

def readunit(sh):
    xlsHead = [u"姓名", u"单位",u"性别", u"军衔级别",  u"籍贯", u"婚姻状况", u"联系电话", u"通讯地址"]
    nUnitIndex = -1
    for index in range(len(xlsHead)):
        if sh.row(0)[index].value == xlsHead[1]:
            nUnitIndex = index
            break
    if nUnitIndex == -1:
        return

    unitlst = []
    unitlst2 = []

    for i in range(1, sh.nrows):
        unittmp = sh.row(i)[nUnitIndex].value[:1]
        if unittmp not in unitlst2:
            unitlst2.append(unittmp)
    unitlst.extend(unitlst2)
    len0 = len(unitlst)
    newUnitlst = []

    index = 1
    maxlength = -1
    while len(newUnitlst) < len0:
        while True:
            index += 1
            unitlst2 = []
            for i in range(1, sh.nrows):
                if maxlength < len(sh.row(i)[nUnitIndex].value):
                    maxlength = len(sh.row(i)[nUnitIndex].value)
                if sh.row(i)[nUnitIndex].value[:index-1] in unitlst:
                    unittmp = sh.row(i)[nUnitIndex].value[:index]
                    if unittmp not in unitlst2:
                        unitlst2.append(unittmp)

            if len(unitlst) == len(unitlst2):
                unitlst = []
                unitlst.extend(unitlst2)
            else:
                break
            
            if maxlength == index:
                break
        # if the length of row is equal to index,then quit
        if maxlength == index:
            newUnitlst = unitlst
            break
        
        # remove the more unit
        tmplst = []
        tmplst.extend(unitlst2)
        for iunit1 in unitlst:
            count = 0
            for iunit2 in unitlst2:
                if iunit1 == iunit2[:len(iunit1)]:
                    count += 1
            # unit number large 1, need be removed
            if count > 1:
                newUnitlst.append(iunit1)
                for iunit2 in unitlst2:
                    if iunit1 == iunit2[:len(iunit1)]:
                        tmplst.remove(iunit2)
                        
        # update the two tmp unit list                
        unitlst = []
        unitlst.extend(tmplst)
        unitlst2 = []
        unitlst2.extend(tmplst)
    
    return newUnitlst
#    for item in newUnitlst:
#        print item.encode('gbk'),

if __name__ == '__main__':
    readunit(u'c:\\花名册2009.8.5.xls')
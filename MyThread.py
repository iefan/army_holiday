#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import  time
import  thread

import  wx
import  wx.lib.newevent
import  ExportXls

import PyDatabase
import xlrd
import string
import ReadUnit
import operateWord
#----------------------------------------------------------------------

# This creates a new Event class and a EVT binder function
(UpdateBarEvent, EVT_UPDATE_BARGRAPH) = wx.lib.newevent.NewEvent()
(ExportXlsEvent, EVT_EXPORT_XLS) = wx.lib.newevent.NewEvent()
(ImportXlsEvent, EVT_IMPORT_XLS) = wx.lib.newevent.NewEvent()
(ImportXlsEvent2, EVT_IMPORT_XLS2) = wx.lib.newevent.NewEvent()
(ExportWordEvent, EVT_EXPORT_WORD) = wx.lib.newevent.NewEvent()

class ImportXlsThread2:
    def __init__(self, win, personSn, xlspath):
        self.win = win
        self.personSn = personSn
        self.xlspath = xlspath

    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        bk = xlrd.open_workbook(self.xlspath)
        sh = bk.sheets()[0]
        
        # 1.rank infomation
        rankExist = PyDatabase.DBSelect("RankSn like '%%%%'", "RankDays", ['RankSn','LevelRank'], 1)
        rankDict = {}
        for item in rankExist:
            rankDict[item[0]] = item[1]
        
        # 2.unit infomation
        unitExist = PyDatabase.DBSelect("UnitSn like '%%%%'", "UnitTab", ['UnitSn', 'UnitName'], 1)
        unitDict = {}
        for item in unitExist:
            unitDict[item[0]] = item[1].split('|->')[-1]
        
        # 3. road infomation
        roadExist = PyDatabase.DBSelect("AddrSn like '%%%%'", "RoadDays", ['AddrSn','Address'], 1)
        roadDict = {}
        for item in roadExist:
            roadDict[item[0]] = item[1][:2]
            
        # 4. personSn infomation
        strResult = PyDatabase.DBSelect("Sn like '%%%%'", "PersonInfo", ['Sn'], 1)
        lstnewSn = []
        lstIntSn = [int(isn[0][-4:]) for isn in strResult]
        for index in range(1, lstIntSn[-1]):
            if index not in lstIntSn:
                lstnewSn.append(index)
        
        if len(lstnewSn) < sh.nrows-1:
            for index in range(lstIntSn[-1]+1, lstIntSn[-1]+1+sh.nrows):
                lstnewSn.append(index)

        #===========================================================
        #===========================================================
        
        xlsTrueHead = [sh.row(0)[i].value for i in range(sh.ncols)]
        for i in range(1, sh.nrows):
            tmplist = [None, self.personSn+string.zfill(str(lstnewSn[i-1]),4)]
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"姓名")].value)  # name
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"性别")].value)  # sex
            tmplist.append(rankDict.keys()[rankDict.values().index(sh.row(i)[xlsTrueHead.index(u"军衔级别")].value)]) # rank
            # unit
            unittmp = sh.row(i)[xlsTrueHead.index(u"单位")].value
            for iUnit in unitDict.values():
                if unittmp[:len(iUnit)] == iUnit:
                    tmplist.append(unitDict.keys()[unitDict.values().index(iUnit)])
            # road 
            roadtmp = sh.row(i)[xlsTrueHead.index(u"籍贯")].value[:2]
            tmplist.append(roadDict.keys()[roadDict.values().index(roadtmp)])
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"通讯地址")].value)
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"婚姻状况")].value)
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"联系电话")].value)
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"军衔时间")].value)
            PyDatabase.DBInsert(tmplist, "PersonInfo")

        time.sleep(0.5)
        evt = ImportXlsEvent2(flag = True)
        wx.PostEvent(self.win, evt)
        self.running = False


class ExportWordThread:
    def __init__(self, win, lstword):
        self.win = win
        self.lstword = lstword
        
    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        time.sleep(0.5)
        operateWord.GenWordList(self.lstword)
        time.sleep(0.5)
        evt = ExportWordEvent(flag = True)
        wx.PostEvent(self.win, evt)
        self.running = False

class ImportXlsThread:
    def __init__(self, win, personSn, xlspath):
        self.win = win
        self.personSn = personSn
        self.xlspath = xlspath

    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        bk = xlrd.open_workbook(self.xlspath)
        sh = bk.sheets()[0]
        # rank infomation
        rankDict = {}
        rankList = [u"列兵", u"上等兵", u"一级士官", u"二级士官", u"三级士官", \
            u"四级士官",u"五级士官", u"六级士官"]
        for index in range(len(rankList)):
            ranklst = [None]
            ranklst.append("RSN"+string.zfill(str(index+1),3))
            ranklst.append(rankList[index])
            ranklst.append(20)
            rankDict["RSN"+string.zfill(str(index+1),3)] = rankList[index]
            PyDatabase.DBInsert(ranklst, "RankDays")
        
        # unit infomation
        strUnitlst = ReadUnit.readunit(sh)
        unitDict = {}
        for i in range(len(strUnitlst)):
            tmpunit = [None]
            tmpunit.append("USN"+string.zfill(str(i+1),3))
            tmpunit.append(strUnitlst[i])
            unitDict["USN"+string.zfill(str(i+1),3)] = strUnitlst[i]
            PyDatabase.DBInsert(tmpunit, "UnitTab")

        #===========================================================
        #===========================================================
        roadtmp = []
        addrCount = 1
        count = 1
        xlsHead = [u"姓名", u"单位",u"性别", u"军衔级别",  u"籍贯", u"婚姻状况", u"联系电话", u"通讯地址", u"军衔时间"]
        xlsTrueHead = [sh.row(0)[i].value for i in range(len(xlsHead))]
        for i in range(1, sh.nrows):
            tmplist = [None, self.personSn+string.zfill(str(i),4)]
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"姓名")].value)  # name
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"性别")].value)  # sex
            tmplist.append(rankDict.keys()[rankDict.values().index(sh.row(i)[xlsTrueHead.index(u"军衔级别")].value)]) # rank
            
            # unit
            unittmp = sh.row(i)[xlsTrueHead.index(u"单位")].value
            for iUnit in unitDict.values():
                if unittmp[:len(iUnit)] == iUnit:
                    tmplist.append(unitDict.keys()[unitDict.values().index(iUnit)])
            # road 
            tmproad = [None]
            if sh.row(i)[xlsTrueHead.index(u"籍贯")].value[:2] not in roadtmp:
                roadtmp.append(sh.row(i)[xlsTrueHead.index(u"籍贯")].value[:2])
                tmproad.append("ASN"+string.zfill(str(addrCount),3))
                tmproad.append(sh.row(i)[xlsTrueHead.index(u"籍贯")].value[:2])
                tmproad.append('4')
                tmproad.append('20')
                addrCount += 1
                PyDatabase.DBInsert(tmproad, "RoadDays")
                
            # road
            roadindex = roadtmp.index(sh.row(i)[xlsTrueHead.index(u"籍贯")].value[:2]) + 1
            tmplist.append("ASN"+string.zfill(str(roadindex),3))
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"通讯地址")].value)
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"婚姻状况")].value)
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"联系电话")].value)
            tmplist.append(sh.row(i)[xlsTrueHead.index(u"军衔时间")].value)
            PyDatabase.DBInsert(tmplist, "PersonInfo")
            
#        strResult = PyDatabase.DBSelect("", "PersonInfo", ["Sn"], 0)
        time.sleep(0.5)
        evt = ImportXlsEvent(flag = True)
        wx.PostEvent(self.win, evt)
        self.running = False
                
class ExportXlsThread:
    def __init__(self, win, lststr, head, xlspath):
        self.win = win
        self.lststr = lststr
        self.head = head
        self.xlspath = xlspath

    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        time.sleep(0.5)
        ExportXls.ExportXls(self.lststr, self.head, self.xlspath)
        time.sleep(0.5)
        evt = ExportXlsEvent(flag = True)
        wx.PostEvent(self.win, evt)
        self.running = False

class CalcGaugeThread:
    def __init__(self, win, count):
        self.win = win
        self.count = count

    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        while self.keepGoing:
            evt = UpdateBarEvent(count = self.count)
            wx.PostEvent(self.win, evt)
                        
            if self.count >= 50:
                self.count = 0
                
            time.sleep(0.1)
            self.count = self.count + 1

        self.running = False
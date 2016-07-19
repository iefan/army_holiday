#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
#import wx.lib.buttons as buttons
import PyDatabase
import images
import string
import MyValidator
from PhrResource import strVersion, g_UnitSnNum, MyListCtrl
import MyThread
import time
import datetime
import types
import os
import operateWord

#import matplotlib.pyplot as plt
import numpy as np

class FrameMaintain(wx.Frame):               
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        self.panel = panel = wx.Panel(self)
        self.Maximize() 
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        StatPic(panel)
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()
        
class StatPic(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)
        
        titleText = wx.StaticText(panel, -1, u"统计数据图表")
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer0 = wx.BoxSizer(wx.HORIZONTAL)
        topsizer0.Add(titleText, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL)
        
        self.lblunit = wx.StaticText(panel, -1, u"休假统计图表")
        
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(self.lblunit, 1, wx.ALL)

#        self.picwin = wx.Window(panel, -1, style= wx.SIMPLE_BORDER)
#        self.picwin.SetBackgroundColour( wx.WHITE )
#        bmp = wx.Bitmap('bitmap/mainCenter.png')
#        self.pie1 = wx.StaticBitmap(self.picwin, -1, bmp)
#        self.pie2 = wx.StaticBitmap(self.picwin, -1, bmp)
#        self.bar1 = wx.StaticBitmap(self.picwin, -1, bmp)
#        self.lblPie1 = wx.StaticText(self.picwin, -1, u"正在休假人员比例(所有人员)", style= wx.ALIGN_CENTER_HORIZONTAL)
#        self.lblPie2 = wx.StaticText(self.picwin, -1, u"已经休假人员比例(只含士官)", style= wx.ALIGN_CENTER_HORIZONTAL)
#        self.lblBar1 = wx.StaticText(self.picwin, -1, u"本年度各月休假人员统计", style= wx.ALIGN_CENTER_HORIZONTAL)
        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        lstHead = [u"单位", u"总人数", u"士官人数", u"已休假人数", u"未休假人数", u"正在休假"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
#        self.list.Hide()
        lblList = wx.StaticText(panel, -1, u"各单位休假详情列表", style= wx.ALIGN_CENTER_HORIZONTAL)
#        lbllst = [self.lblPie1, self.lblPie2, self.lblBar1, self.lblList, self.pie1, self.pie2, self.bar1]
#        [item.Hide() for item in lbllst]
        
        self.listLegend = MyListCtrl(panel, -1, size=(125,80), style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.listLegend.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        lstHead = [u"颜色说明"]
        [self.listLegend.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
        self.listLegend.SetColumnWidth(0, 120)
        self.listLegend.Enable(False)
#        self.listLegend.Hide()
        
        self.pie1box = wx.StaticBox(panel, -1, 'pie1')
        self.pie2box = wx.StaticBox(panel, -1, 'pie2')
        self.bar1box = wx.StaticBox(panel, -1, 'bar1')
        piesizer = wx.BoxSizer()
        piesizer.Add(self.pie1box, 1, wx.ALL| wx.EXPAND, 10)
        piesizer.Add(self.pie2box, 1, wx.ALL| wx.EXPAND, 10)
        piesizer.Add(self.listLegend, 0, wx.ALL, 10)
        
        picsizer = wx.BoxSizer( wx.VERTICAL )
        picsizer.Add(topsizer, 0, wx.ALL, 5)
        picsizer.Add(piesizer, 3, wx.ALL| wx.EXPAND)
        picsizer.Add(self.bar1box, 3, wx.ALL| wx.EXPAND)
        picsizer.Add(lblList, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        picsizer.Add(self.list, 2, wx.ALL| wx.EXPAND)

        self.btn_Print = wx.Button(panel, -1, u"导出至word") 
        btn_Help = wx.Button(panel, -1, u"帮助")        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [self.btn_Print,  btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1) 
        
        lbltree = wx.StaticText(panel, -1, u"单位树")
        self.tree = wx.TreeCtrl(panel, size=(180, -1))
        self.root = self.tree.AddRoot(u"单位")

        treesizer = wx.BoxSizer(wx.VERTICAL)
        treesizer.AddSizer(lbltree, 0, wx.ALL| wx.EXPAND, 5)
        treesizer.Add(self.tree, 1, wx.ALL|wx.EXPAND, 0)        

        midsizer = wx.BoxSizer(wx.HORIZONTAL)
        midsizer.AddSizer(treesizer, 0, wx.ALL| wx.EXPAND, 5)
        midsizer.AddSizer(picsizer, 1, wx.ALL| wx.EXPAND, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer0, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(midsizer, 1, wx.ALL| wx.EXPAND, 0)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      
        
        self.lstPHR = []
        self.treeDict = {}
        self.DictMonth = {}
        self.DTable = 'PersonInfo'
        
        self.lstCenter = [[0,0], [0,0]]
        self.radius = 0
        self.monthBar = []
        self.lstAngle = []
        
        self.InitData()

        panel.Show()
        panel.Bind(wx.EVT_TREE_ITEM_ACTIVATED, self.OnActivate, self.tree)
        panel.Bind(wx.EVT_BUTTON, self.OnBtnPrint, self.btn_Print)
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
        panel.Bind(wx.EVT_SIZE, self.OnSize)
        
    def InitData(self):
        strPHR = PyDatabase.DBSelect("Sn like '%%%%'", 'PPHdRecord', ['Sn', 'HoliKind', 'DateStart', 'DateInfactEnd'], 1)
        [self.lstPHR.append([iPHR[0][:g_UnitSnNum], iPHR[1], iPHR[2], iPHR[3]]) for iPHR in strPHR]
        self.InitTree()

    def InitTree(self):
        strResult = PyDatabase.DBSelect('', 'UnitTab', ['UnitSn'], 0) 
        lstunitname = []
        for row in strResult:
            lstunitname.append((row[1], row[2].split('|->')))
        
        self.CreateTreeByList(lstunitname)
        self.tree.Expand(self.root)
        self.InitListData()
        
    def InitListData(self):
        thisyear = time.localtime()[0]
        for iusn in self.treeDict:
            tmplst = []
            totalPP = PyDatabase.DBSelect("UnitSn = '" + iusn+ "'", self.DTable, ['Sn'], 1)
            partPP = PyDatabase.DBSelect("UnitSn = '" + iusn+ "' and RankSn <> 'RSN001' and RankSn <> 'RSN002'", self.DTable, ['Sn'], 1)
            lstSelSn = []
            [lstSelSn.append(isn[0]) for isn in totalPP]
            lstSelSn2 = []
            [lstSelSn2.append(isn[0]) for isn in partPP]
            
            OutPP = 0
            monthPP = []
            [monthPP.append(0) for i in range(12)]
            
            lstUniHoli = []
            for iphr in self.lstPHR:
                if iphr[0] in lstSelSn:
                    tmpSt = iphr[2].split('-')
                    if int(tmpSt[0]) == thisyear:
                        if iphr[3] == "": 
                            OutPP += 1
                        monthPP[int(tmpSt[1])-1] += 1
                        if iphr[0] not in lstUniHoli:
                            lstUniHoli.append(iphr[0])
                        
            lstUniHoli2 = []
            lstUniHoli2.extend(lstUniHoli)
            for isn in lstUniHoli:
                if isn not in lstSelSn2:
                    lstUniHoli2.remove(isn)
            
            tmplst.append(len(totalPP))
            tmplst.append(len(partPP))
            tmplst.append(len(lstUniHoli2))
            tmplst.append(len(partPP) - len(lstUniHoli2))
            tmplst.append(OutPP)
            self.DictMonth[iusn] = monthPP
            
            self.list.SetColumnWidth(0, wx.LIST_AUTOSIZE)
            
            index = self.list.InsertStringItem(100, "A")
            self.list.SetStringItem(index, 0, self.GetItemText(self.treeDict[iusn]))
            [self.list.SetStringItem(index, i+1, `tmplst[i]`) for i in range(len(tmplst))]
            self.DispColorList(self.list)
        
        for istr, icolor in zip([u"在位（或未休假）", u"休假（或已休假）"], [(0,128,0), (0,0,255)]):
            index = self.listLegend.InsertStringItem(10, "A")
            self.listLegend.SetStringItem(index, 0, istr)
            self.listLegend.SetItemBackgroundColour(index, icolor)
            self.listLegend.SetItemTextColour(index, (255, 255, 255))
        
    def DispColorList(self, list):
        for i in range(list.GetItemCount()):
            if i%2 == 0: list.SetItemBackgroundColour(i, (245, 245, 245))
            if i%2 == 1: list.SetItemBackgroundColour(i, (211, 211, 211))
            
    def CreateTreeByList(self, lststr):
        '''Create Tree by list'''
        if len(lststr) == 0:
            return
        
        flagModRoot = True
        if len(lststr) >= 2:
            if lststr[0][1][0] != lststr[1][1][0]:
                flagModRoot = False
        
        for item in lststr:
            parentItem = self.root
            if flagModRoot:
                itemlst = item[1][1:]
            else:
                itemlst = item[1]
            for ichild in itemlst:
                sibitem, cookie = self.tree.GetFirstChild(parentItem)
                while sibitem.IsOk():
                    '''parent node is the same'''
                    if self.GetItemText(sibitem) == ichild:
                        break
                    sibitem = self.tree.GetNextSibling(sibitem)
                    
                if self.GetItemText(sibitem) != ichild:
                    parentItem = self.tree.AppendItem(parentItem, ichild)
                else:
                    parentItem = sibitem
            # Save the TreeItemId
            self.treeDict[item[0]] = parentItem
            
        if flagModRoot:
            self.tree.SetItemText(self.root, lststr[0][1][0])
            
    def GetItemText(self, item):
        if item:
            return self.tree.GetItemText(item)
        else:
            return ""

    def OnActivate(self, event):
        item = event.GetItem()
        listIndex = -1
        for index in range(self.list.GetItemCount()):
            if self.GetItemText(item) == self.list.GetItem(index, col=0).GetText():
                listIndex = index
                break
            
        if listIndex != -1:
            outPP = int(self.list.GetItem(listIndex, col=5).GetText())
            inPP = int(self.list.GetItem(listIndex, col=1).GetText()) - outPP
            holiPP = int(self.list.GetItem(listIndex, col=3).GetText())
            noholiPP = int(self.list.GetItem(listIndex, col=2).GetText()) - holiPP
            monthPP = self.DictMonth[self.treeDict.keys()[self.treeDict.values().index(item)]]
        else:
            outPP = 0
            inPP = 0
            holiPP = 0
            noholiPP = 0
            for index in range(self.list.GetItemCount()):
                outPP += int(self.list.GetItem(index, col=5).GetText())
                inPP += int(self.list.GetItem(index, col=1).GetText())
                holiPP = int(self.list.GetItem(index, col=3).GetText())
                noholiPP = int(self.list.GetItem(index, col=2).GetText())
            inPP = inPP - outPP
            noholiPP = noholiPP - holiPP
            
            monthPP = self.DictMonth.values()[0]
            for imon in self.DictMonth.values()[1:]:
                monthPP = list(np.sum(zip(monthPP, imon), axis=1))
        
        self.monthBar = monthPP
        lstUnitPie1 = [outPP, inPP]
        lstUnitPie2 = [holiPP, noholiPP]
        self.lstAngle = []
        for ipie in [lstUnitPie1, lstUnitPie2]:
            self.lstAngle.append(float(ipie[0])/(ipie[0] + ipie[1])*360)
        
        self.lblunit.SetLabel(self.GetItemText(item) + u"　休假统计图表")
        
        pielst = [self.pie1box, self.pie2box]
        for i in range(2):
            self.lstCenter[i][0] = pielst[i].GetPosition()[0] + pielst[i].GetSize()[0]/2
            self.lstCenter[i][1] = pielst[i].GetPosition()[1] + pielst[i].GetSize()[1]/2
            pielst[i].Hide()

        self.radius = min(self.pie1box.GetSize())/2
        self.panel.Refresh()

    def DispBarChart(self, monthPP, bar1, size):
        month1 = range(1,13)
        try:
            fig = plt.figure(2, figsize=size)
            rects1 = plt.bar(month1, monthPP, width=0.3,color='r')
        except:
            wx.MessageBox(u"无法创建柱状图！", u"提示")
            return
        
        width = 0.3        
        plt.axis([1, 12+2*width, 0, max(monthPP)*1.2])
        plt.xticks(np.arange(1+width,13+width),("1", "2","3", "4","5", "6","7", "8","9", "10","11", "12"))
        
        def autolabel(rects):
            # attach some text labels
            for rect in rects:
                height = rect.get_height()
                if height != 0:
                    plt.text(rect.get_x()+rect.get_width()/2., 1.05*height, '%d'%int(height),
                        ha='center', va='bottom')
        autolabel(rects1)
        plt.savefig(bar1)
        plt.clf()
        plt.cla()
 
    def OnBtnPrint(self, event):
        return
        if not self.list.IsShown():
            wx.MessageBox(u"请先点击查看各单位的休假统计！", u"提示")
            return
        self.btn_Print.Enable(False)
        title = self.lblunit.GetLabel()
        picpath = [os.getcwd()+'\\pie1.png', os.getcwd()+'\\pie2.png', os.getcwd()+'\\bar1.png']
        
        lsthead = [self.list.GetColumn(i).GetText() for i in range(self.list.GetColumnCount())]
        lstdata = [lsthead]
        for irow in range(self.list.GetItemCount()):
            lsttmp = []
            for icol in range(self.list.GetColumnCount()):
                lsttmp.append(self.list.GetItem(irow, icol).GetText())
            lstdata.append(lsttmp)
        try:
            operateWord.GenStatWord(title, picpath, lstdata)
        except:
            pass
        self.btn_Print.Enable(True)
        
    def OnHelp(self, event):        
        try:
            os.startfile(r'help\index.htm') 
        except:
            wx.MessageBox(u'当前版本没有找到帮助文档!', u'提示')
    
    def OnSize(self, event):
        lstbox = [self.pie1box, self.pie2box, self.bar1box]
        [item.Show() for item in lstbox]
        event.Skip()
        
    def OnPaint(self, event):
        pdc = wx.PaintDC(self.panel)
        try:
            dc = wx.GCDC(pdc)
        except:
            dc = pdc
#        rect = wx.Rect(300,300, 100, 100)
#        self.center1 = (350,  180)
#        self.center2 = (630, 180)
#        self.radius = 100
        
        center1 = tuple(self.lstCenter[0])
        center2 = tuple(self.lstCenter[1])
        
        for RGB, pos in [((0, 128,  0), center1),
                         ((0, 128, 0), center2)
                         ]:
            r, g, b = RGB
            penclr   = wx.Colour(r, g, b, wx.ALPHA_OPAQUE)
            brushclr = wx.Colour(r, g, b, 128)   # half transparent
            dc.SetPen(wx.Pen(penclr))
            dc.SetBrush(wx.Brush(brushclr))
#            rect.SetPosition(pos)
            dc.DrawCirclePoint(pos, self.radius)
            
        dc.SetPen(wx.Pen(wx.BLACK))
        for iangle, icenter in zip(self.lstAngle, self.lstCenter):
            if iangle != 0:
                dc.DrawLinePoint(icenter, (icenter[0], icenter[1]+self.radius))
                line1pos = (icenter[0] + self.radius*np.sin(iangle/180.*np.pi), icenter[1] + self.radius*np.cos(iangle/180.*np.pi))
                dc.DrawLinePoint(icenter, line1pos)
                dc.SetBrush(wx.Brush(wx.BLUE))
                dc.DrawArcPoint((icenter[0], icenter[1]+self.radius), line1pos, icenter)

        # some additional testing stuff
        #dc.SetPen(wx.Pen(wx.Colour(0,0,255, 196)))
        #dc.SetBrush(wx.Brush(wx.Colour(0,0,255, 64)))
        #dc.DrawCircle(50, 275, 25)
        #dc.DrawEllipse(100, 275, 75, 50)
#        dc = wx.PaintDC(self.panel)
#        dc.BeginDrawing()
#        dc.EndDrawing()

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frmMaintain = FrameMaintain()
    frmMaintain.Show()
    app.MainLoop()
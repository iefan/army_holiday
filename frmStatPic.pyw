#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import PyDatabase
import images
import string
from PhrResource import strVersion, g_UnitSnNum, MyListCtrl
import time
import operateWord
import os
#from Image import new

from math import cos, sin, pi

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
        
        titleText = wx.StaticText(panel, -1, u"统计数据图表", size=(300, -1), style= wx.ALIGN_CENTER)
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))

        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        lstHead = [u"单位", u"总人数", u"士官人数", u"已休假人数", u"未休假人数", u"正在休假"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]

        lblList = wx.StaticText(panel, -1, u"各单位休假详情列表", style= wx.ALIGN_CENTER_HORIZONTAL)
        
        self.listLegend = MyListCtrl(panel, -1, size=(125,80), style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.listLegend.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        lstHead = [u"颜色说明"]
        [self.listLegend.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
        self.listLegend.SetColumnWidth(0, 120)
        self.listLegend.Enable(False)
#        self.listLegend.Hide()
        
        self.pie1box = wx.StaticBox(panel)
        self.pie2box = wx.StaticBox(panel)
        self.bar1box = wx.StaticBox(panel)
        lblPie1 = wx.StaticText(panel, -1, u"正在休假人员比例(所有人员)")
        lblPie2 = wx.StaticText(panel, -1, u"已经休假人员比例(只含士官)")
        lblBar1 = wx.StaticText(panel, -1, `time.localtime()[0]` + u"年度各月休假人次统计")
        
        pie1sizer = wx.BoxSizer( wx.VERTICAL)
        pie1sizer.Add(lblPie1, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 0)
        pie1sizer.Add(self.pie1box, 1, wx.ALL| wx.EXPAND, 10)
        
        pie2sizer = wx.BoxSizer( wx.VERTICAL)
        pie2sizer.Add(lblPie2, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 0)
        pie2sizer.Add(self.pie2box, 1, wx.ALL| wx.EXPAND, 10)

        piesizer = wx.BoxSizer()
        piesizer.Add(pie1sizer, 1, wx.ALL| wx.EXPAND, 10)
        piesizer.Add(pie2sizer, 1, wx.ALL| wx.EXPAND, 10)
        piesizer.Add(self.listLegend, 0, wx.ALL, 10)
        
        picsizer = wx.BoxSizer( wx.VERTICAL )
        picsizer.Add(piesizer, 3, wx.ALL| wx.EXPAND)
        picsizer.Add(lblBar1, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 10)
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
        
        self.lbltree = wx.StaticText(panel, -1, u"单位树")
        self.tree = wx.TreeCtrl(panel, size=(180, -1))
        self.root = self.tree.AddRoot(u"单位")

        treesizer = wx.BoxSizer(wx.VERTICAL)
        treesizer.AddSizer(self.lbltree, 0, wx.ALL| wx.EXPAND, 5)
        treesizer.Add(self.tree, 1, wx.ALL|wx.EXPAND, 0)        

        midsizer = wx.BoxSizer(wx.HORIZONTAL)
        midsizer.AddSizer(treesizer, 0, wx.ALL| wx.EXPAND, 5)
        midsizer.AddSizer(picsizer, 1, wx.ALL| wx.EXPAND, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(titleText, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
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
        self.lstAngle = []
        
        self.barlstRect = []
        self.barBottomLinePos = [(0,0), (0,0)]
        self.barlstData = []
        self.InitData()

        panel.Show()
        panel.Bind(wx.EVT_TREE_ITEM_ACTIVATED, self.OnActivate, self.tree)
        panel.Bind(wx.EVT_BUTTON, self.OnBtnPrint, self.btn_Print)
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
#        panel.Bind(wx.EVT_SIZE, self.OnSize)
        
    def InitData(self):
        strPHR = PyDatabase.DBSelect("Sn like '%%%%'", 'PPHdRecord', ['Sn', 'HoliKind', 'DateStart', 'DateInfactEnd'], 1)
        [self.lstPHR.append([iPHR[0][:g_UnitSnNum], iPHR[1], iPHR[2], iPHR[3]]) for iPHR in strPHR]
        self.InitTree()
        self.btn_Print.Enable(False)

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
            
            index = self.list.InsertStringItem(100, "A")
            self.list.SetStringItem(index, 0, self.GetItemText(self.treeDict[iusn]))
            [self.list.SetStringItem(index, i+1, `tmplst[i]`) for i in range(len(tmplst))]
            self.DispColorList(self.list)
            
        self.list.SetColumnWidth(0, 200)
        
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
            for idata in self.DictMonth.values():
                tmp = []
                for ia, ib in zip(monthPP, idata):
                    tmp.append(ia+ib)
                monthPP = []
                monthPP.extend(tmp)
            
        #====================================================
        lstUnitPie1 = [outPP, inPP]
        lstUnitPie2 = [holiPP, noholiPP]
        self.lstAngle = []
        for ipie in [lstUnitPie1, lstUnitPie2]:
            self.lstAngle.append(float(ipie[0])/(ipie[0] + ipie[1])*360)
        
        self.lbltree.SetLabel(u"当前显示："+self.GetItemText(item))
        self.lbltree.SetBackgroundColour(wx.Colour(0, 240, 200, 128))
        
        pielst = [self.pie1box, self.pie2box]
        for i in range(2):
            self.lstCenter[i][0] = pielst[i].GetPosition()[0] + pielst[i].GetSize()[0]/2
            self.lstCenter[i][1] = pielst[i].GetPosition()[1] + pielst[i].GetSize()[1]/2
            
        self.radius = min(self.pie1box.GetSize())/2
        #==============================================
        self.barlstRect = []
        barW = float(self.bar1box.GetSize()[0]-30)
        barH = float(self.bar1box.GetSize()[1] - 30)
        if max(monthPP) != 0:
            eachH = barH/(max(monthPP)*1.1)
        else:
            eachH = 0
        eachW = barW/len(monthPP)
        startPos = self.bar1box.GetPosition()
        for i in range(len(monthPP)):
            x = startPos[0] + i*eachW + eachW/3. +15
            y = startPos[1] + (barH-monthPP[i]*eachH)
            self.barlstRect.append((int(x), int(y), int(eachW/3.), int(monthPP[i]*eachH)))
        
        self.barBottomLinePos[0] = (startPos[0]+15, startPos[1]+barH)
        self.barBottomLinePos[1] = (startPos[0]+barW, startPos[1]+barH)
        self.barlstData = monthPP
        #===============================================
        [item.Hide() for item in [self.bar1box,self.pie1box, self.pie2box]]
        self.panel.Refresh()
        self.btn_Print.Enable(True)
        
    def OnBtnPrint(self, event):
        self.btn_Print.Enable(False)
        lstBmpPath = ['pie1.bmp','pie2.bmp','bar1.bmp']
        lstImg = [self.pie1box, self.pie2box, self.bar1box]
        
        dcsource = wx.ClientDC(self.panel)
        for ipath, ibox in zip(lstBmpPath, lstImg):
            size = ibox.GetSize()
            pos = ibox.GetPosition()
            if ipath != 'bar1.bmp':
                pos = (pos[0] + (size[0]-min(size))/2.-1, pos[1] + (size[1]-min(size))/2.-1)
                size = (min(size)+2, min(size)+2)
            
            bmp = wx.EmptyBitmap(size[0], size[1])
#            img = new('RGB', size)
#            img.save(ipath)
            
#            bmp = wx.Image(ipath).ConvertToBitmap()
            dc = wx.MemoryDC(bmp)
            
            dc.Blit(0, 0, bmp.GetWidth(),bmp.GetHeight(), dcsource, pos[0], pos[1])
            img = dc.GetAsBitmap()
            img.SaveFile(ipath, wx.BITMAP_TYPE_BMP)

        self.btn_Print.Enable(False)
        title = self.lbltree.GetLabel()[5:] + u"休假数据统计图表"
        picpath = []
        [picpath.append(os.getcwd() + "\\" + ipath) for ipath in lstBmpPath]
#        [os.getcwd()+'\\pie1.bmp', os.getcwd()+'\\pie2.bmp', os.getcwd()+'\\bar1.bmp']
        
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
    
#    def OnSize(self, event):
#        [item.Show() for item in [self.bar1box,self.pie1box, self.pie2box]]
    
    def OnHelp(self, event):        
        try:
            os.startfile(r'help\index.htm') 
        except:
            wx.MessageBox(u'当前版本没有找到帮助文档!', u'提示')
            
    def OnPaint(self, event):
        pdc = wx.PaintDC(self.panel)
        try:
            dc = wx.GCDC(pdc)
        except:
            dc = pdc

        center1 = tuple(self.lstCenter[0])
        center2 = tuple(self.lstCenter[1])
        color = (0, 128, 0)
        for icenter in self.lstCenter:
            r, g, b = color
            penclr   = wx.Colour(r, g, b, wx.ALPHA_OPAQUE)
            brushclr = wx.Colour(r, g, b, 128)   # half transparent
            dc.SetPen(wx.Pen(penclr))
            dc.SetBrush(wx.Brush(brushclr))
            dc.DrawCirclePoint(icenter, self.radius)
        
        dc.SetPen(wx.Pen(wx.BLACK))
        for iangle, icenter in zip(self.lstAngle, self.lstCenter):
            if iangle != 0:
                dc.DrawLinePoint(icenter, (icenter[0], icenter[1]+self.radius))
                line1pos = (icenter[0] + self.radius*sin(iangle/180.*pi), icenter[1] + self.radius*cos(iangle/180.*pi))
                dc.DrawLinePoint(icenter, line1pos)
                dc.SetBrush(wx.Brush(wx.Colour(0, 0, 255, 128)))
                dc.DrawArcPoint((icenter[0], icenter[1]+self.radius), line1pos, icenter)
            
            # Draw the pie label
            pos = (icenter[0]+self.radius*sin(iangle/400.*pi)/2., icenter[1] + self.radius*cos(iangle/400.*pi)/2.)
            dc.DrawTextPoint('%1.1f' % (iangle/360.*100) + "%", pos)
            jangle = 360+iangle
            pos = (icenter[0]+self.radius*sin(jangle/360.*pi)/2., icenter[1] + self.radius*cos(jangle/360.*pi)/2.)
            dc.DrawTextPoint('%1.1f' % ((1-iangle/360.)*100) + "%", pos)
            
        # Draw the bar
        dc.SetBrush(wx.Brush(wx.RED))
        dc.DrawLinePoint(self.barBottomLinePos[0], self.barBottomLinePos[1])
        dc.DrawRectangleList(self.barlstRect)
        
        for i in range(len(self.barlstRect)):
            pos = [self.barlstRect[i][0], self.barlstRect[i][1]+self.barlstRect[i][3]+2]
            dc.DrawTextPoint(`i+1`+u"月", pos)
            if self.barlstData[i] != 0:
                pos = [self.barlstRect[i][0]+2, self.barlstRect[i][1]-14]
                dc.DrawTextPoint(`self.barlstData[i]`, pos)

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frmMaintain = FrameMaintain()
    frmMaintain.Show()
    app.MainLoop()
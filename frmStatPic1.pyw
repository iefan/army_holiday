#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
#import wx.lib.buttons as buttons
import PyDatabase
import images
import string
import MyValidator
from PhrResource import ALPHA_ONLY, DIGIT_ONLY, strVersion, g_UnitSnNum
import MyThread
import time
import datetime
import types
import os

import matplotlib.pyplot as plt
import numpy as np

#PIE1_W = 200
#PIE1_H = 200

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

        self.ckbRank = wx.CheckBox(panel, -1, u"军衔")
        self.ckbRankTime = wx.CheckBox(panel, -1, u"时间")
        self.ckbAddr = wx.CheckBox(panel, -1, u"籍贯")
        self.ckbSex = wx.CheckBox(panel, -1, u"性别")
        self.ckbMarried = wx.CheckBox(panel, -1, u"婚否")
        
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        itemSelect = [self.ckbRank, (-1,-1), self.ckbRankTime, (-1,-1), self.ckbAddr, (-1,-1), self.ckbSex, (-1,-1), self.ckbMarried]
        for item in itemSelect:
            if types.TupleType == type(item):
                topsizer.Add(item, 1, wx.ALL| wx.EXPAND)
            else:                
                topsizer.Add(item, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL)
                
        self.picwin = wx.Window(panel, -1, style= wx.SIMPLE_BORDER)
#        bmp = wx.Image('c:\\tmp.png', wx.BITMAP_TYPE_PNG).ConvertToBitmap()
        self.pie1 = wx.StaticBitmap(self.picwin, -1, wx.NullBitmap)
        self.pie2 = wx.StaticBitmap(self.picwin, -1, wx.NullBitmap)
        self.bar1 = wx.StaticBitmap(self.picwin, -1, wx.NullBitmap)
        self.lblPie1 = wx.StaticText(self.picwin, style= wx.ALIGN_CENTER_HORIZONTAL)
        self.lblPie2 = wx.StaticText(self.picwin)
        self.lblBar1 = wx.StaticText(self.picwin)
        self.list = wx.ListCtrl(self.picwin)
        
        picsizer = wx.BoxSizer( wx.VERTICAL )
        picsizer.Add(topsizer, 0, wx.ALL, 5)
        picsizer.Add(self.picwin, 1, wx.ALL| wx.EXPAND)

        btn_Print = wx.Button(panel, -1, u"打印") 
        btn_Help = wx.Button(panel, -1, u"帮助")        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [btn_Print,  btn_Help]: 
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
        self.treeSel = None
        self.listIndex = 0
        self.unitDict = {}
        self.rankDict = {}
        self.roadDict = {}
        self.DTable = 'PersonInfo'
        
        self.InitData()

        panel.Show()
        itemchecklst = [self.ckbRank, self.ckbRankTime, self.ckbAddr, self.ckbSex, self.ckbMarried]
        [panel.Bind(wx.EVT_CHECKBOX, self.OnCkbInfo, item) for item in itemchecklst]
        panel.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelChanged, self.tree)
        panel.Bind(wx.EVT_TREE_ITEM_ACTIVATED, self.OnActivate, self.tree)
        panel.Bind(wx.EVT_BUTTON, self.OnBtnPrint, btn_Print)
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
        
    def InitData(self):
        strPHR = PyDatabase.DBSelect("Sn like '%%%%'", 'PPHdRecord', ['Sn', 'HoliKind', 'DateStart', 'DateInfactEnd'], 1)
        [self.lstPHR.append([iPHR[0][:g_UnitSnNum], iPHR[1], iPHR[2], iPHR[3]]) for iPHR in strPHR]
        
        lstRank = PyDatabase.DBSelect(u"ID like '%%%%'", "RankDays", ['RankSn', 'LevelRank'], 1)
        list_Rank = []
        for item in lstRank:
            self.rankDict[item[0]] = item[1]
            list_Rank.append(item[1])

        lstRTime = []
        for i in range(2020, 2000, -1):
            lstRTime.append(`i`)
        lstAddr = PyDatabase.DBSelect(u"ID like '%%%%'", "RoadDays", ['AddrSn','Address'], 1)
        list_Addr = []
        for item in lstAddr:
            self.roadDict[item[0]] = item[1]
            list_Addr.append(item[1])

        self.InitTree()

    def OnCkbInfo(self, event):
        self.OnSelect(None)

    def InitTree(self):
        strResult = PyDatabase.DBSelect('', 'UnitTab', ['UnitSn'], 0) 
        lstunitname = []
        for row in strResult:
            self.unitDict[row[1]] = row[2]
            lstunitname.append((row[1], row[2].split('|->')))
        
        self.CreateTreeByList(lstunitname)
        self.tree.Expand(self.root)
    
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
    
    def OnSelChanged(self, event):
        item = event.GetItem()
        if item in self.treeDict.values():
            self.treeSel = item
        else:
            self.treeSel = None
        
    def OnActivate(self, event):
        item = event.GetItem()
        
        if item in self.treeDict.values():
            curUsn = self.treeDict.keys()[self.treeDict.values().index(item)]
            strResult = PyDatabase.DBSelect("UnitSn = '" + curUsn+ "'", self.DTable, ['Sn'], 1)
            self.treeSel = item            
        else:
            strResult = PyDatabase.DBSelect("UnitSn like '%%%%'", self.DTable, ['Sn'], 1)
            self.treeSel = None
        
        lstSelSn = []
        [lstSelSn.append(isn[0]) for isn in strResult]
        OutPP = 0
        monthPP = []
        [monthPP.append(0) for i in range(12)]
        
        thisyear = time.localtime()[0]
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
        
#        HoliPP = sum(monthPP)
        
        lstUnitPie1 = [OutPP, len(lstSelSn) - OutPP]
        lstUnitPie2 = [len(lstUniHoli), len(lstSelSn)-len(lstUniHoli)]
        
#        for i in range(12):
#            monthPP[i] = float(monthPP[i])/len(lstSelSn)
        
        pie_w_h = self.picwin.GetSize()[1]*0.35
        pie_pos = self.picwin.GetSize()[0] - 2*pie_w_h
        
        self.pie1.SetPosition((pie_pos/3, 5))
        self.pie1.SetSize((pie_w_h, pie_w_h))
        self.lblPie1.SetPosition((pie_pos/3+pie_w_h/4, pie_w_h-10))
#        self.lblPie1.SetSize((pie_w_h, -1))
        self.pie2.SetPosition((pie_pos/3*2+pie_w_h, 5))
        self.pie2.SetSize((pie_w_h, pie_w_h))
        self.lblPie2.SetPosition((pie_pos/3*2+pie_w_h/4*5, pie_w_h-10))
#        self.lblPie2.SetSize((pie_w_h, -1))
        pie1 = 'pie1.png'
        pie2 = 'pie2.png'
        bar1 = 'bar1.png'
        try:
            plt.figure(1, figsize=(float(pie_w_h)/100, float(pie_w_h)/100))
            for ipieData, ipieName in zip([lstUnitPie1, lstUnitPie2], [pie1, pie2]):
                plt.pie(ipieData, autopct='%1.1f%%', shadow=True)
                plt.savefig(ipieName)
                plt.clf()
                plt.cla()
        except:
            wx.MessageBox(u"无法创建饼状图！", u"提示")
            return
            
        if os.path.exists(pie1):
            bmp = wx.Image(pie1, wx.BITMAP_TYPE_PNG).Rescale(pie_w_h, pie_w_h).ConvertToBitmap()
            self.pie1.SetBitmap(bmp)
            self.lblPie1.SetLabel(u"正在休假人员比例")
        if os.path.exists(pie2):
            bmp = wx.Image(pie2, wx.BITMAP_TYPE_PNG).Rescale(pie_w_h, pie_w_h).ConvertToBitmap()
            self.pie2.SetBitmap(bmp)
            self.lblPie2.SetLabel(u"已经休假人员比例")
        
        bar_w = self.picwin.GetSize()[0]-10
        bar_h = self.picwin.GetSize()[1]*0.3
        self.DispBarChart(monthPP, bar1, (float(bar_w)/100, float(bar_h)/100))
        self.bar1.SetPosition((5, self.pie1.GetSize()[1]+10))
        self.bar1.SetSize((bar_w, bar_h))
        self.lblBar1.SetPosition((5+bar_w/5*2, self.pie1.GetSize()[1]+20+bar_h))
#        self.lblBar1.SetSize((bar_w, -1))
        bmp = wx.Image(bar1, wx.BITMAP_TYPE_PNG).Rescale(bar_w, bar_h).ConvertToBitmap()
        self.bar1.SetBitmap(bmp)
        self.lblBar1.SetLabel(u"本年度各月休假人员统计")
    
    def DispBarChart(self, monthPP, bar1, size):
        month1 = range(1,13)
        try:
            plt.figure(2, figsize=size)
            plt.bar(month1, monthPP, width=0.3,color='r')
        except:
            wx.MessageBox(u"无法创建柱状图！", u"提示")
            return
        
        width = 0.3
        plt.axis([1, 12+2*width, 0, max(monthPP)*1.1])
        plt.xticks(np.arange(1+width,13+width),("1", "2","3", "4","5", "6","7", "8","9", "10","11", "12"))
        plt.savefig(bar1)
        plt.clf()
        plt.cla()
        
#        month2 = range(1+width, 13+width)
        
#        print len(lstSelSn), InPP, HoliPP, len(lstUniHoli), OutPP
#        print monthPP
    
    def OnOutXls(self):
        dlg = wx.DirDialog(self.panel, u"请选择一个保存目录:", style=wx.DD_DEFAULT_STYLE)
        
        pathXls = ""
        if dlg.ShowModal() == wx.ID_OK:
            pathXls =  dlg.GetPath()
        dlg.Destroy()
        if pathXls == "":
            wx.MessageBox(u"未选择保存目录！", u"提示")
            return
        
        head = [self.list.GetColumn(index).GetText() for index in range(self.list.GetColumnCount())]
        lstStr = []
        for index in range(self.list.GetItemCount()):
            lstStr.append([self.list.GetItem(index, col=icol).GetText() for icol in range(self.list.GetColumnCount())])
        
        self.thread.append(MyThread.ExportXlsThread(self.panel, lstStr, head, pathXls))
        self.g1.Show(True)
        for item in self.thread:
            item.Start()
    
    def OnShowPop(self, event):
        if self.list.GetItemCount() != 0:
            self.list.PopupMenu(self.panel.popMenuPPFix)
         
    def OnPopItemSelected(self, event):
        try:
            item = self.panel.popMenuPPFix.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"删除":
                self.OnDelete(None)
            if text == u"将所有结果导出为xls文件":
                self.OnOutXls()
        except:
            pass
    
    def OnSelect(self, event):
        # Clear last select result
        if len(self.rankDict) == 0:
            return
        self.ClearTxt()
        strName = self.Text_Select.GetValue()
        # fuzzy select
        lstsql = []
        if self.ckbRank.GetValue():
            lstsql.append("RankSn = '" + self.rankDict.keys()[self.rankDict.values().index(self.cbRank.GetValue())] + "'")
        if self.ckbRankTime.GetValue():
            lstsql.append("RankTime = '" + self.cbRankTime.GetValue() + "-12-1" + "'")
        if self.ckbAddr.GetValue():
            lstsql.append("AddrSn = '" + self.roadDict.keys()[self.roadDict.values().index(self.cbAddr.GetValue())] + "'")
        if self.ckbMarried.GetValue():
            lstsql.append("Married = '" + self.cbMarried.GetValue() + "'")
        if self.ckbSex.GetValue():
            lstsql.append("Sex = '" + self.cbSex.GetValue() + "'")
        
        if self.treeSel is not None:
            curUsn = self.treeDict.keys()[self.treeDict.values().index(self.treeSel)]
            lstsql.append("UnitSn = '" + curUsn + "'")

        strsql = ""
        for item in lstsql:
            strsql += item + " and "
        strsql += "Name"
        
        strResult = PyDatabase.DBSelect(strName, self.DTable, [strsql], 0)
        self.FlashList(strResult)
        self.listIndex = -1
        self.OnDispPPNum()
        
        lstPInfo = [self.cbRank, self.cbRankTime, self.cbAddr, self.cbSex, self.cbMarried, self.tree, self.Text_Name, self.Text_Tel, self.Text_AddrAll]
        [item.Enable(True) for item in lstPInfo]
    
    def FlashList(self, strResult):
        self.list.DeleteAllItems()
        for row in strResult:
            index = self.list.InsertStringItem(1000, "A") 
            self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体')) 
            self.list.SetStringItem(index, 0, `index+1`)
            [self.list.SetStringItem(index, i+1, row[i+1]) for i in range(3)]
            self.list.SetStringItem(index, 4, self.rankDict[row[4]])
            self.list.SetStringItem(index, 5, self.unitDict[row[5]].split('|->')[-1])
            self.list.SetStringItem(index, 6, self.roadDict[row[6]])
            [self.list.SetStringItem(index, i, row[i]) for i in range(7,11)]
        if len(strResult) != 0:
            [self.list.SetColumnWidth(i, wx.LIST_AUTOSIZE) for i in [5,7,9]]
        self.DispColorList(self.list)
       
    def OnBtnPrint(self, event):
        pass
    
    def OnHelp(self, event):        
        try:
            os.startfile(r'help\index.htm') 
        except:
            wx.MessageBox(u'当前版本没有找到帮助文档!', u'提示')
        
    def OnPaint(self, event):
        dc = wx.PaintDC(self.panel)
        dc.BeginDrawing()
        dc.EndDrawing()

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frmMaintain = FrameMaintain()
    frmMaintain.Show()
    app.MainLoop()
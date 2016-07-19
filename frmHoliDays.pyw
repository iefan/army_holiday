#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import PyDatabase
import images
import string
#import MyValidator
from PhrResource import ALPHA_ONLY, DIGIT_ONLY, strVersion, MyListCtrl
import os
import datetime
import PyLunar
import wx.lib.masked as masked
import time

class FrmSubMaintain(wx.Frame):
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        panel = wx.Panel(self)
        
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        HoliDaysFix(panel)
        self.Maximize()
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()

class HoliDaysFix(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)
        
        titleText = wx.StaticText(panel, -1, u"法定假期信息维护")
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))

        btn_Add = wx.Button(panel, -1, u"添加")
        btn_Delete = wx.Button(panel, -1, u"删除")
        btn_Modify = wx.Button(panel, -1, u"初始化")
        btn_Help = wx.Button(panel, -1, u"帮助")
        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [btn_Add, btn_Delete, btn_Modify, btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1)
        
        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        listsizer = wx.BoxSizer(wx.HORIZONTAL)
        listsizer.Add(self.list, 1, wx.ALL | wx.EXPAND, 10)
        
        sampleList = [u"公历", u"农历"]
        self.rbDate = wx.RadioBox(panel, choices=sampleList)
        self.TxtDate = masked.TextCtrl(panel, -1, "", mask="####-##-##")
        self.TxtDate.SetFieldParameters(0, defaultValue='2009')
        self.TxtDate.SetFieldParameters(1, defaultValue='01')
        self.TxtDate.SetFieldParameters(2, defaultValue='01')
        self.TxtDate.SetValue(datetime.datetime.now().strftime("%Y%m%d"))
        
        lbl_Date = wx.StaticText(panel, -1, u"对应公历日期")
        self.TxtDateDisp = masked.TextCtrl(panel, -1, "", mask="####-##-##")
        self.TxtDateDisp.SetEditable(False)

        lbl1 = wx.StaticText(panel, -1, u"假期名称")
        self.Text = wx.TextCtrl(panel, -1, u"") 
        lbl2 = wx.StaticText(panel, -1, u"补假天数")
        self.Text2 = masked.TextCtrl(panel, -1, "", mask="#")

        infosizer = wx.BoxSizer(wx.VERTICAL)
        [infosizer.Add(item, 0, wx.ALL, 5) for item in [self.rbDate, self.TxtDate, lbl_Date, self.TxtDateDisp, lbl1, self.Text, lbl2, self.Text2]]
        
        listsizer.AddSizer(infosizer, 0, wx.ALL, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(titleText, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 
        
        self.DTable = 'HoliDays'
        
        lstHead = [u"序号", u"编号", u"日期", u"假期名称", u"补假天数"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
        [self.list.SetColumnWidth(i, 200) for i in [2,3]]
        
        self.OnSelect(None)
        self.OnRBSelect(None)
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      
        mainSizer.SetSizeHints(panel)
        
        panel.popHoliday = wx.Menu()
        pmList_1 = panel.popHoliday.Append(1181, u"删除")        
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_1)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop)
        
        panel.Show()
        
        panel.Bind( wx.EVT_RADIOBOX, self.OnRBSelect, self.rbDate)
        panel.Bind(wx.EVT_BUTTON, self.OnAdd, btn_Add)       
        panel.Bind(wx.EVT_BUTTON, self.OnDelete, btn_Delete)       
        panel.Bind(wx.EVT_BUTTON, self.OnModify, btn_Modify)       
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
        self.TxtDate.Bind(wx.EVT_KEY_UP, self.EvtChar)
    
    def OnRBSelect(self, event):
        self.CalcDate()
    
    def CalcDate(self):
        dateTuple = self.TxtDate.GetValue().split('-')
        if self.rbDate.GetSelection() == 0:
            try:
                dateDisp = datetime.date(int(dateTuple[0]), int(dateTuple[1]), int(dateTuple[2])).strftime("%Y%m%d")
            except:
                dateDisp = ""
        else:
            dateDisp = PyLunar.get_Calender_date(dateTuple)
            if dateDisp == None:
                dateDisp = ""
            else:
                dateDisp = dateDisp.strftime("%Y%m%d")
        self.TxtDateDisp.SetValue(dateDisp)
    
    def EvtChar(self, event):
        self.CalcDate()
        event.Skip()
        
    def OnShowPop(self, event):
        if self.list.GetItemCount() != 0:
            self.list.PopupMenu(self.panel.popHoliday)
         
    def OnPopItemSelected(self, event):
        try:
            item = self.panel.popHoliday.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"删除":
                self.OnDelete(None)
        except:
            pass
            
    def DispColorList(self, list):
        for i in range(list.GetItemCount()):
            if i%4 == 0: list.SetItemBackgroundColour(i, (233, 233, 247))
            if i%4 == 1: list.SetItemBackgroundColour(i, (247, 247, 247))
            if i%4 == 2: list.SetItemBackgroundColour(i, (247, 233, 233))
            if i%4 == 3: list.SetItemBackgroundColour(i, (233, 247, 247))
            
    def OnSelect(self, event):        
        self.list.DeleteAllItems()                
        strResult = PyDatabase.DBSelect('', self.DTable, ['HolidayTime'], 0)
        for row in strResult:
            index = self.list.InsertStringItem(100, "A")
            self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))  
            self.list.SetStringItem(index, 0, `index+1`)
            self.list.SetStringItem(index, 1, row[1])
            self.list.SetStringItem(index, 2, row[2])
            self.list.SetStringItem(index, 3, row[3])
            self.list.SetStringItem(index, 4, row[4])
        
        self.DispColorList(self.list)
    
    def ClearTxt(self):
        self.TxtDateDisp.SetValue("")
        self.Text.SetValue("")
        self.Text2.SetValue("")
        
    def OnAdd(self, event):
        if self.Text.GetValue().strip() == "" or self.Text2.GetValue().strip() == ""\
            or self.TxtDateDisp.GetValue()[0] == " ":
            wx.MessageBox(u"请填写法定假期时间和补假时间！", u"提示")
            return
        
        newHdDate = self.TxtDateDisp.GetValue()
        newHdName = self.Text.GetValue().strip()
        newDays = self.Text2.GetValue().strip()
        
        if newHdName in [self.list.GetItem(i, col=3).GetText() for i in range(self.list.GetItemCount())]:
            wx.MessageBox(u"已经存在此法定假期！", u"提示")
            return
        
        strResult = PyDatabase.DBSelect("HoliSn like '%%%%'", self.DTable, ['HoliSn'], 1)
        lstIntSn = [int(isn[0][-3:]) for isn in strResult]
        index = 0;  newsn = -1
        for index in range(1, len(strResult)+1):
            if index < lstIntSn[index-1]:
                newsn = index
                break
        if newsn == -1:
            newsn = index+1
        
        listHd = [None]
        listHd.append('HSN' + string.zfill(str(newsn),3))
        listHd.append(newHdDate)
        listHd.append(newHdName)
        listHd.append(newDays)
        # Update Database  
        PyDatabase.DBInsert(listHd, self.DTable)
        # Update List
        index = self.list.InsertStringItem(10000, "A")
        self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))              
        self.list.SetStringItem(index, 0, `index+1`)
        [self.list.SetStringItem(index, i+1, listHd[i+1]) for i in range(4)]
        self.DispColorList(self.list)
        
        self.ClearTxt()
    
    def OnModify(self, event):
        lstGreg = [["1.1", u"元旦",1], ["4.5",u"清明节",1], ["5.1",u"国际劳动节",1],\
            ["10.1",u"国庆节",3]]
        lstLunar = [["1.1",u"春节",3], ["5.5",u"端午节",1],["8.15",u"中秋节",1]]

        today = datetime.datetime.now().strftime('%Y-%m-%d')
        today_lunar = PyLunar.get_ludar_date(datetime.datetime.now()).strftime('%Y-%m-%d')
        lstHoliday = []
        for item in lstGreg:
            tmp = []
            mon = string.zfill(item[0].split('.')[0],2)
            day = string.zfill(item[0].split('.')[1],2)
            tmpday = `time.localtime()[0]`+ "-" + mon + "-" + day
            if tmpday < today:
                tmpday = `time.localtime()[0]+1`+ "-" + mon + "-" + day
            tmp.append(tmpday)
            tmp.append(item[1])
            tmp.append(`item[2]`)
            lstHoliday.append(tmp)
        
        for item in lstLunar:
            tmp = []
            mon = string.zfill(item[0].split('.')[0],2)
            day = string.zfill(item[0].split('.')[1],2)
            tmpday = `time.localtime()[0]`+ "-" + mon + "-" + day
            if tmpday < today_lunar:
                tmpday = PyLunar.get_Calender_date((time.localtime()[0]+1, mon, day)).strftime("%Y-%m-%d")
            else:
                tmpday = PyLunar.get_Calender_date((time.localtime()[0], mon, day)).strftime("%Y-%m-%d")
            tmp.append(tmpday)
            tmp.append(item[1])
            tmp.append(`item[2]`)
            lstHoliday.append(tmp)
        lstHoliday.sort()
        
        self.list.DeleteAllItems()
        # Update database
        PyDatabase.DBDeleteTab(self.DTable)
        
        for row in lstHoliday:
            index = self.list.InsertStringItem(100, "A")
            self.list.SetStringItem(index, 0, `index+1`)
            self.list.SetStringItem(index, 1, "HSN"+ string.zfill(index+1,3))
            self.list.SetStringItem(index, 2, row[0])
            self.list.SetStringItem(index, 3, row[1])
            self.list.SetStringItem(index, 4, row[2])
            
            listHd = [None]
            listHd.append("HSN"+ string.zfill(index+1,3))
            listHd.extend(row)
            # Update Database  
            PyDatabase.DBInsert(listHd, self.DTable)
        
        self.DispColorList(self.list)
        
    def OnDelete(self, event):
        listIndex = self.list.GetFirstSelected()
        if listIndex == -1:
            wx.MessageBox(u"请在列表中点击要删除的法定假期！", u"提示")
            return
        
        # Update Database 
        lstHd = [self.list.GetItem(listIndex, col=i).GetText() for i in range(1,5)]
        PyDatabase.DBDelete(lstHd, self.DTable)
        # Update list
        self.list.DeleteItem(listIndex)
        [self.list.SetStringItem(i, 0, `i+1`) for i in range(self.list.GetItemCount())]
        
        self.DispColorList(self.list)
               
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
    fmmaintain = FrmSubMaintain()
    fmmaintain.Show()
    app.MainLoop()
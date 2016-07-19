#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import PyDatabase
import images
import string
import MyValidator
from PhrResource import ALPHA_ONLY, DIGIT_ONLY, strVersion, MyListCtrl
import os

class FrmSubMaintain(wx.Frame):
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        panel = wx.Panel(self)
        self.Maximize()        
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        RankDaysFix(panel)
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()

class RankDaysFix(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)

        titleText = wx.StaticText(panel, -1, u"军衔信息维护")
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(titleText, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL)
        
        btn_Add = wx.Button(panel, -1, u"添加")
        btn_Delete = wx.Button(panel, -1, u"删除")
        btn_Modify = wx.Button(panel, -1, u"修改")
        btn_Help = wx.Button(panel, -1, u"帮助")
        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [btn_Add, btn_Delete, btn_Modify, btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1)

        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        listsizer = wx.BoxSizer(wx.HORIZONTAL)
        listsizer.Add(self.list, 1, wx.ALL|wx.EXPAND, 5)
        
        lbl1 = wx.StaticText(panel, -1, u"军衔")
        self.Text = wx.TextCtrl(panel, -1, u"一期士官", size=(80, -1)) 
        lbl2 = wx.StaticText(panel, -1, u"天数")
        self.Text2 = wx.TextCtrl(panel, -1, "20", size=(80, -1),validator = MyValidator.MyValidator(DIGIT_ONLY)) 
        infosizer = wx.BoxSizer(wx.VERTICAL)
        [infosizer.Add(item, 0, wx.ALL, 5) for item in [lbl1, self.Text, lbl2, self.Text2]]
        
        listsizer.AddSizer(infosizer, 0, wx.ALL, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 

        self.listIndex = 0
        self.DTable = 'RankDays'
        
        lstHead = [u"序号", u"编号", u"军衔", u"天数"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]

        self.OnSelect(None)
        self.infoTxt = [self.Text, self.Text2]
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      
        mainSizer.SetSizeHints(panel)
        
        panel.popMenuRank = wx.Menu()
        pmList_1 = panel.popMenuRank.Append(1161, u"删除")        
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_1)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop)
        
        panel.Show()

        panel.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.list)
        panel.Bind(wx.EVT_BUTTON, self.OnAdd, btn_Add)       
        panel.Bind(wx.EVT_BUTTON, self.OnDelete, btn_Delete)       
        panel.Bind(wx.EVT_BUTTON, self.OnModify, btn_Modify)       
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
    
    def OnShowPop(self, event):
        if self.list.GetItemCount() != 0:
            self.list.PopupMenu(self.panel.popMenuRank)
         
    def OnPopItemSelected(self, event):
        try:
            item = self.panel.popMenuRank.FindItemById(event.GetId())        
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
        strResult = PyDatabase.DBSelect('', self.DTable, ['LevelRank'], 0)
        for row in strResult:
            index = self.list.InsertStringItem(100, "A")
            self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))  
            self.list.SetStringItem(index, 0, `index+1`)
            self.list.SetStringItem(index, 1, row[1])
            self.list.SetStringItem(index, 2, row[2])
            self.list.SetStringItem(index, 3, row[3])
        
        self.DispColorList(self.list)

    def OnItemSelected(self, event):
        self.listIndex = event.GetIndex()
        item1 = self.list.GetItem(self.listIndex, col=2)
        self.Text.SetValue(item1.GetText())
        item2 = self.list.GetItem(self.listIndex, col=3)
        self.Text2.SetValue(item2.GetText())
                
    def ClearTxt(self):
        self.Text.SetValue("")
        self.Text2.SetValue("")
        
    def OnAdd(self, event):
        if self.Text.GetValue().strip() == "" or self.Text2.GetValue().strip() == "":
            wx.MessageBox(u"请填写军衔和对应天数！", u"提示")
            return
        
        newRank = self.Text.GetValue().strip()
        newDays = self.Text2.GetValue().strip()
        
        if newRank in [self.list.GetItem(i, col=2).GetText() for i in range(self.list.GetItemCount())]:
            wx.MessageBox(u"已经存在此军衔！", u"提示")
            return
        
        strResult = PyDatabase.DBSelect("RankSn like '%%%%'", self.DTable, ['RankSn'], 1)
        lstIntSn = [int(isn[0][-3:]) for isn in strResult]
        index = 0;  newsn = -1
        for index in range(1, len(strResult)+1):
            if index < lstIntSn[index-1]:
                newsn = index
                break
        if newsn == -1:
            newsn = index+1
            
        listRank = [None]
        listRank.append('RSN' + string.zfill(str(newsn),3))
        [listRank.append(item.GetValue()) for item in self.infoTxt]
        # Update Database  
        PyDatabase.DBInsert(listRank, self.DTable)
        # Update List
        index = self.list.InsertStringItem(10000, "A")
        self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))              
        self.list.SetStringItem(index, 0, `index+1`)
        [self.list.SetStringItem(index, i+1, listRank[i+1]) for i in range(3)]
        
        self.ClearTxt()
        self.DispColorList(self.list)
        self.list.SetFocus()
    
    def OnModify(self, event): 
        if self.Text.GetValue().strip() == "" or self.Text2.GetValue().strip() == "":
            wx.MessageBox(u"请在列表中点击所要修改的某一军衔和对应天数！", u"提示")
            return
        
        oldRank = self.list.GetItem(self.listIndex, col=2).GetText()
        oldDays = self.list.GetItem(self.listIndex, col=3).GetText()
        newRank = self.Text.GetValue().strip()
        newDays = self.Text2.GetValue().strip()
        
        if newRank == oldRank and newDays == oldDays:
            return
        
        if newDays in [self.list.GetItem(i, col=2).GetText() for i in range(self.list.GetItemCount())] \
            and newRank in [self.list.GetItem(i, col=1).GetText() for i in range(self.list.GetItemCount())]:
            wx.MessageBox(u"已经存在此军衔，请检查！", u"提示")
            return
        
        oldRanklst = [self.list.GetItem(self.listIndex, col=1).GetText(), oldRank, oldDays]
        newRanklst = [self.list.GetItem(self.listIndex, col=1).GetText(), newRank, newDays]
        # Update Database
        PyDatabase.DBUpdate(oldRanklst, newRanklst, self.DTable)
        # Update list
        [self.list.SetStringItem(self.listIndex, i+1, newRanklst[i]) for i in range(3)]
                        
        self.ClearTxt()
        self.list.SetFocus()
    
    def OnDelete(self, event):
        if self.Text.GetValue().strip() == "" or self.Text2.GetValue().strip() == "" or self.listIndex == -1:
            wx.MessageBox(u"请在列表中点击要删除的军衔及对应天数！", u"提示")
            return
        
        rankSn = self.list.GetItem(self.listIndex, col=1).GetText()
        strResult = PyDatabase.DBSelect(rankSn, "PersonInfo", ['RankSn'], 2)
        if len(strResult) > 0:
            wx.MessageBox(u"当前数据库中存在该军衔的人员，无法删除该军衔！", u"提示")
            return
    
        # Update Database 
        lstRank = [self.list.GetItem(self.listIndex, col=i).GetText() for i in range(1,4)]
        PyDatabase.DBDelete(lstRank, self.DTable)
        # Update list
        self.list.DeleteItem(self.listIndex)
        self.listIndex = -1        
        [self.list.SetStringItem(i, 0, `i+1`) for i in range(self.list.GetItemCount())]
        
        self.ClearTxt()
        self.DispColorList(self.list)
        self.list.SetFocus()
       
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
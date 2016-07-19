#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
#import wx.lib.buttons as buttons       
import PyDatabase
import images
import string
import datetime
import operateWord
import pickle
import MyThread
import  time
from PhrResource import strVersion, g_UnitSnNum
try:
    import huBarcode.code128 as hc
except:
    pass
import os

class FrameMaintain(wx.Frame):               
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        self.panel = panel = wx.Panel(self)
        self.Maximize() 
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        HoliPHFix(panel)
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()
        
class HoliPHFix(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)

        titleText = wx.StaticText(panel, -1, u"休假人员信息维护")
        self.dispText = wx.StaticText(panel, -1, u"当前浏览人数：")
        self.g1 = wx.Gauge(panel, -1, 50, size=(-1,10))
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer0 = wx.BoxSizer(wx.HORIZONTAL)
        topsizer0.Add(titleText, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL)
        numsizer = wx.BoxSizer(wx.HORIZONTAL)
        numsizer.Add(self.dispText, 0, wx.ALL| wx.ALIGN_BOTTOM)
        numsizer.Add(self.g1, 1, wx.ALL|wx.EXPAND| wx.ALIGN_BOTTOM)
        
        self.ckbRankLbl = wx.CheckBox(panel, -1, u"军衔")
        self.cbRank = wx.ComboBox(panel)
        self.ckbDateLbl = wx.CheckBox(panel, -1, u"时间：起")
        self.dateStart = wx.DatePickerCtrl(panel, style = wx.DP_DROPDOWN | wx.DP_SHOWCENTURY)
        lbl1 = wx.StaticText(panel, -1, u"止")
        self.dateEnd = wx.DatePickerCtrl(panel, style = wx.DP_DROPDOWN | wx.DP_SHOWCENTURY)
        self.ckbIsHEnd = wx.CheckBox(panel, -1, u"正在休假")
        
        lblname = wx.StaticText(panel, -1, u"姓名")
        self.Text_Select = wx.TextCtrl(panel, -1, u"")
        btn_Select = wx.Button(panel, -1, u"查询", size=(60,-1))
        btn_Select.SetDefault()
        
        topsizer = wx.BoxSizer(wx.HORIZONTAL)        
        itemSelect = [self.ckbRankLbl, self.cbRank, self.ckbDateLbl, self.dateStart, lbl1, self.dateEnd]
        for item in itemSelect:
            topsizer.Add(item, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
            if item.GetClassName() == u"wxComboBox":
                topsizer.Add((20, -1))
        topsizer.Add((-1,-1), 1, wx.ALL|wx.EXPAND)
        topsizer.Add(self.ckbIsHEnd, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL)
        topsizer.Add((-1,-1), 1, wx.ALL|wx.EXPAND)
        topsizer.Add(lblname, 0, wx.ALL| wx.ALIGN_RIGHT| wx.ALIGN_CENTER_VERTICAL)
        topsizer.Add(self.Text_Select, 0, wx.ALL| wx.ALIGN_RIGHT)
        topsizer.Add(btn_Select, 0, wx.ALL| wx.ALIGN_RIGHT)

        lbl2 = wx.StaticText(panel, -1, u"销假日期：")
        self.dateInfactEnd = wx.DatePickerCtrl(panel, style = wx.DP_DROPDOWN | wx.DP_SHOWCENTURY)
        lbl3 = wx.StaticText(panel, -1, u"假期详情：")
        self.Text_Demo = wx.TextCtrl(panel, -1)
        self.Text_Demo.SetEditable(False)
        
        lbl4 = wx.StaticText(panel, -1, u"备　　注：")
        self.Text_Demo2 = wx.TextCtrl(panel)
        
        lbl5 = wx.StaticText(panel, -1, u"请输入休假编号：")
        self.Text_HDSn = wx.TextCtrl(panel, -1, size=(160,-1), style=wx.WANTS_CHARS)
        
        infosizer1 = wx.BoxSizer(wx.HORIZONTAL)
        infosizer1.Add(lbl2, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer1.Add(self.dateInfactEnd, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer1.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        infosizer1.Add(lbl5, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer1.Add(self.Text_HDSn, 0, wx.ALL, 5)
        infosizer1.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        
        infosizer2 = wx.BoxSizer(wx.HORIZONTAL)
        infosizer2.Add(lbl3, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer2.Add(self.Text_Demo, 1, wx.ALL| wx.EXPAND, 0)
        
        infosizer3 = wx.BoxSizer(wx.HORIZONTAL)
        infosizer3.Add(lbl4, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer3.Add(self.Text_Demo2, 1, wx.ALL| wx.EXPAND, 0)
        
        self.btn_HoliEnd = wx.Button(panel, -1, u"销假") 
        btn_Delete = wx.Button(panel, -1, u"删除") 
        btn_Modify = wx.Button(panel, -1, u"修改备注") 
        btn_Help = wx.Button(panel, -1, u"帮助")        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [self.btn_HoliEnd, btn_Delete, btn_Modify, btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1) 
        
        self.list = wx.ListCtrl(panel, -1, style=wx.LC_REPORT| wx.LC_VRULES | wx.LC_HRULES|wx.LC_SINGLE_SEL|wx.LC_SINGLE_SEL)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        listsizer = wx.BoxSizer(wx.VERTICAL)
        listsizer.AddSizer(topsizer,0, wx.ALL| wx.EXPAND, 0)
        listsizer.Add(self.list, 1, wx.ALL|wx.EXPAND, 0)
        
        lbltree = wx.StaticText(panel, -1, u"单位树")
        self.tree = wx.TreeCtrl(panel, size=(160, 300))
        self.root = self.tree.AddRoot(u"单位")
        
        treesizer = wx.BoxSizer(wx.VERTICAL)
        treesizer.AddSizer(lbltree, 0, wx.ALL, 5)
        treesizer.Add(self.tree, 1, wx.ALL| wx.EXPAND)
        
        midsizer = wx.BoxSizer(wx.HORIZONTAL)
        midsizer.AddSizer(treesizer, 0, wx.ALL| wx.EXPAND, 5)
        midsizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer0, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.AddSizer(numsizer, 0, wx.ALL| wx.ALIGN_LEFT | wx.EXPAND, 0)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(midsizer, 1, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 2)
        mainSizer.AddSizer(infosizer1, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.AddSizer(infosizer2, 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.AddSizer(infosizer3, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      

        panel.popMenuHoliMan = wx.Menu()
        pmList_1 = panel.popMenuHoliMan.Append(1141, u"删除")
        panel.popMenuHoliMan.AppendSeparator()
        pmList_2 = panel.popMenuHoliMan.Append(1142, u"导出批假通知单")
        pmList_3 = panel.popMenuHoliMan.Append(1143, u"将所有结果导出为xls文件")
        self.list.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_1)
        self.list.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_2)
        self.list.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_3)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop)
        
        self.treeDict = {}
        self.treeSel = None
        self.listIndex = -1
        self.unitDict = {}
        self.rankDict = {}
        self.DTable = 'PersonInfo'
        self.DTable2 = 'PPHdRecord'
        
        panel.Bind(MyThread.EVT_UPDATE_BARGRAPH, self.OnUpdate)
        panel.Bind(MyThread.EVT_EXPORT_XLS, self.OnExport)
        
        self.thread = []
        
        lstHead = [u"序号", u"编号", u"姓名",u"请假事由",u"天数", u"离队日期",u"到假日期",u"销假日期",u"经手",u"批假", u"剩余", u"假期详情", u"类别", u"备注"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
        self.list.SetColumnWidth(0, 50)


        infoItem1 = [self.ckbDateLbl, self.ckbRankLbl]
        [item1.SetValue(False) for item1 in infoItem1]
        self.Text_Select.Enable(True)
        
        self.InitData()

        panel.Show()
        self.Text_HDSn.Bind(wx.EVT_KEY_UP, self.EvtChar)
        itemchecklst = [self.ckbRankLbl, self.ckbDateLbl, self.ckbIsHEnd]
        [panel.Bind(wx.EVT_CHECKBOX, self.OnCkbInfo, item) for item in itemchecklst]
        panel.Bind( wx.EVT_COMBOBOX, self.OnCkbInfo, self.cbRank)
        panel.Bind(wx.EVT_DATE_CHANGED, self.OnDateChanged, self.dateStart)
        panel.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelChanged, self.tree)
        panel.Bind(wx.EVT_TREE_ITEM_ACTIVATED, self.OnActivate, self.tree)
        panel.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.list)
        panel.Bind(wx.EVT_BUTTON, self.OnHoliEnd, self.btn_HoliEnd)       
        panel.Bind(wx.EVT_BUTTON, self.OnDelete, btn_Delete)       
        panel.Bind(wx.EVT_BUTTON, self.OnModify, btn_Modify)       
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_BUTTON, self.OnSelect, btn_Select)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
    
    def EvtChar(self, event):
        if len(self.Text_HDSn.GetValue()) >= 16:
            self.Text_HDSn.SelectAll()
#            self.Text_HDSn.SetValue(self.Text_HDSn.GetValue()[16:])
        self.Text_Select.SetValue("")
        self.ckbRankLbl.SetValue(False)
        self.treeSel = None
        self.tree.SelectItem(self.root)
        self.OnSelect(None)
        
        txtInput = self.Text_HDSn.GetValue().strip().upper()
        for index in range(self.list.GetItemCount()-1, -1, -1):
            if txtInput not in self.list.GetItem(index, col=1).GetText():
                self.list.DeleteItem(index)
                [self.list.SetStringItem(i, 0, `i+1`) for i in range(self.list.GetItemCount())]

        event.Skip()
    
    def OnCkbInfo(self, event):
        self.OnSelect(None)
        
    def OnDispPPNum(self):
        strDisp = self.list.GetItemCount()
        self.dispText.SetLabel(u"当前浏览人数：" + `strDisp`)
        
    def OnExport(self, evt):
        if evt.flag:
            for item in self.thread:
                item.Stop()
                
            running = 1

            while running:
                running = 0
                for t in self.thread:
                    running = running + t.IsRunning()
                time.sleep(0.1)
            self.g1.Show(False)
            wx.MessageBox(u"导出成功！", u"提示")
    
    def OnUpdate(self, evt):
        self.g1.SetValue(evt.count)
        
    def InitData(self):
        lstRank = PyDatabase.DBSelect(u"ID like '%%%%'", "RankDays", ['RankSn', 'LevelRank'], 1)
        list_Rank = []
        for item in lstRank:
            self.rankDict[item[0]] = item[1]
            list_Rank.append(item[1])
        self.cbRank.SetItems(list_Rank)
        self.cbRank.Select(0)
        
        self.InitTree()
        self.OnSelect(None)
        self.g1.Show(False)
    
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
        head.insert(1, u"单位")
        head.insert(4, u"军衔")
        lstStr = []
        for index in range(self.list.GetItemCount()):
            tmplst = [self.list.GetItem(index, col=0).GetText()]
            tmpSn = self.list.GetItem(index, col=1).GetText()[:g_UnitSnNum]
            strSelect = "Sn = '" + tmpSn + "'"
            tmpResult = PyDatabase.DBSelect(strSelect, self.DTable, ['RankSn', 'UnitSn'], 1)[0]
            tmplst.append(self.tree.GetItemText(self.treeDict[tmpResult[1]]))
            tmplst.extend([self.list.GetItem(index, col=icol).GetText() for icol in range(1, 3)])
            tmplst.append(self.rankDict[tmpResult[0]])
            tmplst.extend([self.list.GetItem(index, col=icol).GetText() for icol in range(3, self.list.GetColumnCount())])
            lstStr.append(tmplst)
            
        self.thread = []
        self.thread.append(MyThread.CalcGaugeThread(self.panel, 0))
        self.thread.append(MyThread.ExportXlsThread(self.panel, lstStr, head, pathXls))
        self.g1.Show(True)
        for item in self.thread:
            item.Start()
        
    def OnOutWord(self):
        if self.listIndex == -1:
            wx.MessageBox(u"请选中表中一行！", u"提示")
            return
        
        lstword = []
        lstword.append(self.list.GetItem(self.listIndex, col=1).GetText())
        lstword.append(self.list.GetItem(self.listIndex, col=2).GetText())
        lstword.append(self.tree.GetItemText( self.tree.GetSelection()))
        lstword.append(self.list.GetItem(self.listIndex, col=11).GetText())
        lstword.append(self.list.GetItem(self.listIndex, col=5).GetText())
        lstword.append(self.list.GetItem(self.listIndex, col=6).GetText())
        try:
            pkl_file = open('ByData.pkl', 'rb')
        except:
            wx.MessageBox(u"请先在单位维护界面填写批假单位！", u"提示")
            return
        
        data = pickle.load(pkl_file)
        pkl_file.close()
        if data['ByUnit'] == "":
            wx.MessageBox(u"请先在单位维护界面填写批假单位！", u"提示")
            return
        lstword.append(data['ByUnit'])
        operateWord.MyGenWordList(lstword)

    def OnDateChanged(self, event): 
        pass
    
    def InitTree(self):
        strResult = PyDatabase.DBSelect('', 'UnitTab', ['UnitSn'], 0) 
#        self.unitDict = []
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
        
    def OnActivate(self, event):
        infoItem1 = [self.ckbDateLbl, self.ckbRankLbl]
        [item1.SetValue(False) for item1 in infoItem1]
        
        item = event.GetItem()
        if item in self.treeDict.values():
            curUsn = self.treeDict.keys()[self.treeDict.values().index(item)]
#            strResult = PyDatabase.DBSelect(curUsn, self.DTable, ['UnitSn'], 2)
            self.treeSel = item
        else:
            self.treeSel = None
        self.OnSelect(None)
#            strResult = PyDatabase.DBSelect("", self.DTable, ['Sn'], 0)            
#        self.FlashList(strResult)
        
    def OnShowPop(self, event):
        if self.list.GetItemCount() != 0:
            self.list.PopupMenu(self.panel.popMenuHoliMan)
         
    def OnPopItemSelected(self, event):
        try:
            item = self.panel.popMenuHoliMan.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"删除":
                if wx.OK == wx.MessageBox(u"确定要删除当前记录吗？", u"提示", wx.ICON_QUESTION|wx.OK| wx.CANCEL):
                    self.OnDelete(None)
            if text == u"导出批假通知单":                
                self.OnOutWord()
            if text == u"将所有结果导出为xls文件":
                self.OnOutXls()            
        except:
            pass
            
    def DispColorList(self, list):
        for i in range(list.GetItemCount()):
            if i%4 == 0: list.SetItemBackgroundColour(i, (233, 233, 247))
            if i%4 == 1: list.SetItemBackgroundColour(i, (247, 247, 247))
            if i%4 == 2: list.SetItemBackgroundColour(i, (247, 233, 233))
            if i%4 == 3: list.SetItemBackgroundColour(i, (233, 247, 247))
        
    def OnSelect(self, event):
        # Clear last select result
        if len(self.rankDict) == 0:
            return
        
        strName = self.Text_Select.GetValue()
        # fuzzy select
        lstsql = []
        if self.ckbRankLbl.GetValue():
            lstsql.append("RankSn = '" + self.rankDict.keys()[self.rankDict.values().index(self.cbRank.GetValue())] + "'")
        if self.treeSel is not None:
            curUsn = self.treeDict.keys()[self.treeDict.values().index(self.treeSel)]
            lstsql.append("UnitSn = '" + curUsn + "'")

        strsql = ""
        for item in lstsql:
            strsql += item + " and "
        strsql += "Name"
        
        strResult = PyDatabase.DBSelect(strName, self.DTable, [strsql], 0)
        
        lstItem = ["Sn", "Name"]
        strResultHoliday = PyDatabase.DBSelect("", self.DTable2, ['Sn'], 0)
        
        lstResult = []
        if self.ckbDateLbl.GetValue():
            datestart = self.dateStart.GetLabel()
            dateend = self.dateEnd.GetLabel()
            startday = datetime.date(int(datestart.split('-')[0]), int(datestart.split('-')[1]), int(datestart.split('-')[2]))
            endday = datetime.date(int(dateend.split('-')[0]), int(dateend.split('-')[1]), int(dateend.split('-')[2]))
            for item in strResultHoliday:
                date1 = datetime.date(int(item[4].split('-')[0]), int(item[4].split('-')[1]), int(item[4].split('-')[2]))
                if (startday-date1).days <= 0 and (endday - date1).days>=0:
                    lstResult.append(item)
                    
        elif self.ckbIsHEnd.GetValue():
            for item in strResultHoliday:
                if item[6] == "":
                    lstResult.append(item)                
        else:
            lstResult.extend(strResultHoliday)
        
        dispResult = []
        for item1 in lstResult:
            tmp = []
            for item2 in strResult:
                if item1[1][:g_UnitSnNum] == item2[1]:
                    tmp.extend([item1[1], item2[2]])
                    tmp.extend(item1[2:])
                    dispResult.append(tmp)
        self.FlashList(dispResult)
        self.listIndex = -1
        self.OnDispPPNum()
    
    def FlashList(self, strResult):
        self.list.DeleteAllItems()
        sortResult = []
        for row in strResult:
            tmplst = [row.pop(4)]
            tmplst.extend(row)
            sortResult.append(tmplst)
        sortResult.sort()
        strResult = []
        for row in sortResult:
            tmplst = row[1:5]
            tmplst.append(row[0])
            tmplst.extend(row[5:])
            strResult.append(tmplst)
            
        for row in strResult:
            index = self.list.InsertStringItem(10000, "A") 
            self.list.SetStringItem(index, 0, `index+1`)
            [self.list.SetStringItem(index, i+1, row[i]) for i in range(len(row))]
            self.list.SetItemTextColour(index, wx.BLACK)
            self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
            dateEnd = row[5]
            dateNow = time.localtime()[:3]
            endday = datetime.date(int(dateEnd.split('-')[0]), int(dateEnd.split('-')[1]), int(dateEnd.split('-')[2]))
            today = datetime.date(dateNow[0], dateNow[1], dateNow[2])
                        
            if (endday - today).days > 0 and (endday - today).days <= 3:
                self.list.SetItemTextColour(index, (0, 0, 255))
            if (endday - today).days == 0 and row[6] == "":
                self.list.SetItemTextColour(index, (0, 0, 255))
            if (endday - today).days < 0 and row[6] == "":
                self.list.SetItemTextColour(index, (0, 0, 255))
                self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.ITALIC, wx.NORMAL, face=u'宋体'))
                
            if row[6] != "":
                dateInfactEnd = row[6]
                endInfactday = datetime.date(int(dateInfactEnd.split('-')[0]), int(dateInfactEnd.split('-')[1]), int(dateInfactEnd.split('-')[2]))
                if (endInfactday - endday).days > 0:
                    self.list.SetItemTextColour(index, (255, 0, 0))
                elif (endInfactday - endday).days < 0:
                    self.list.SetItemTextColour(index, (0, 100,0))
                    
        if len(strResult) != 0:
            [self.list.SetColumnWidth(i, wx.LIST_AUTOSIZE) for i in [1,3,5,6,7, 11,13]]
        self.DispColorList(self.list)

    def OnItemSelected(self, event):
        self.listIndex = event.GetIndex()
        strSn = self.list.GetItem(self.listIndex, col=1).GetText()[:g_UnitSnNum]
        
        strSelect = "Sn = '" + strSn + "'"
        strResult = PyDatabase.DBSelect(strSelect, self.DTable, ['RankSn', 'UnitSn'], 1)[0]
        self.cbRank.SetValue(self.rankDict[strResult[0]])
        self.tree.SelectItem(self.treeDict[strResult[1]])
        startdate = self.list.GetItem(self.listIndex, col=5).GetText().split('-')
        enddate = self.list.GetItem(self.listIndex, col=6).GetText().split('-')
        self.dateStart.SetValue(wx.DateTimeFromDMY(int(startdate[2]), int(startdate[1])-1, int(startdate[0]), 0, 0, 0))
        self.dateEnd.SetValue(wx.DateTimeFromDMY(int(enddate[2]), int(enddate[1])-1, int(enddate[0]), 0, 0, 0))
        self.Text_Demo.SetValue(self.list.GetItem(self.listIndex, col=11).GetText())
        self.Text_Demo2.SetValue(self.list.GetItem(self.listIndex, col=13).GetText())
        
        if self.list.GetItem(self.listIndex, col=7).GetText() == "":
            self.btn_HoliEnd.Enable(True)
        else:
            self.btn_HoliEnd.Enable(False)
        
        path = os.getcwd() +  "\\tmp.png"
        try:            
            hc.Code128Encoder(str(self.list.GetItem(self.listIndex, col=1).GetText())).save(path,1) 
        except:
            if os.path.exists(path):
                os.system("del "+ path)

    def OnHoliEnd(self, event):
        if self.listIndex == -1:
            wx.MessageBox(u"请点击选中已经被批假的人员。", u"提示")
            return
        
        dateEnd = self.list.GetItem(self.listIndex, col=6).GetText()
        dateEndInfact = self.dateInfactEnd.GetLabel()
        oldPersonlst = [self.list.GetItem(self.listIndex, col=i+1).GetText() for i in range(self.list.GetColumnCount()-1)]
        oldPersonlst.pop(1)
        newPersonlst = []
        newPersonlst.extend(oldPersonlst)
        newPersonlst[5] = dateEndInfact
        
        endday = datetime.date(int(dateEnd.split('-')[0]), int(dateEnd.split('-')[1]), int(dateEnd.split('-')[2]))
        infactendday = datetime.date(int(dateEndInfact.split('-')[0]), int(dateEndInfact.split('-')[1]), int(dateEndInfact.split('-')[2]))
        newPersonlst[8] = `(endday - infactendday).days`
        newPersonlst[-1] = self.Text_Demo2.GetValue().strip()
        
        strtip = u"您确定要为 " + self.list.GetItem(self.listIndex, col=2).GetText() + u" 销假吗？"
        
        if wx.OK == wx.MessageBox(strtip, u"询问", wx.OK | wx.CANCEL | wx.ICON_QUESTION):
            if wx.CANCEL == wx.MessageBox(u"请注意，是 " + self.list.GetItem(self.listIndex, col=2).GetText() + u"，确定吗？", u"询问", wx.OK | wx.CANCEL | wx.ICON_QUESTION):
                return
        else:            
            return
        
        # Update Database
        PyDatabase.DBUpdate(oldPersonlst, newPersonlst, self.DTable2)
        # Update the list
        self.list.SetStringItem(self.listIndex, 7, newPersonlst[5])
        self.list.SetStringItem(self.listIndex, 10, newPersonlst[8])
        self.list.SetStringItem(self.listIndex, 13, newPersonlst[-1])
        [self.list.SetColumnWidth(i, wx.LIST_AUTOSIZE) for i in [1,3,5,6,7, 11,13]]
        
        if int(newPersonlst[8]) < 0:
            self.list.SetItemTextColour(self.listIndex, (255, 0, 0))
        elif int(newPersonlst[8]) > 0:
            self.list.SetItemTextColour(self.listIndex, (0, 100,0))
        else:
            self.list.SetItemTextColour(self.listIndex, wx.BLACK)
            
        self.list.SetItemFont(self.listIndex, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        self.list.Select(self.listIndex)
        self.btn_HoliEnd.Enable(False)
        self.list.SetFocus()
        
    def OnModify(self, event):
        if self.listIndex == -1:
            wx.MessageBox(u"请点击选中已经被批假的人员。", u"提示")
            return
                
        oldPersonlst = [self.list.GetItem(self.listIndex, col=i+1).GetText() for i in range(self.list.GetColumnCount()-1)]
        oldPersonlst.pop(1)
        newPersonlst = []
        newPersonlst.extend(oldPersonlst)
        newPersonlst[-1] = self.Text_Demo2.GetValue().strip()
        
        # Update Database
        PyDatabase.DBUpdate(oldPersonlst, newPersonlst, self.DTable2)        
        # Update list
        self.list.SetStringItem(self.listIndex, 13, newPersonlst[-1])
        self.list.Select(self.listIndex)
        self.list.SetFocus()
    
    def OnDelete(self, event):
        if self.listIndex == -1:
            wx.MessageBox(u"请点击选中已经被批假的人员。", u"提示")
            return
        
        oldPersonlst = [self.list.GetItem(self.listIndex, col=i+1).GetText() for i in range(self.list.GetColumnCount()-1)]
        oldPersonlst.pop(1)
        # Update Database
        PyDatabase.DBDelete(oldPersonlst, self.DTable2)
        # Update list
        self.list.DeleteItem(self.listIndex)
        [self.list.SetStringItem(i, 0, `i+1`) for i in range(self.list.GetItemCount())]
        self.listIndex = -1
        self.DispColorList(self.list)
        self.Text_Demo.SetValue("")
        self.Text_Demo2.SetValue("")
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
    frmMaintain = FrameMaintain()
    frmMaintain.Show()
    app.MainLoop()
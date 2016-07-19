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
import types
import os

class FrameMaintain(wx.Frame):               
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        self.panel = panel = wx.Panel(self)
        self.Maximize() 
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        PersonFix(panel)
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()
        
class PersonFix(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)
        
        titleText = wx.StaticText(panel, -1, u"人员信息维护")
        self.dispText = wx.StaticText(panel, -1, u"当前浏览人数：")
        self.g1 = wx.Gauge(panel, -1, 50, size=(-1,10))
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer0 = wx.BoxSizer(wx.HORIZONTAL)
        topsizer0.Add(titleText, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL)
        numsizer = wx.BoxSizer(wx.HORIZONTAL)
        numsizer.Add(self.dispText, 0, wx.ALL| wx.ALIGN_BOTTOM)
        numsizer.Add(self.g1, 1, wx.ALL|wx.EXPAND| wx.ALIGN_BOTTOM)

        self.ckbRank = wx.CheckBox(panel, -1, u"军衔")
        self.cbRank = wx.ComboBox(panel, size=(100,-1))
        self.ckbRankTime = wx.CheckBox(panel, -1, u"时间")
        self.cbRankTime = wx.ComboBox(panel, size=(60,-1))
        self.ckbAddr = wx.CheckBox(panel, -1, u"籍贯")
        self.cbAddr = wx.ComboBox(panel, size=(60,-1))
        self.ckbSex = wx.CheckBox(panel, -1, u"性别")
        self.cbSex = wx.ComboBox(panel, -1, u"男", choices=[u"男", u"女"], size=(50,-1))
        self.ckbMarried = wx.CheckBox(panel, -1, u"婚否")
        self.cbMarried = wx.ComboBox(panel, -1, u"未婚", choices=[u"未婚", u"已婚"], size=(50,-1))
        lblname = wx.StaticText(panel, -1, u"姓名：")

        self.Text_Select = wx.TextCtrl(panel, size=(100,-1))
        btn_Select = wx.Button(panel, -1, u"查询", size=(60,-1))
        btn_Select.SetDefault()
        
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        itemSelect = [self.ckbRank, self.cbRank, (-1,-1), self.ckbRankTime, self.cbRankTime, (-1,-1), self.ckbAddr, self.cbAddr, (-1,-1), self.ckbSex, self.cbSex, (-1,-1), self.ckbMarried, self.cbMarried]
        for item in itemSelect:
            if types.TupleType == type(item):
                topsizer.Add(item, 1, wx.ALL| wx.EXPAND)
            else:                
                topsizer.Add(item, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL)

        topsizer.Add((-1,-1), 5, wx.ALL| wx.EXPAND)
        topsizer.Add(lblname, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL)
        topsizer.Add(self.Text_Select, 0, wx.ALL)
        topsizer.Add(btn_Select, 0, wx.ALL)

        lbl5 = wx.StaticText(panel, -1, u"姓　　名：")
        self.Text_Name = wx.TextCtrl(panel)
        lbl6 = wx.StaticText(panel, -1, u"联系电话：")
        self.Text_Tel = wx.TextCtrl(panel,validator = MyValidator.MyValidator(DIGIT_ONLY))
        lbl7 = wx.StaticText(panel, -1, u"详细地址：")
        self.Text_AddrAll = wx.TextCtrl(panel)
        
        iteminfo = [lbl5, self.Text_Name]
        infosizer1 = wx.BoxSizer(wx.HORIZONTAL)
        for item in iteminfo:
            infosizer1.Add(item, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer1.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        infosizer1.Add(lbl6, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer1.Add(self.Text_Tel, 3, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer1.Add((-1,-1), 1, wx.ALL| wx.EXPAND)

        infosizer2 = wx.BoxSizer(wx.HORIZONTAL)
        infosizer2.Add(lbl7, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)
        infosizer2.Add(self.Text_AddrAll, 1, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 5)

        btn_Add = wx.Button(panel, -1, u"添加") 
        btn_Delete = wx.Button(panel, -1, u"删除") 
        btn_Modify = wx.Button(panel, -1, u"修改") 
        btn_Help = wx.Button(panel, -1, u"帮助")        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [btn_Add, btn_Delete, btn_Modify, btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1) 
        
        self.list = wx.ListCtrl(panel, -1, style=wx.LC_REPORT| wx.LC_VRULES | wx.LC_HRULES)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        listsizer = wx.BoxSizer(wx.VERTICAL)
        listsizer.AddSizer(topsizer,0, wx.ALL| wx.EXPAND, 0)
        listsizer.Add(self.list, 1, wx.ALL|wx.EXPAND, 0)
        
        lbltree = wx.StaticText(panel, -1, u"单位树")
        self.ckbModify = wx.CheckBox(panel, -1, u"编辑模式")
        self.tree = wx.TreeCtrl(panel, size=(180, -1))
        self.root = self.tree.AddRoot(u"单位")
        
        treetopsizer = wx.BoxSizer(wx.HORIZONTAL)
        treetopsizer.Add(lbltree, 0, wx.ALL)
        treetopsizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        treetopsizer.Add(self.ckbModify, 0, wx.ALL)
        treesizer = wx.BoxSizer(wx.VERTICAL)
        treesizer.AddSizer(treetopsizer, 0, wx.ALL| wx.EXPAND, 5)
        treesizer.Add(self.tree, 1, wx.ALL|wx.EXPAND, 0)        

        midsizer = wx.BoxSizer(wx.HORIZONTAL)
        midsizer.AddSizer(treesizer, 0, wx.ALL| wx.EXPAND, 5)
        midsizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer0, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.AddSizer(numsizer, 0, wx.ALL| wx.ALIGN_LEFT | wx.EXPAND, 0)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(midsizer, 1, wx.ALL| wx.EXPAND, 0)
        mainSizer.AddSizer(infosizer1, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.AddSizer(infosizer2, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      
        
        panel.popMenuPPFix = wx.Menu()
        pmList_1 = panel.popMenuPPFix.Append(1151, u"删除")
        panel.popMenuPPFix.AppendSeparator()
        pmList_2 = panel.popMenuPPFix.Append(1152, u"将所有结果导出为xls文件")
        self.list.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_1)
        self.list.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_2)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop)
        
        self.treeDict = {}
        self.treeSel = None
        self.listIndex = 0
        self.unitDict = {}
        self.rankDict = {}
        self.roadDict = {}
        self.DTable = 'PersonInfo'
        
        panel.Bind(MyThread.EVT_UPDATE_BARGRAPH, self.OnUpdate)
        panel.Bind(MyThread.EVT_EXPORT_XLS, self.OnExport)
        
        self.thread = []
        self.thread.append(MyThread.CalcGaugeThread(panel, 0))
        
        self.lstHead = [u"序号", u"编号", u"姓名",u"性别",u"军衔",u"单位",u"籍贯", u"详细地址",u"婚否",u"电话", u"军衔时间"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(self.lstHead)), self.lstHead)]
        self.list.SetColumnWidth(0, 50)
        self.list.SetColumnWidth(3, 60)
        
        infoPerson = [self.cbSex, self.cbMarried, self.cbRank, self.cbAddr]
        [item.SetEditable(False) for item in infoPerson]
        
        self.InitData()

        panel.Show()
        itemchecklst = [self.ckbRank, self.ckbRankTime, self.ckbAddr, self.ckbSex, self.ckbMarried]
        [panel.Bind(wx.EVT_CHECKBOX, self.OnCkbInfo, item) for item in itemchecklst]
        panel.Bind(wx.EVT_CHECKBOX, self.OnCkbModify, self.ckbModify)
        panel.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelChanged, self.tree)
        panel.Bind(wx.EVT_TREE_ITEM_ACTIVATED, self.OnActivate, self.tree)
        panel.Bind(wx.EVT_COMBOBOX, self.OnCbChanged, self.cbRank)
        panel.Bind(wx.EVT_COMBOBOX, self.OnCbChanged, self.cbRankTime)
        panel.Bind(wx.EVT_COMBOBOX, self.OnCbChanged, self.cbAddr)
        panel.Bind(wx.EVT_COMBOBOX, self.OnCbChanged, self.cbSex)
        panel.Bind(wx.EVT_COMBOBOX, self.OnCbChanged, self.cbMarried)
        panel.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.list)
        panel.Bind(wx.EVT_BUTTON, self.OnAdd, btn_Add)       
        panel.Bind(wx.EVT_BUTTON, self.OnDelete, btn_Delete)       
        panel.Bind(wx.EVT_BUTTON, self.OnModify, btn_Modify)       
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_BUTTON, self.OnSelect, btn_Select)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
    
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
        lstType = PyDatabase.DBSelect(u"ID like '%%%%'", "RankDays", ['RankSn', 'LevelRank'], 1)
        list_Type = []
        for item in lstType:
            self.rankDict[item[0]] = item[1]
            list_Type.append(item[1])
        self.cbRank.SetItems(list_Type)
        self.cbRank.Select(0)
        
        lstRTime = []
        for i in range(2020, 2000, -1):
            lstRTime.append(`i`)
        self.cbRankTime.SetItems(lstRTime)
        self.cbRankTime.Select(12)
        
        lstAddr = PyDatabase.DBSelect(u"ID like '%%%%'", "RoadDays", ['AddrSn','Address'], 1)
        list_Addr = []
        for item in lstAddr:
            self.roadDict[item[0]] = item[1]
            list_Addr.append(item[1])
        self.cbAddr.SetItems(list_Addr)
        self.cbAddr.Select(0)
        
        self.InitTree()
        self.OnSelect(None)
        self.g1.Show(False)
    
    def OnDispPPNum(self):
        strDisp = self.list.GetItemCount()
        self.dispText.SetLabel(u"当前浏览人数：" + `strDisp`)
    
    def OnCkbInfo(self, event):
        self.OnSelect(None)
    
    def OnCkbModify(self, event):
        itemchecklst = [self.ckbRank, self.ckbRankTime, self.ckbAddr, self.ckbSex, self.ckbMarried]
        [item.SetValue(False) for item in itemchecklst]
        [item.Enable(not self.ckbModify.GetValue()) for item in itemchecklst]

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
        else:
            self.treeSel = None
        
    def OnActivate(self, event):
        item = event.GetItem()
        if item in self.treeDict.values():
            curUsn = self.treeDict.keys()[self.treeDict.values().index(item)]
            strResult = PyDatabase.DBSelect(curUsn, self.DTable, ['UnitSn'], 2)
            self.treeSel = item            
        else:
            strResult = PyDatabase.DBSelect("", self.DTable, ['Sn'], 0)
            self.treeSel = None
        self.FlashList(strResult)
        
        itemchecklst = [self.ckbRank, self.ckbAddr, self.ckbSex, self.ckbMarried,self.ckbModify]
        [item.SetValue(False) for item in itemchecklst]
        [item.Enable(not self.ckbModify.GetValue()) for item in itemchecklst[:-1]]
        self.OnDispPPNum()
    
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
            
    def DispColorList(self, list):
        for i in range(list.GetItemCount()):
            if i%4 == 0: list.SetItemBackgroundColour(i, (233, 233, 247))
            if i%4 == 1: list.SetItemBackgroundColour(i, (247, 247, 247))
            if i%4 == 2: list.SetItemBackgroundColour(i, (247, 233, 233))
            if i%4 == 3: list.SetItemBackgroundColour(i, (233, 247, 247))
            
    def OnCbChanged(self, event):
        itemchecklst = [self.ckbRank, self.ckbRankTime, self.ckbAddr, self.ckbSex, self.ckbMarried]
        flagNum = 0
        for item in itemchecklst:
            if not item.GetValue():
                flagNum += 1
        if flagNum == len(itemchecklst):
            return
        
        if self.ckbModify.GetValue():
            return
        self.OnSelect(None)
        
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
            index = self.list.InsertStringItem(10000, "A") 
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
    
    def EnableCombox(self):
        lstSelected = []
        item = self.list.GetFirstSelected()
        lstSelected.append(item)
        while self.list.GetNextSelected(item) != -1:
            item = self.list.GetNextSelected(item)
            lstSelected.append(item)
        
        strcoldisp = [u"军衔", u"军衔时间", u"籍贯", u"性别", u"婚否", u"单位"]
        lstPInfo = [self.cbRank, self.cbRankTime, self.cbAddr, self.cbSex, self.cbMarried, self.tree]
        
        lstColDisp = {}
        for item in strcoldisp:
            lstColDisp[item] = []

        for idItem in lstSelected:
            for item in strcoldisp:
                if self.list.GetItem(idItem, self.lstHead.index(item)).GetText() not in lstColDisp[item]:
                    lstColDisp[item].append(self.list.GetItem(idItem, self.lstHead.index(item)).GetText())
      
        for item in zip(strcoldisp, lstPInfo):
            if len(lstColDisp[item[0]]) != 1:
                item[1].Enable(False)
            else:
                item[1].Enable(True)
    
    def OnItemSelected(self, event):
        lstPInfo = [self.cbRank, self.cbRankTime, self.cbAddr, self.cbSex, self.cbMarried, self.tree, self.Text_Name, self.Text_Tel, self.Text_AddrAll]
        if self.list.GetSelectedItemCount() != 1:
            self.EnableCombox()
            [item.Enable(False) for item in lstPInfo[-3:]]
            return
        
        [item.Enable(True) for item in lstPInfo]
#        self.list.GetNextSelected()
#        print self.list.GetFirstSelected()
        
        self.listIndex = event.GetIndex()
        strSn = self.list.GetItem(self.listIndex, col=1).GetText()
            
        strResult = PyDatabase.DBSelect(strSn, self.DTable, ['Sn'], 2)[0][2:]
        self.Text_Name.SetValue(strResult[0])
        self.cbSex.SetValue(strResult[1])
        self.cbRank.SetValue(self.rankDict[strResult[2]])
        self.tree.SelectItem(self.treeDict[strResult[3]])
        self.cbAddr.SetValue(self.roadDict[strResult[4]])
        self.Text_AddrAll.SetValue(strResult[5])
        self.cbMarried.SetValue(strResult[6])
        self.Text_Tel.SetValue(strResult[7])
        self.cbRankTime.SetValue(strResult[-1][:4])
        
    def GetPersonInfo(self):
        lstinput = [None, '']
        lstinput.append(self.Text_Name.GetValue().strip())
        lstinput.append(self.cbSex.GetValue())
        lstinput.append(self.rankDict.keys()[self.rankDict.values().index(self.cbRank.GetValue())])
        lstinput.append(self.treeDict.keys()[self.treeDict.values().index(self.treeSel)])
        lstinput.append(self.roadDict.keys()[self.roadDict.values().index(self.cbAddr.GetValue())])
        lstinput.append(self.Text_AddrAll.GetValue().strip())
        lstinput.append(self.cbMarried.GetValue())
        lstinput.append(self.Text_Tel.GetValue().strip())
        lstinput.append(self.cbRankTime.GetValue() + "-12-1")
        return lstinput
        
    def ClearTxt(self):
        self.Text_Name.SetValue("")
        self.Text_Tel.SetValue("")
        self.Text_AddrAll.SetValue("") 
        self.DispColorList(self.list)
        
    def OnAdd(self, event):
        if not self.ckbModify.GetValue():
            wx.MessageBox(u"请进入编辑模式！", u"提示")
            return
        
        if self.cbRank.GetValue() == "":
            wx.MessageBox(u"请进入军衔维护界面添加相应军衔！", u"提示")
            return
        
        if self.cbAddr.GetValue() == "":
            wx.MessageBox(u"请进入籍贯维护添加相应籍贯！", u"提示")
            return
                
        if self.Text_Name.GetValue().strip() == "" or self.treeSel == None:
            wx.MessageBox(u"请点击单位并填写姓名！", u"提示")
            return

        strResult = PyDatabase.DBSelect("Sn like '%%%%'", self.DTable, ['Sn'], 1)
        if len(strResult) == 0:
            dlg = wx.TextEntryDialog(self.panel, u"请输入四位字母编号：", u"提示", 'BZDD')
            if dlg.ShowModal() == wx.ID_OK:
                personSn = dlg.GetValue() 
                if len(personSn) != 4:
                    wx.MessageBox(u"请输入四位的英文字符", u"提示")
                    dlg.Destroy()
                    return
                for iab in personSn:
                    if iab not in string.ascii_letters:
                        wx.MessageBox(u"请输入全部英文字符", u"提示")
                        dlg.Destroy()
                        return
            dlg.Destroy()
        else:
            personSn = strResult[0][0][:4]
            
        personSn = string.upper(personSn)
        
        listPerson = self.GetPersonInfo()
        lstIntSn = [int(isn[0][-4:]) for isn in strResult]
        newsn = -1
        for index in range(1, lstIntSn[-1]):
            if index not in lstIntSn:
                newsn = index
                break
        if newsn == -1:
            newsn = lstIntSn[-1]+1
            
        listPerson[1] = personSn + string.zfill(str(newsn),4)
        # Update Database
        PyDatabase.DBInsert(listPerson, self.DTable)
        # Update the list
        index = self.list.InsertStringItem(10000, "A")             
        self.list.SetStringItem(index, 0, `index+1`)
        self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体')) 
        [self.list.SetStringItem(index, i+1, listPerson[i+1]) for i in range(3)]
        self.list.SetStringItem(index, 4, self.rankDict[listPerson[4]])
        self.list.SetStringItem(index, 5, self.unitDict[listPerson[5]].split('|->')[-1])
        self.list.SetStringItem(index, 6, self.roadDict[listPerson[6]])
        [self.list.SetStringItem(index, i, listPerson[i]) for i in range(7,11)]

        self.ClearTxt()
        
    def OnModify(self, event):
        if not self.ckbModify.IsChecked():
            wx.MessageBox(u"请进入编辑模式！", u"提示")
            return
        
        if self.list.GetFirstSelected() == -1:
            wx.MessageBox(u"请选中表中一列数据！", u"提示")
            return
        
        lstSelected = []
        item = self.list.GetFirstSelected()
        lstSelected.append(item)
        while self.list.GetNextSelected(item) != -1:
            item = self.list.GetNextSelected(item)
            lstSelected.append(item)
        
        lstSelected.sort()
        for index in lstSelected:
            strSn = self.list.GetItem(index, col=1).GetText()
            oldPersonlst = PyDatabase.DBSelect(strSn, self.DTable, ['Sn'], 2)[0][1:]
            newPersonlst = []
            newPersonlst.extend(oldPersonlst)
            
            # modify the newPersonlst by the personinfo isenabled.
            if self.cbRank.IsEnabled():
                newPersonlst[3] = self.rankDict.keys()[self.rankDict.values().index(self.cbRank.GetValue())]
            if self.cbRankTime.IsEnabled():
                newPersonlst[9] = self.cbRankTime.GetValue() + "-12-1"
            if self.cbAddr.IsEnabled():
                newPersonlst[5] = self.roadDict.keys()[self.roadDict.values().index(self.cbAddr.GetValue())]
            if self.cbSex.IsEnabled():
                newPersonlst[2] = self.cbSex.GetValue()
            if self.cbMarried.IsEnabled():
                newPersonlst[7] = self.cbMarried.GetValue()
            if self.tree.IsEnabled():
                newPersonlst[4] = self.treeDict.keys()[self.treeDict.values().index(self.treeSel)]
            
            #self.Text_Name, self.Text_Tel, self.Text_AddrAll
            if self.Text_Name.IsEnabled():                
                newPersonlst[1] = self.Text_Name.GetValue().strip()
                newPersonlst[6] = self.Text_AddrAll.GetValue().strip()
                newPersonlst[8] = self.Text_Tel.GetValue().strip()
                        
            # Update Database
            PyDatabase.DBUpdate(oldPersonlst, newPersonlst, self.DTable)        
            # Update list
            [self.list.SetStringItem(index, i+1, newPersonlst[i]) for i in range(3)]
            self.list.SetStringItem(index, 4, self.rankDict[newPersonlst[3]])
            self.list.SetStringItem(index, 5, self.unitDict[newPersonlst[4]].split('|->')[-1])
            self.list.SetStringItem(index, 6, self.roadDict[newPersonlst[5]])
            [self.list.SetStringItem(index, i+1, newPersonlst[i]) for i in range(6,10)]

        self.ClearTxt()
    
    def OnDelete(self, event):
        if self.list.GetFirstSelected() == -1:
            wx.MessageBox(u"请选中表中一列数据！", u"提示")
            return
        
        if wx.CANCEL == wx.MessageBox(u"删除本记录将删除这些人对应的休假记录！是否要继续？", u"提示", wx.OK| wx.CANCEL):
            return
        
        lstSelected = []
        item = self.list.GetFirstSelected()
        lstSelected.append(item)
        while self.list.GetNextSelected(item) != -1:
            item = self.list.GetNextSelected(item)
            lstSelected.append(item)
        
        lstsn = []
        lstSelected.sort(reverse=True)
        for index in lstSelected:
            strSn = self.list.GetItem(index, col=1).GetText()
            oldPersonlst = PyDatabase.DBSelect(strSn, self.DTable, ['Sn'], 2)[0][1:]
            # Update Database
            PyDatabase.DBDelete(oldPersonlst, self.DTable)
            lstsn.append(strSn)        
            # Update list        
            self.list.DeleteItem(index)
        
        strHDR = PyDatabase.DBSelect("", 'PPHdRecord', ['Sn'], 0)
        for item in strHDR:
            if item[1][:g_UnitSnNum] in lstsn:
                PyDatabase.DBDelete(item[1:], "PPHdRecord")
        
        [self.list.SetStringItem(i, 0, `i+1`) for i in range(self.list.GetItemCount())]        
        self.ClearTxt() 

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
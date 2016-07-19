#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
#import wx.lib.buttons as buttons
import PyDatabase
import images
import string
import datetime, time
import operateWord
import MyValidator
import types
import pickle
from PhrResource import ALPHA_ONLY, DIGIT_ONLY, strVersion,g_UnitSnNum, MyListCtrl
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
        
        HoliPermit(panel)
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()
        
class HoliPermit(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)
        
        titleText = wx.StaticText(panel, -1, u"人员批假信息")
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        self.dispText = wx.StaticText(panel)
        topsizer0 = wx.BoxSizer(wx.HORIZONTAL)
        topsizer0.Add(titleText, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL)
        
        self.ckbRank = wx.CheckBox(panel, -1, u"军衔")
        self.cbRank = wx.ComboBox(panel)
        self.ckbAddr = wx.CheckBox(panel, -1, u"籍贯")
        self.cbAddr = wx.ComboBox(panel)
        self.ckbSex = wx.CheckBox(panel, -1, u"性别")
        self.cbSex = wx.ComboBox(panel, -1, u"男", choices=[u"男", u"女"])
        self.ckbMarried = wx.CheckBox(panel, -1, u"婚否")
        self.cbMarried = wx.ComboBox(panel, -1, u"未婚", choices=[u"未婚", u"已婚"])
        lblname = wx.StaticText(panel, -1, u"姓名")

        self.Text_Select = wx.TextCtrl(panel)
        btn_Select = wx.Button(panel, -1, u"查询", size=(60,-1))
        btn_Select.SetDefault()
        
        topsizer = wx.BoxSizer(wx.HORIZONTAL)        
        itemSelect = [self.ckbRank, self.cbRank, self.ckbAddr, self.cbAddr,self.ckbSex, self.cbSex, self.ckbMarried, self.cbMarried]
        for item in itemSelect:
            topsizer.Add(item, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
            if item.GetClassName() == u"wxComboBox":
                topsizer.Add((5, -1))
        topsizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        topsizer.Add(lblname, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        topsizer.Add(self.Text_Select, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        topsizer.Add((10, -1), 0, wx.ALL|wx.EXPAND)
        topsizer.Add(btn_Select, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)   

        lbl5 = wx.StaticText(panel, -1, u"请假事由：")
        self.Text_Reason = wx.TextCtrl(panel, -1, `time.localtime()[0]` + u"年正常假")
        self.ckbHK = wx.CheckBox(panel, -1, u"正常假：")
        self.ckbHK.SetValue(True)
        lbl6 = wx.StaticText(panel, -1, u"基本天数：")
        self.Text_BasicDays = wx.TextCtrl(panel, -1, u"", size=(30,-1),validator = MyValidator.MyValidator(DIGIT_ONLY))
        self.ckbAddr2 = wx.CheckBox(panel, -1, u"路途：")
        self.ckbAddr2.SetValue(True)
        self.cbAddr2 = wx.ComboBox(panel)
        self.cbAddrDays = wx.ComboBox(panel)
        lbl7 = wx.StaticText(panel, -1, u"已婚假：")
        self.Text_MarriedDays = wx.TextCtrl(panel, -1, u"10", size=(30,-1),validator = MyValidator.MyValidator(DIGIT_ONLY))
        lbl9 = wx.StaticText(panel, -1, u"额外假：")
        self.Text_ExtraDays = wx.TextCtrl(panel, -1, u"0")
        self.Text_ExtraDays.SetEditable(False)
        lbl10 = wx.StaticText(panel, -1, u"离队日期：")
        self.dateStart = wx.DatePickerCtrl(panel, style = wx.DP_DROPDOWN | wx.DP_ALLOWNONE|wx.DP_SHOWCENTURY, size=(120,-1))
            
        lbl11 = wx.StaticText(panel, -1, u"假期详情：")
        self.Text_Demo = wx.TextCtrl(panel)

        lbl12 = wx.StaticText(panel, -1, u"经 手 人：")
        self.Text_ByMan = wx.TextCtrl(panel)
        lbl13 = wx.StaticText(panel, -1, u"批准领导：")
        self.Text_ByLeader = wx.TextCtrl(panel)
        
        self.list3= wx.ListCtrl(panel, -1, size=(100, 80), style=wx.LC_REPORT| wx.LC_VRULES | wx.LC_HRULES|wx.LC_SINGLE_SEL)
        self.list3.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        lstExtrahd = [u"晚婚假",u"补假",u"晚育假",u"扣除事假",u"其他假"]
        self.cbExtraHd = wx.ComboBox(panel, -1, lstExtrahd[0], choices=lstExtrahd)
        self.cbExtraHd.SetEditable(False)
        self.Text_ExtraItem = wx.TextCtrl(panel, -1, u"10", validator = MyValidator.MyValidator(DIGIT_ONLY))
        self.btn_AddExtra = wx.Button(panel, -1, u"添加", size=(40,-1))
        
        infosizer1 = wx.BoxSizer(wx.HORIZONTAL)
        infosizer1.Add(lbl5, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 0)
        infosizer1.Add(self.Text_Reason, 1, wx.ALL| wx.EXPAND, 0)
        
        iteminfo = [self.ckbHK, lbl6, self.Text_BasicDays, (-1,-1), lbl7, self.Text_MarriedDays, (-1,-1), self.ckbAddr2, self.cbAddr2, self.cbAddrDays, (-1,-1)]
        infosizer2 = wx.BoxSizer(wx.HORIZONTAL)
        for item in iteminfo:
            if type(item) == types.TupleType:
                infosizer2.Add(item, 1, wx.ALIGN_RIGHT| wx.EXPAND)
            else:
                infosizer2.Add(item, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)

        iteminfo = [lbl12, self.Text_ByMan, (-1,-1), lbl13, self.Text_ByLeader, (-1,-1), lbl10, self.dateStart, (-1,-1)]
        infosizer3 = wx.BoxSizer(wx.HORIZONTAL)
        for item in iteminfo:
            if type(item) == types.TupleType:
                infosizer3.Add(item, 1, wx.ALIGN_RIGHT| wx.EXPAND)
            else:
                infosizer3.Add(item, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL, 0)
                
        infosizer4 = wx.BoxSizer(wx.HORIZONTAL)
        infosizer4.Add(lbl11, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL, 0)
        infosizer4.Add(self.Text_Demo, 1, wx.ALL| wx.EXPAND, 0)
        
        infoLeft = wx.BoxSizer( wx.VERTICAL)
        [infoLeft.AddSizer(item, 1, wx.ALL| wx.EXPAND, 5) for item in [infosizer1,infosizer2,infosizer3,infosizer4]]

        infosizer5 = wx.BoxSizer( wx.HORIZONTAL)
        infosizer5.Add(lbl9, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)
        infosizer5.Add(self.Text_ExtraDays, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)
        infosizer6 = wx.BoxSizer( wx.HORIZONTAL)
        infosizer6.Add(self.cbExtraHd, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)
        infosizer6.Add(self.Text_ExtraItem, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)
        infosizer6.Add(self.btn_AddExtra, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 10)
        infoRight = wx.BoxSizer( wx.VERTICAL)
        infoRight.AddSizer(infosizer5, 0, wx.ALL, 5)
        infoRight.Add(self.list3, 0, wx.ALL|wx.EXPAND|wx.ALIGN_RIGHT)
        infoRight.AddSizer(infosizer6, 0, wx.ALL, 5)
        
        midsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        midsizer2.AddSizer(infoLeft, 1, wx.ALL| wx.EXPAND, 5)
        midsizer2.Add(wx.StaticLine(panel, style= wx.LI_VERTICAL), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        midsizer2.AddSizer(infoRight, 0, wx.ALL, 5)

        btn_Add = wx.Button(panel, -1, u"批假") 
        btn_Delete = wx.Button(panel, -1, u"删除") 
        btn_Modify = wx.Button(panel, -1, u"修改") 
        btn_Help = wx.Button(panel, -1, u"帮助")
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [btn_Add, btn_Delete, btn_Modify, btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1) 
        
        lbltoday = wx.StaticText(panel, -1, u" 休假人员信息：");
        self.list2 = wx.ListCtrl(panel, -1, size=(-1, 80), style=wx.LC_REPORT| wx.LC_VRULES | wx.LC_HRULES|wx.LC_SINGLE_SEL)
        self.list2.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        
        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT| wx.LC_VRULES | wx.LC_HRULES|wx.LC_SINGLE_SEL)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        listsizer = wx.BoxSizer(wx.VERTICAL)
        listsizer.AddSizer(topsizer,0, wx.ALL| wx.EXPAND, 0)
        listsizer.Add(self.list, 1, wx.ALL|wx.EXPAND, 0)
        
        lbltree = wx.StaticText(panel, -1, u"单位树")
        self.tree = wx.TreeCtrl(panel, size=(160, -1))
        self.root = self.tree.AddRoot(u"单位")
        
        treesizer = wx.BoxSizer(wx.VERTICAL)
        treesizer.Add(lbltree, 0, wx.ALL, 0)
        treesizer.Add(self.tree, 1, wx.ALL|wx.EXPAND, 0)

        midsizer = wx.BoxSizer(wx.HORIZONTAL)
        midsizer.AddSizer(treesizer, 0, wx.ALL| wx.EXPAND, 5)
        midsizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer0, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add(self.dispText, 0, wx.ALL|wx.BOTTOM| wx.ALIGN_LEFT, 0)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(midsizer, 1, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 2)
        mainSizer.Add(lbltoday, 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(self.list2, 0, wx.ALL|wx.EXPAND|wx.ALIGN_RIGHT, 5)
        mainSizer.AddSizer(midsizer2, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      

        panel.popMenuList0 = wx.Menu()
        pmList_11 = panel.popMenuList0.Append(1309, u"显示休假详情")
        self.list.Bind(wx.EVT_MENU, self.OnPopItemSelected1, pmList_11)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop1)
        
        panel.popMenuList2 = wx.Menu()
        pmList_21 = panel.popMenuList2.Append(1301, u"删除")
        panel.popMenuList2.AppendSeparator()
        pmList_22 = panel.popMenuList2.Append(1302, u"导出批假通知单")
        self.list2.Bind(wx.EVT_MENU, self.OnPopItemSelected2, pmList_21)
        self.list2.Bind(wx.EVT_MENU, self.OnPopItemSelected2, pmList_22)
        self.list2.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop2)
        
        panel.popMenuList3 = wx.Menu()
        pmList_31 = panel.popMenuList3.Append(1310, u"删除")
        self.list3.Bind(wx.EVT_MENU, self.OnPopItemSelected3, pmList_31)
        self.list3.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop3)
        
        self.treeDict = {}
        self.treeSel = None
        self.lstPHR = []
        self.phrcount = 0

        self.unitDict = {}
        self.rankDict = {}
        self.roadDict = {}
        self.roadDaysDict = {}
        self.rankDaysDict = {}
        self.holiDays = {}
        self.DTable = 'PersonInfo'
        self.DTable2 = 'PPHdRecord'
        
        lstHead = [u"序号", u"编号", u"姓名",u"性别",u"军衔",u"单位",u"籍贯", u"详细地址",u"婚否",u"电话"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
        self.list.SetColumnWidth(0, 50)
        self.list.SetColumnWidth(3, 60)
        
        lstHead = [u"序号", u"编号", u"姓名",u"请假事由",u"天数", u"离队日期",u"归队日期",u"",u"经手",u"批假", u"", u"假期详情", u"类别", u"备注"]
        [self.list2.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]
        [self.list2.SetColumnWidth(i, 0) for i in [7,10]]
        
        lstHead = [u"序号", u"名称", u"天数"]
        [self.list3.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]

        infoPerson = [self.cbSex, self.cbMarried, self.cbRank, self.cbAddr]
        [item.SetEditable(False) for item in infoPerson]

        self.InitData()
        
        panel.Show()
        itemchecklst = [self.ckbRank, self.ckbAddr, self.ckbSex, self.ckbMarried]
        [panel.Bind(wx.EVT_CHECKBOX, self.OnCkbInfo, item) for item in itemchecklst]
        panel.Bind(wx.EVT_TEXT, self.OnBasicText, self.Text_Reason)
        panel.Bind(wx.EVT_COMBOBOX, self.OnBasicText, self.cbAddrDays)
        panel.Bind(wx.EVT_TEXT, self.OnBasicText, self.Text_ExtraDays)
        panel.Bind(wx.EVT_TEXT, self.OnBasicText, self.Text_BasicDays)
        panel.Bind(wx.EVT_TEXT, self.OnBasicText, self.Text_MarriedDays)
        panel.Bind(wx.EVT_CHECKBOX, self.OnBasicText, self.ckbAddr2)
        panel.Bind(wx.EVT_CHECKBOX, self.OnHoliKind, self.ckbHK)
        panel.Bind(wx.EVT_COMBOBOX, self.OnCbAddr, self.cbAddr2)
        panel.Bind(wx.EVT_DATE_CHANGED, self.OnDateChanged, self.dateStart)
        panel.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelChanged, self.tree)
        panel.Bind(wx.EVT_TREE_ITEM_ACTIVATED, self.OnActivate, self.tree)
        [panel.Bind(wx.EVT_COMBOBOX, self.OnCbChanged, item) for item in [self.cbRank, self.cbAddr, self.cbSex, self.cbMarried]]
        panel.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.list)
        panel.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected2, self.list2)
        panel.Bind(wx.EVT_BUTTON, self.OnAdd, btn_Add)       
        panel.Bind(wx.EVT_BUTTON, self.OnDelete, btn_Delete)       
        panel.Bind(wx.EVT_BUTTON, self.OnModify, btn_Modify)       
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_BUTTON, self.OnSelect, btn_Select)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
        panel.Bind(wx.EVT_BUTTON, self.OnAddExtra, self.btn_AddExtra)
    
    def OnDispPPNum(self, strunit):
        strResult = PyDatabase.DBSelect("Sn like '%%%%'", self.DTable2, ['Sn', 'DateInfactEnd'], 1)
        
        snLst = [self.list.GetItem(i, col=1).GetText() for i in range(self.list.GetItemCount())]

        strtext = " " + strunit + u"在位人数：" + `len(snLst)` + u"，当前正在休假人数：" + `self.phrcount`
        self.dispText.SetLabel(strtext)
    
    def InitData(self):
        strPHR = PyDatabase.DBSelect("Sn like '%%%%' and DateInfactEnd = ''", 'PPHdRecord', ['Sn'], 1)
        [self.lstPHR.append(iPHR[0][:g_UnitSnNum]) for iPHR in strPHR]
        
        lstHoli = PyDatabase.DBSelect(u"ID like '%%%%'", "HoliDays", ['HolidayTime', 'HolidayName', 'Days'], 1)
        for item in lstHoli:
            self.holiDays[item[0]] = [item[1],item[2]]

        lstRank = PyDatabase.DBSelect(u"ID like '%%%%'", "RankDays", ['RankSn', 'LevelRank', 'Days'], 1)
        list_Rank = []
        for item in lstRank:
            self.rankDict[item[0]] = item[1]
            self.rankDaysDict[item[0]] = item[2]
            list_Rank.append(item[1])
        self.cbRank.SetItems(list_Rank)
        self.cbRank.Select(0)
        
        lstAddr = PyDatabase.DBSelect(u"ID like '%%%%'", "RoadDays", ['AddrSn','Address', 'MinDays', 'MaxDays'], 1)
        list_Addr = []
        for item in lstAddr:
            self.roadDict[item[0]] = item[1]
            list_Addr.append(item[1])
            self.roadDaysDict[item[1]] = [item[2], item[3]]
        self.cbAddr.SetItems(list_Addr)
        self.cbAddr.Select(0)
        self.cbAddr2.SetItems(list_Addr)
        self.cbAddr2.Select(0)
        
        self.InitTree()
        self.OnSelect(None)
        self.OnCbAddr(None)
        
        try:
            pkl_file = open('ByData.pkl', 'rb')
        except:
            data = {'ByMan': u"",\
            'ByLeader': u"", \
            'ByUnit': u""}        
            output = open('ByData.pkl', 'wb')
            pickle.dump(data, output)
            output.close()
            self.Text_ByMan.SetValue(data['ByMan'])
            self.Text_ByLeader.SetValue(data['ByLeader'])
            return
        
        data = pickle.load(pkl_file)
        pkl_file.close()
        self.Text_ByMan.SetValue(data['ByMan'])
        self.Text_ByLeader.SetValue(data['ByLeader'])
                
    def OnCkbInfo(self, event):
        self.OnSelect(None)
        
    def OnOutWord(self):
        listIndex2 = self.list2.GetFirstSelected()
        lstword = []
        lstword.append(self.list2.GetItem(listIndex2, col=1).GetText())
        lstword.append(self.list2.GetItem(listIndex2, col=2).GetText())
        lstword.append(self.list.GetItem(0, col=5).GetText())
        lstword.append(self.list2.GetItem(listIndex2, col=11).GetText())
        lstword.append(self.list2.GetItem(listIndex2, col=5).GetText())
        lstword.append(self.list2.GetItem(listIndex2, col=6).GetText())
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
        
    def OnBasicText(self, event):
        if not self.dateStart.GetValue().IsOk():
            self.Text_Demo.SetValue("")
            return
        self.cbAddr2.Enable(self.ckbAddr2.GetValue())
        self.cbAddrDays.Enable(self.ckbAddr2.GetValue())
        if self.list.GetFirstSelected() == -1:
            return
        basicdays = int(self.Text_BasicDays.GetValue()=="" and '0' or self.Text_BasicDays.GetValue())
        if not self.ckbHK.GetValue():
            strDisp = self.Text_Reason.GetValue().strip() + `basicdays` + u"天，"
            strDisp += u"共计" + str(basicdays) + u"天。" 
        else: 
            basicdays += int(self.Text_MarriedDays.GetValue()=="" and '0' or self.Text_MarriedDays.GetValue())
            strDisp = self.Text_Reason.GetValue().strip() + `basicdays` + u"天，"   
            lstExtraHd = []
            for index in range(self.list3.GetItemCount()):
                lstExtraHd.append([self.list3.GetItem(index, col=1).GetText(), self.list3.GetItem(index, col=2).GetText()])

            roaddays = int(self.cbAddrDays.GetValue())
            extradays = int(self.Text_ExtraDays.GetValue()=='' and '0' or self.Text_ExtraDays.GetValue())

            for item in lstExtraHd:
                strDisp += item[0] + " " + item[1] + u" 天，"
            
            if self.ckbAddr2.GetValue():
                strDisp += u"路途（" + self.cbAddr2.GetValue() + u"）" + `roaddays` + u"天，"
                strDisp += u"共计" + str(basicdays + roaddays + extradays) + u"天。" 
            else:
                strDisp += u"共计" + str(basicdays + extradays) + u"天。" 
        
        self.Text_Demo.SetValue(strDisp)

    def OnDelete3(self):
        listIndex3 = self.list3.GetFirstSelected()
        if listIndex3 == -1:
            return
        days = int(self.Text_ExtraDays.GetValue())
        days -= int(self.list3.GetItem(listIndex3, col=2).GetText())
        self.list3.DeleteItem(listIndex3)
        [self.list2.SetStringItem(i, 0, `i+1`) for i in range(self.list2.GetItemCount())]
        self.Text_ExtraDays.SetValue(`days`)
        
    def OnAddExtra(self, event):
        if self.cbExtraHd.GetValue().strip() == "" or self.Text_ExtraItem.GetValue().strip() == "":
            wx.MessageBox(u"请填写额外假期名称及天数", u"提示")
            return
        lsthdname = [ self.list3.GetItem(i, col=1).GetText() for i in range( self.list3.GetItemCount())]
        if self.cbExtraHd.GetValue().strip() in lsthdname:
            wx.MessageBox(u"该假已添加！如需修改，请删除后重新添加！", u"提示")
            return
        
        index = self.list3.InsertStringItem(100, "A")
        self.list3.SetStringItem(index, 0, `index+1`)
        self.list3.SetStringItem(index, 1, self.cbExtraHd.GetValue().strip())
        if self.cbExtraHd.GetValue() == u"扣除事假":
            tmpday = '-' + self.Text_ExtraItem.GetValue().strip()
        else:
            tmpday = self.Text_ExtraItem.GetValue().strip()
        self.list3.SetStringItem(index, 2, tmpday)
        
        days = int(self.Text_ExtraDays.GetValue())
        days += int(tmpday)
        self.Text_ExtraDays.SetValue(`days`)
        
    def OnHoliKind(self, event):
        infohd = [self.ckbAddr2, self.Text_MarriedDays, self.Text_ExtraDays, self.btn_AddExtra]
        [item.Enable(self.ckbHK.GetValue()) for item in infohd]
        self.ckbAddr2.SetValue(self.ckbHK.GetValue())
        
        if self.ckbHK.GetValue():
            self.Text_Reason.SetValue(`time.localtime()[0]` + u"年正常假")
        else:
            self.Text_Reason.SetValue(`time.localtime()[0]` + u"年事假")
            self.Text_BasicDays.SetValue('10')
        
        self.OnDateChanged(None)
        
    def OnCbAddr(self, event):
        self.cbAddrDays.Clear()
        if len(self.roadDaysDict) == 0:
            return
        roaddays = self.roadDaysDict[self.cbAddr2.GetValue()]
        lstroaddays = range(int(roaddays[0]), int(roaddays[1])+1)
        for i in lstroaddays:
            self.cbAddrDays.Append(str(i))
        self.cbAddrDays.Select(len(lstroaddays)-1)
    
    def OnDateChanged(self, event):
        if not self.dateStart.GetValue().IsOk():
            self.list3.DeleteAllItems()
            self.Text_ExtraDays.SetValue('0')
            return
        
        if not self.ckbHK.GetValue():
            self.list3.DeleteAllItems()
            self.Text_ExtraDays.SetValue('0')
            return
        
        days = 0
        for item in [self.Text_BasicDays, self.Text_MarriedDays, self.cbAddrDays]:
            if item.GetValue() != "":
                days += int(item.GetValue())

#        days = sum([int(item.GetValue()) for item in [self.Text_BasicDays, self.Text_MarriedDays, self.cbAddrDays]])
        
        daysExtra = 0
        self.list3.DeleteAllItems()
        self.Text_ExtraDays.SetValue(`daysExtra`)
                    
        startdayStr = self.dateStart.GetLabel().split('-')
        startday = datetime.date(int(startdayStr[0]), int(startdayStr[1]), int(startdayStr[2]))
        endday = startday + datetime.timedelta(days-1)  
        
        for item in self.holiDays:
            hddate = datetime.date(int(item.split('-')[0]), int(item.split('-')[1]), int(item.split('-')[2]))
            if hddate >=startday and hddate <= endday:
                index = self.list3.InsertStringItem(100, "A")
                self.list3.SetStringItem(index, 0, `index+1`)
                self.list3.SetStringItem(index, 1, self.holiDays[item][0])
                self.list3.SetStringItem(index, 2, self.holiDays[item][1])
                daysExtra += int(self.holiDays[item][1])
            
        [self.list3.SetStringItem(i, 0, `i+1`) for i in range(self.list3.GetItemCount())]
        self.DispColorList( self.list3)
        
        self.Text_ExtraDays.SetValue(`daysExtra`)
    
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
        curUsn = ""
        if item in self.treeDict.values():
            curUsn = self.treeDict.keys()[self.treeDict.values().index(item)]
            strResult = PyDatabase.DBSelect(curUsn, self.DTable, ['UnitSn'], 2)
            self.treeSel = item
        else:
            strResult = PyDatabase.DBSelect("", self.DTable, ['Sn'], 0)
            self.treeSel = None
        self.FlashList(strResult)
        
        self.OnDispPPNum(self.GetItemText(item))
        
        itemchecklst = [self.ckbRank, self.ckbAddr, self.ckbSex, self.ckbMarried]
        [item.SetValue(False) for item in itemchecklst]
        
    def DispPersonHR(self):
        strPPInfo = u"　　　编号\t\t休假时间\t天数\t类别\t剩余天数\n"
        strPPInfo += "_"*65 + "\n\n"
        
        listIndex = self.list.GetFirstSelected()
        strSn = self.list.GetItem(listIndex, col=1).GetText()
        strPHR = PyDatabase.DBSelect("Sn like '%" + strSn + "%'", 'PPHdRecord', ['Sn', 'DateStart', 'Days', 'HoliKind', 'DaysDelta'], 1)
        
        if len(strPHR) == 0:
            wx.MessageBox(u"该同志暂时没有休假记录!", u"提示")
        else:
            for item in strPHR:
                strPPInfo += string.join(item, '\t') + '\n'
            wx.MessageBox(strPPInfo, u"提示")
       
    def OnShowPop1(self, event):
        if self.list.GetItemCount() != 0:
            self.list.PopupMenu(self.panel.popMenuList0)
        
    def OnPopItemSelected1(self, event):
        try:
            item = self.panel.popMenuList0.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"显示休假详情":
                self.DispPersonHR()
        except:
            pass
        
    def OnShowPop2(self, event):
        if self.list2.GetItemCount() != 0:
            self.list2.PopupMenu(self.panel.popMenuList2)
         
    def OnPopItemSelected2(self, event):
        try:
            item = self.panel.popMenuList2.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"删除":
                self.OnDelete(None)
            if text == u"导出批假通知单":
                self.OnOutWord()
        except:
            pass
    
    def OnShowPop3(self, event):
        if self.list3.GetItemCount() != 0:
            self.list3.PopupMenu(self.panel.popMenuList3)
                
    def OnPopItemSelected3(self, event):
        try:
            item = self.panel.popMenuList3.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"删除":
                self.OnDelete3()
        except:
            pass
            
    def DispColorList(self, list):
        for i in range(list.GetItemCount()):
            if i%4 == 0: list.SetItemBackgroundColour(i, (233, 233, 247))
            if i%4 == 1: list.SetItemBackgroundColour(i, (247, 247, 247))
            if i%4 == 2: list.SetItemBackgroundColour(i, (247, 233, 233))
            if i%4 == 3: list.SetItemBackgroundColour(i, (233, 247, 247))
            
    def OnCbChanged(self, event):
        itemchecklst = [self.ckbRank, self.ckbAddr, self.ckbSex, self.ckbMarried]
        flagNum = 0
        for item in itemchecklst:
            if not item.GetValue():
                flagNum += 1
        if flagNum == len(itemchecklst):
            return
        
        self.OnSelect(None)
        
    def OnSelect(self, event):
        if len(self.rankDict) == 0:
            return
        strName = self.Text_Select.GetValue()
        # fuzzy select
        lstsql = []
        if self.ckbRank.GetValue():
            lstsql.append("RankSn = '" + self.rankDict.keys()[self.rankDict.values().index(self.cbRank.GetValue())] + "'")
        if self.ckbAddr.GetValue():
            lstsql.append("AddrSn = '" + self.roadDict.keys()[self.roadDict.values().index(self.cbAddr.GetValue())] + "'")
        if self.ckbMarried.GetValue():
            lstsql.append("Married = '" + self.cbMarried.GetValue() + "'")
        if self.ckbSex.GetValue():
            lstsql.append("Sex = '" + self.cbSex.GetValue() + "'")
        
        curUsn = ""
        if self.treeSel is not None:
            curUsn = self.treeDict.keys()[self.treeDict.values().index(self.treeSel)]
            lstsql.append("UnitSn = '" + curUsn + "'")

        strsql = ""
        for item in lstsql:
            strsql += item + " and "
        strsql += "Name"
        
        strResult = PyDatabase.DBSelect(strName, self.DTable, [strsql], 0)
        self.FlashList(strResult)
        self.Text_Reason.SetValue("")
        self.Text_Demo.SetValue("")

        if len(lstsql) == 0 and self.treeSel is None:
            self.OnDispPPNum(self.GetItemText(self.root))
        elif lstsql[0] == "UnitSn = '" + curUsn + "'":
            self.OnDispPPNum(self.GetItemText(self.treeSel))
        else:
            self.OnDispPPNum(u"当前浏览")
    
    def FlashList(self, strResult):
        self.list.DeleteAllItems()
        self.phrcount = 0       # numbers of the PHR

        for row in strResult:
            if row[1] not in self.lstPHR:
                index = self.list.InsertStringItem(10000, "A") 
                self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体')) 
                self.list.SetStringItem(index, 0, `index+1`)
                [self.list.SetStringItem(index, i+1, row[i+1]) for i in range(3)]
                self.list.SetStringItem(index, 4, self.rankDict[row[4]])
                self.list.SetStringItem(index, 5, self.unitDict[row[5]].split('|->')[-1])
                self.list.SetStringItem(index, 6, self.roadDict[row[6]])
                [self.list.SetStringItem(index, i, row[i]) for i in range(7,10)]
            else:
                self.phrcount += 1
        if len(strResult) != 0:
            [self.list.SetColumnWidth(i, wx.LIST_AUTOSIZE) for i in [5,7,9]]
        self.DispColorList(self.list)
    
    def OnItemSelected2(self, event):
        listIndex2 = event.GetIndex()
        strName = self.list2.GetItem(listIndex2, col=2).GetText()
        self.Text_Reason.SetValue(self.list2.GetItem(listIndex2, col=3).GetText())
        self.Text_ByMan.SetValue(self.list2.GetItem(listIndex2, col=8).GetText())
        self.Text_ByLeader.SetValue(self.list2.GetItem(listIndex2, col=9).GetText())
        self.Text_Demo.SetValue(self.list2.GetItem(listIndex2, col=11).GetText())
        if self.list2.GetItem(listIndex2, col=12).GetText() == u"正常假":
            self.ckbHK.SetValue(True)
        else:
            self.ckbHK.SetValue(False)
        path = os.getcwd() +  "\\tmp.png"
        hc.Code128Encoder(str(self.list2.GetItem(listIndex2, col=1).GetText())).save(path,1) 
        
    def OnItemSelected(self, event):
        listIndex = event.GetIndex()
        strSn = self.list.GetItem(listIndex, col=1).GetText()
        self.OnHoliKind(None)
                
        strResult = PyDatabase.DBSelect(strSn, self.DTable, ['Sn'], 2)[0][2:]
        
        self.cbSex.SetValue(strResult[1])
        self.cbRank.SetValue(self.rankDict[strResult[2]])
        self.Text_BasicDays.SetValue(self.rankDaysDict[strResult[2]])
        self.tree.SelectItem(self.treeDict[strResult[3]])
        self.cbAddr.SetValue(self.roadDict[strResult[4]])
        self.cbAddr2.SetValue(self.roadDict[strResult[4]])
        self.OnCbAddr(None)
        if self.list.GetItem(listIndex, col=8).GetText() == u"已婚":
            self.Text_MarriedDays.SetValue('10')
        else:
            self.Text_MarriedDays.SetValue('0')
        self.cbMarried.SetValue(strResult[6])
        
        self.list3.DeleteAllItems()
        self.Text_ExtraDays.SetValue('0')
        self.OnDateChanged(None)
        
    def GetPersonInfo(self):
        lstinput = [self.Text_Reason.GetValue().strip()]
        if self.ckbHK.GetValue():
            days = int(self.Text_BasicDays.GetValue()) + \
                int(self.Text_MarriedDays.GetValue())+ int(self.Text_ExtraDays.GetValue())
            if self.ckbAddr2.GetValue():
                days += int(self.cbAddrDays.GetValue())
        else:
            days = int(self.Text_BasicDays.GetValue())
        lstinput.append(`days`)
        startday = self.dateStart.GetLabel()
        lstinput.append(str(datetime.date(int(startday.split('-')[0]), int(startday.split('-')[1]), int(startday.split('-')[2]))))
        endday = datetime.date(int(startday.split('-')[0]), int(startday.split('-')[1]), int(startday.split('-')[2])) + datetime.timedelta(days-1)  
        lstinput.append(str(endday))
        lstinput.append("")
        lstinput.append( self.Text_ByMan.GetValue() )
        lstinput.append( self.Text_ByLeader.GetValue() )
        lstinput.append("")
        lstinput.append(self.Text_Demo.GetValue())
        if self.ckbHK.GetValue():
            lstinput.append(u"正常假")
        else:
            lstinput.append(u"事假")
        lstinput.append("")
        return lstinput
    
    def OnAdd(self, event):
        listIndex = self.list.GetFirstSelected()
        if listIndex == -1 or self.Text_BasicDays.GetValue().strip() == "":
            wx.MessageBox(u"请点击选中要批假的人员，并填写基本假期天数。", u"提示")
            return
        
        if not self.dateStart.GetValue().IsOk():
            wx.MessageBox(u"请填写批假时间。", u"提示")
            return
        
        if self.Text_Reason.GetValue().strip() == "" or \
            self.Text_ByMan.GetValue().strip() == "" or \
            self.Text_ByLeader.GetValue().strip() == "" or \
            self.Text_Demo.GetValue().strip() == "":
            wx.MessageBox(u"请将请假事由、经手人和批准领导均填写完整。", u"提示")
            return
        
        listPerson = [None]
        startday = self.dateStart.GetLabel()
        phrsn = self.list.GetItem(listIndex, col=1).GetText() + str(datetime.date(int(startday.split('-')[0]), int(startday.split('-')[1]), int(startday.split('-')[2])))
        listPerson.append(phrsn.replace('-',''))
        listPerson.extend(self.GetPersonInfo())
        # Update Database
        PyDatabase.DBInsert(listPerson, self.DTable2)
        # Update the list
        listPerson.insert(2, self.list.GetItem(listIndex, col=2).GetText())
        index = self.list2.InsertStringItem(10000, "A")             
        self.list2.SetStringItem(index, 0, `index+1`)
        [self.list2.SetStringItem(index, i+1, listPerson[i+1]) for i in range(len(listPerson)-1)]        
        [self.list2.SetColumnWidth(i, wx.LIST_AUTOSIZE) for i in [1,3,5,6,11]]
        
        self.list.DeleteItem(listIndex)
        
        # Update the self.lstPHR
        self.lstPHR.append(phrsn[:g_UnitSnNum])
        
        self.Text_Reason.SetValue("")
        self.DispColorList(self.list2)
        self.Text_Demo.SetValue("")
        
        pkl_file = open('ByData.pkl', 'rb')
        data = pickle.load(pkl_file)
        pkl_file.close()
        
        data['ByMan'] = self.Text_ByMan.GetValue().strip()
        data['ByLeader'] = self.Text_ByLeader.GetValue().strip()
        output = open('ByData.pkl', 'wb')
        pickle.dump(data, output)
        output.close()
        
    def OnModify(self, event):
        listIndex2 = self.list2.GetFirstSelected()
        if listIndex2 == -1:
            wx.MessageBox(u"请点击选中已经被批假的人员。", u"提示")
            return
        if self.Text_BasicDays.GetValue().strip() == "":
            wx.MessageBox(u"请填写基本假期。", u"提示")
            return
        
        oldPersonlst = [self.list2.GetItem(listIndex2, col=i+1).GetText() for i in range(self.list2.GetColumnCount()-1)]
        oldPersonlst.pop(1)
        newPersonlst = self.GetPersonInfo()
        newPersonlst.insert(0, oldPersonlst[0])
        # Update Database
        PyDatabase.DBUpdate(oldPersonlst, newPersonlst, self.DTable2)        
        # Update list
        [self.list2.SetStringItem(listIndex2, i+3, newPersonlst[i+1]) for i in range(len(newPersonlst)-2)]        
        
        self.Text_Reason.SetValue("")
        self.Text_Demo.SetValue("")
    
    def OnDelete(self, event):
        listIndex2 = self.list2.GetFirstSelected()
        if listIndex2 == -1:
            wx.MessageBox(u"请点击选中已经被批假的人员。", u"提示")
            return
        
        oldPersonlst = [self.list2.GetItem(listIndex2, col=i+1).GetText() for i in range(self.list2.GetColumnCount()-1)]
        oldPersonlst.pop(1)
        # Update Database
        PyDatabase.DBDelete(oldPersonlst, self.DTable2)
        # Update list
        self.list2.DeleteItem(listIndex2)
        [self.list2.SetStringItem(i, 0, `i+1`) for i in range(self.list2.GetItemCount())]
        
        # Update the self.lstPHR
        self.lstPHR.remove(oldPersonlst[0][:g_UnitSnNum])
        
        self.OnSelect(None)
        self.Text_Reason.SetValue("")
        self.Text_Demo.SetValue("")
        self.DispColorList(self.list2)
    
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
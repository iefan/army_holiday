#!/usr/local/bin/python
# -*- coding: utf-8 -*-

#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import images
import os
import PyDatabase
import xlrd
import string
import MyThread
import time
import ReadUnit

from PhrResource import strVersion

class FrmImport(wx.Dialog): 
    def __init__(self): 
        title = strVersion
        wx.Dialog.__init__(self, None, -1, title, \
            style= wx.DEFAULT_DIALOG_STYLE)
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
        
        self.panel = panel = wx.Panel(self, -1)
        panel.Bind(MyThread.EVT_UPDATE_BARGRAPH, self.OnUpdate)
        panel.Bind(MyThread.EVT_IMPORT_XLS, self.OnImportXLS)
        panel.Bind(MyThread.EVT_IMPORT_XLS2, self.OnImportXLS)
        
        self.DTable = 'PersonInfo'
        
        self.thread = []
        
        lbltitle = wx.StaticText(self, -1, u"数据导入")
        lbltitle.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(lbltitle, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 0)
        
        lbl1 = wx.StaticText(self, -1, u"请选择人员信息的 Excel 文件路径：")
        self.ckbDeltaAdd = wx.CheckBox(self, -1, u"增量导入")
        self.Text_Path = wx.TextCtrl(self, -1, size=(300,-1))
        self.Text_Path.SetEditable(False)
        self.btn_Choice = wx.Button(self, -1, u"浏览")
        
        ckbsizer = wx.BoxSizer(wx.HORIZONTAL)
        ckbsizer.Add(lbl1, 0, wx.ALL| wx.ALIGN_LEFT)
        ckbsizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        ckbsizer.Add(self.ckbDeltaAdd, 0, wx.ALL| wx.ALIGN_RIGHT)
        
        pathsizer = wx.BoxSizer(wx.HORIZONTAL)
        pathsizer.Add(self.Text_Path, 0, wx.ALL, 0)
        pathsizer.Add(self.btn_Choice, 0, wx.ALL, 0)
        
        lbl2 = wx.StaticText(self, -1, u"请输入人员编号(四个字符，以字母开头)：")
        self.Text_Sn = wx.TextCtrl(self, -1, "BZDD")
        
        snsizer = wx.BoxSizer(wx.HORIZONTAL)
        snsizer.Add(lbl2, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL, 0)
        snsizer.Add(self.Text_Sn, 0, wx.ALL, 5)
        
        self.btn_InitBase = wx.Button(self, -1, u"初始化数据库")
        self.btn_CheckData = wx.Button(self, -1, u"检查")
        self.btn_Importdata = wx.Button(self, -1, u"导入")
        
        importsizer = wx.BoxSizer(wx.HORIZONTAL)
        importsizer.Add(self.btn_InitBase, 1, wx.ALL| wx.ALIGN_LEFT, 10)
        importsizer.Add(self.btn_CheckData, 1, wx.ALL| wx.ALIGN_CENTER, 10)
        importsizer.Add(self.btn_Importdata, 1, wx.ALL| wx.ALIGN_RIGHT, 10)
        
        btn_Close = wx.Button(self, -1, u"关闭")
        self.g1 = wx.Gauge(self, -1, 50, size=(-1,10))
        btnsizer = wx.BoxSizer(wx.HORIZONTAL)
        btnsizer.Add(self.g1, 10, wx.ALL| wx.EXPAND)
        btnsizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        btnsizer.Add(btn_Close, 0, wx.ALL, 0)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 10)        
        mainSizer.Add( wx.StaticLine(self), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.Add(ckbsizer, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.AddSizer(pathsizer, 0, wx.ALL, 5)
        mainSizer.AddSizer(snsizer, 0, wx.ALL, 5)
        mainSizer.AddSizer(importsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add( wx.StaticLine(self), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.AddSizer(btnsizer, 0, wx.ALL| wx.EXPAND, 5)
        
        self.SetSizer(mainSizer)
        mainSizer.Fit(self)
        self.Center()
        self.OnInitData()
        
        self.Bind( wx.EVT_CHECKBOX, self.OnCkbDelta, self.ckbDeltaAdd) 
        self.Bind(wx.EVT_BUTTON, self.OnInitBase, self.btn_InitBase) 
        self.Bind(wx.EVT_BUTTON, self.OnCheckData, self.btn_CheckData) 
        self.Bind(wx.EVT_BUTTON, self.OnImportdata, self.btn_Importdata) 
        self.Bind(wx.EVT_BUTTON, self.OnChoicePath, self.btn_Choice) 
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
    
    def OnInitData(self):
        self.g1.SetValue(1)
        self.g1.Show(False)
        
    def OnCkbDelta(self, event):
        strResult = PyDatabase.DBSelect("", self.DTable, ['Sn'], 0)
        if len(strResult) == 0:
            wx.MessageBox(u"数据库为空，无法进行增量导入！", u"提示")
            self.ckbDeltaAdd.SetValue(False)
            return
        
        self.btn_InitBase.Enable(not self.ckbDeltaAdd.GetValue())
        self.Text_Sn.Enable(not self.ckbDeltaAdd.GetValue())
    
    def OnImportXLS(self, evt):
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
            wx.MessageBox(u"导入成功！", u"提示")
            
            self.btn_InitBase.Enable(True)
            self.btn_CheckData.Enable(True)
            self.btn_Choice.Enable(True)
    
    def OnUpdate(self, evt):
        self.g1.SetValue(evt.count)
        
    def OnInitBase(self, event):
        wx.MessageBox(u"初始化将会丢掉您当前数据库中所有的人员信息，请做好备份！", u"提示")
        if wx.OK == wx.MessageBox(u"您确定要初始化吗！", u"提示", wx.OK| wx.CANCEL):
            PyDatabase.DBInit()
        self.btn_InitBase.Enable(False)

    def CheckData2(self, sh):
        # 1. check the unit
        strUnitlst = ReadUnit.readunit(sh)
        unitExist = PyDatabase.DBSelect("UnitSn like '%%%%'", "UnitTab", ['UnitName'], 1)
        tmplst = []
        for item in unitExist:
            tmplst.append(item[0].split('|->')[-1])
            
        for iuNew in strUnitlst:
            if iuNew not in tmplst:
                wx.MessageBox(u"新增人员中有目前数据库未包含的单位！请手动添加！", u"提示")
                return False
            
        # 2. check the rank
        xlsTrueHead = [sh.row(0)[i].value for i in range(sh.ncols)]
        rankExist = PyDatabase.DBSelect("RankSn like '%%%%'", "RankDays", ['LevelRank'], 1)
        tmplst = []
        for item in rankExist:
            tmplst.append(item[0])
        for i in range(1, sh.nrows):
            ranktmp = sh.row(i)[xlsTrueHead.index(u"军衔级别")].value
            if ranktmp not in tmplst:
                wx.MessageBox(u"新增人员中有目前数据库未包含的军衔级别！请手动添加！", u"提示")
                return False

        # 3. check the road
        roadExist = PyDatabase.DBSelect("AddrSn like '%%%%'", "RoadDays", ['Address'], 1)
        tmplst = []
        for item in roadExist:
            tmplst.append(item[0][:2])
        for i in range(1, sh.nrows):
            roadtmp = sh.row(i)[xlsTrueHead.index(u"籍贯")].value[:2]
            if roadtmp not in tmplst:
                wx.MessageBox(u"新增人员中有目前数据库未包含的籍贯！请手动添加！", u"提示")
                return False
            
        # 4. Set the self.Text_Sn.SetValue()
        strResult = PyDatabase.DBSelect("Sn like '%%%%'", self.DTable, ['Sn'], 1)
        self.Text_Sn.SetValue(strResult[0][0][:4])
        return True
        
    def OnCheckData(self, event):
        if self.Text_Path.GetValue() == "":
            wx.MessageBox(u"请点击浏览人员信息文件！", u"提示")
            return
        try:
            bk = xlrd.open_workbook(self.Text_Path.GetValue())
            sh = bk.sheets()[0]
        except:
            wx.MessageBox(u"无法读取数据文件！请检查数据文件格式！", u"错误", wx.ICON_ERROR)
            return
        
        # check the first row of the xls file
        xlsHead = [u"姓名", u"单位",u"性别", u"军衔级别",  u"籍贯", u"婚姻状况", u"联系电话", u"通讯地址", u"军衔时间"]
        if sh.ncols < len(xlsHead):
            wx.MessageBox(u"文件的第一行必须只包含\n\n" + string.join(xlsHead, '-') + u"\n\n来排列！", u"错误", wx.ICON_ERROR)
            return
        
        for index in range(len(xlsHead)):
            if sh.row(0)[index].value not in xlsHead:
                wx.MessageBox(u"文件的第一行必须只包含\n\n" + string.join(xlsHead, '-') + u"\n\n来排列！", u"错误", wx.ICON_ERROR)
                return
        
        if self.ckbDeltaAdd.GetValue():
            if not self.CheckData2(sh):
                return
        else:
            if self.btn_InitBase.IsEnabled():
                wx.MessageBox(u"请先初始化数据库！！！", u"提示")
                return

        wx.MessageBox(u"检查通过", u"提示") 
        self.btn_Importdata.Enable(True)
    
    def OnImportdata(self, event):
        if self.Text_Path.GetValue() == "":
            wx.MessageBox(u"请点击浏览人员信息文件！", u"提示")
            return
        if self.Text_Sn.GetValue() == "":
            wx.MessageBox(u"请输入四位英文字符", u"提示")
            return
        
        personSn = self.Text_Sn.GetValue().strip()                
        if len(personSn) != 4:
            wx.MessageBox(u"请输入四位的英文字符", u"提示")
            return
        for iab in personSn:
            if iab not in string.ascii_letters:
                wx.MessageBox(u"请输入全部英文字符", u"提示")
                return
        personSn = string.upper(personSn)
        
        self.btn_CheckData.Enable(False)
        self.btn_Importdata.Enable(False)
        self.btn_Choice.Enable(False)
        
        if self.ckbDeltaAdd.GetValue():
            self.thread = []
            self.thread.append(MyThread.CalcGaugeThread(self.panel, 0))
            self.thread.append(MyThread.ImportXlsThread2(self.panel, personSn, self.Text_Path.GetValue()))
            self.g1.Show(True)
            for item in self.thread:
                item.Start()
        else:
            self.thread = []
            self.thread.append(MyThread.CalcGaugeThread(self.panel, 0))
            self.thread.append(MyThread.ImportXlsThread(self.panel, personSn, self.Text_Path.GetValue()))
            self.g1.Show(True)
            for item in self.thread:
                item.Start()
  
    def OnChoicePath(self, event):
        try:
            dlg = wx.FileDialog(self, u"选择文件", "", "",
                                u"人员信息文件 (*.xls)|*.xls", wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                self.Text_Path.SetValue(dlg.GetPath())
            dlg.Destroy()
        except:
            wx.MessageBox(u"未能打开文件浏览对话框，请关闭本窗体，再试一下！", u"提示")
            return
        self.btn_Importdata.Enable(False)
    
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        if self.g1.IsShown():
            return
        else:
            self.Destroy()

if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmImport() 
    frame.ShowModal()
    app.MainLoop()
    

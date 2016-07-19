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
import ReadUnit

from PhrResource import strVersion

class FrmImport(wx.Dialog): 
    def __init__(self): 
        title = strVersion
        wx.Dialog.__init__(self, None, -1, title, \
            style= wx.DEFAULT_DIALOG_STYLE)
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
        
        panel = wx.Panel(self, -1)
        
        lbltitle = wx.StaticText(self, -1, u"数据导入")
        lbltitle.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(lbltitle, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 0)
        
        lbl1 = wx.StaticText(self, -1, u"请选择人员信息的 Excel 文件路径：")
        self.Text_Path = wx.TextCtrl(self, -1, size=(300,-1))
        self.Text_Path.SetEditable(False)
        btn_Choice = wx.Button(self, -1, u"浏览")
        
        pathsizer = wx.BoxSizer(wx.HORIZONTAL)
        pathsizer.Add(self.Text_Path, 0, wx.ALL, 0)
        pathsizer.Add(btn_Choice, 0, wx.ALL, 0)
        
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
        btnsizer = wx.BoxSizer(wx.HORIZONTAL)
        btnsizer.Add(btn_Close, 0, wx.ALL| wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM, 0)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 10)        
        mainSizer.Add( wx.StaticLine(self), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.Add(lbl1, 0, wx.ALL| wx.ALIGN_LEFT, 5)
        mainSizer.AddSizer(pathsizer, 0, wx.ALL, 5)
        mainSizer.AddSizer(snsizer, 0, wx.ALL, 5)
        mainSizer.AddSizer(importsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add( wx.StaticLine(self), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.AddSizer(btnsizer, 0, wx.ALL| wx.ALIGN_RIGHT, 5)
        
        self.SetSizer(mainSizer)
        mainSizer.Fit(self)
        self.Center()
        
        self.Bind(wx.EVT_BUTTON, self.OnInitBase, self.btn_InitBase) 
        self.Bind(wx.EVT_BUTTON, self.OnCheckData, self.btn_CheckData) 
        self.Bind(wx.EVT_BUTTON, self.OnImportdata, self.btn_Importdata) 
        self.Bind(wx.EVT_BUTTON, self.OnChoicePath, btn_Choice) 
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

    def OnInitBase(self, event):
        wx.MessageBox(u"初始化将会丢掉您当前数据库中所有的人员信息，请做好备份！", u"提示")
        if wx.OK == wx.MessageBox(u"您确定要初始化吗！", u"提示", wx.OK| wx.CANCEL):
            PyDatabase.DBInit()
            wx.MessageBox(u"初始化完成！请导入数据", u"提示")
#        self.btn_InitBase.Enable(False)
                
    def OnCheckData(self, event):
        if self.Text_Path.GetValue() == "":
            return
        try:
            bk = xlrd.open_workbook(self.Text_Path.GetValue())
            sh = bk.sheets()[0]
        except:
            wx.MessageBox(u"无法读取数据文件！请检查数据文件格式！", u"错误", wx.ICON_ERROR)
            return
        
        # check the first row of the xls file
        xlsHead = [u"姓名", u"性别", u"军衔级别", u"单位", u"籍贯", u"婚姻状况"]
        for index in range(6):
            if sh.row(0)[index].value != xlsHead[index]:
                wx.MessageBox(u"文件的第一行请按照\n\n姓名 性别 军衔级别 单位 籍贯 婚姻状况\n\n来排列！", u"错误", wx.ICON_ERROR)
                return
        
        wx.MessageBox(u"检查通过", u"提示")
        self.btn_Importdata.Enable(True)
    
    def OnImportdata(self, event):
        if self.Text_Path.GetValue() == "":
            wx.MessageBox(u"请点击浏览人员信息文件！", u"提示")
            return
        if self.Text_Sn.GetValue() == "":
            wx.MessageBox(u"请输入四位的英文字符", u"提示")
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
        bk = xlrd.open_workbook(self.Text_Path.GetValue())
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
        strUnitlst = ReadUnit.readunit(self.Text_Path.GetValue())
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
        for i in range(1, sh.nrows):
            tmplist = [None, personSn+string.zfill(str(i),4)]
            tmplist.append(sh.row(i)[0].value)  # name
            tmplist.append(sh.row(i)[1].value)  # sex
            tmplist.append(rankDict.keys()[rankDict.values().index(sh.row(i)[2].value)]) # rank
            
            # unit
            unittmp = sh.row(i)[3].value
            for iUnit in unitDict.values():
                if unittmp[:len(iUnit)] == iUnit:
                    tmplist.append(unitDict.keys()[unitDict.values().index(iUnit)])
            # road 
            tmproad = [None]
            if sh.row(i)[4].value[:2] not in roadtmp:
                roadtmp.append(sh.row(i)[4].value[:2])
                tmproad.append("ASN"+string.zfill(str(addrCount),3))
                tmproad.append(sh.row(i)[4].value[:2])
                tmproad.append('4')
                tmproad.append('20')
                addrCount += 1
                PyDatabase.DBInsert(tmproad, "RoadDays")
                
            # road
            roadindex = roadtmp.index(sh.row(i)[4].value[:2]) + 1
            tmplist.append("ASN"+string.zfill(str(roadindex),3))
            tmplist.append(sh.row(i)[4].value)
            tmplist.append(sh.row(i)[5].value)
            # tel
            tmplist.append("")
            PyDatabase.DBInsert(tmplist, "PersonInfo")
            
        strResult = PyDatabase.DBSelect("", "PersonInfo", ["Sn"], 0)
        
        wx.MessageBox(u"导入完成！", u"提示")
  
    def OnChoicePath(self, event):
        path = ""
        name = ""
        if self.Text_Path.GetValue():
            path, name = os.path.split(self.Text_Path.GetValue())
        
        dlg = wx.FileDialog(self, u"选择文件", path, name,
                            u"人员信息文件 (*.xls)|*.xls", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.Text_Path.SetValue(dlg.GetPath())
        dlg.Destroy()
        self.btn_Importdata.Enable(False)
    
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()

if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmImport() 
    frame.ShowModal()
    app.MainLoop()
    

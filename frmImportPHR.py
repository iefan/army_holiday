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
        
        self.lstUser = []
        self.lstPwd = []
        self.lstAdminLevel = []
        
        panel = wx.Panel(self, -1)
        
        lbltitle = wx.StaticText(self, -1, u"数据导入")
        lbltitle.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(lbltitle, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 0)
        
        lbl1 = wx.StaticText(self, -1, u"请选择备份文件路径：")
        self.Text_Path = wx.TextCtrl(self, -1, size=(300,-1))
        self.Text_Path.SetEditable(False)
        btn_Choice = wx.Button(self, -1, u"浏览")
        
        pathsizer = wx.BoxSizer(wx.HORIZONTAL)
        pathsizer.Add(self.Text_Path, 0, wx.ALL, 0)
        pathsizer.Add(btn_Choice, 0, wx.ALL, 0)
        
        lbl2 = wx.StaticText(self, -1, u"请输入所备份文件用户名：")
        self.Text_Name = wx.TextCtrl(self, -1, "")
        lbl3 = wx.StaticText(self, -1, u"请输入密码：")
        self.Text_Pwd = wx.TextCtrl(self, -1, "", style = wx.TE_PASSWORD)
        
        snsizer = wx.BoxSizer(wx.VERTICAL)
        snsizer.Add(lbl2, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL, 5)
        snsizer.Add(self.Text_Name, 0, wx.ALL, 5)
        snsizer.Add(lbl3, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL, 5)
        snsizer.Add(self.Text_Pwd, 0, wx.ALL, 5)
                
        self.btn_Importdata = wx.Button(self, -1, u"导入")
        btn_Close = wx.Button(self, -1, u"关闭")
        btnsizer = wx.BoxSizer(wx.HORIZONTAL)
        btnsizer.Add(self.btn_Importdata, 0, wx.ALL| wx.ALIGN_BOTTOM, 10)
        btnsizer.Add(btn_Close, 0, wx.ALL| wx.ALIGN_BOTTOM, 10)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 10)        
        mainSizer.Add( wx.StaticLine(self), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.Add(lbl1, 0, wx.ALL| wx.ALIGN_LEFT, 5)
        mainSizer.AddSizer(pathsizer, 0, wx.ALL, 5)
        mainSizer.AddSizer(snsizer, 0, wx.ALL, 5)
        mainSizer.Add( wx.StaticLine(self), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.AddSizer(btnsizer, 0, wx.ALL| wx.ALIGN_RIGHT, 5)
        
        self.SetSizer(mainSizer)
        mainSizer.Fit(self)
        self.Center()

        self.Bind(wx.EVT_BUTTON, self.OnImportdata, self.btn_Importdata) 
        self.Bind(wx.EVT_BUTTON, self.OnChoicePath, btn_Choice) 
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

    def OnImportdata(self, event):
        if self.Text_Path.GetValue() == "":
            wx.MessageBox(u"请点击浏览人员备份文件！", u"提示")
            return
        
        txtUser = self.Text_Name.GetValue().strip()
        txtPwd = self.Text_Pwd.GetValue().strip() 
        
        if txtUser == "" or txtPwd == "":
            wx.MessageBox(u"请输入所导入的备份文件的用户名和密码！", u"提示")
            return
        
        userLib = PyDatabase.DBSelect('', 'UserLib', ['UserName'], 0)        
        for item in userLib:
            self.lstUser.append(item[1])
            self.lstPwd.append(item[2])
            self.lstAdminLevel.append(item[3])
            
        for itemUser, itemPwd in zip(self.lstUser, self.lstPwd):
            if itemUser == txtUser and itemPwd == txtPwd:
                index = self.lstUser.index(txtUser) 
                if self.lstAdminLevel[index] == u'管理员':
                    strCmd = 'copy /Y "' + self.Text_Path.GetValue() + '"' + " HolidayData"
                    if os.system(strCmd.encode('gbk')) == 0:
                        wx.MessageBox(u"导入成功！", u"提示")
                    else:
                        wx.MessageBox(u"未能成功导入，请检查所选择的备份文件！", u"提示")
                    return
                else:
                    wx.MessageBox(u"请输入所导入的备份文件的管理员用户名和密码！", u"提示")
                    return
            else:
                wx.MessageBox(u"对不起，您输入的用户名与密码错误！数据未能导入！", u"错误")
                return
        
    def OnChoicePath(self, event):
        dlg = wx.FileDialog(self, u"选择文件", "", "",
                            u"休假系统数据备份文件 (*.phr)|*.phr", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.Text_Path.SetValue(dlg.GetPath())
        dlg.Destroy()
    
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()

if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmImport() 
    frame.ShowModal()
    app.MainLoop()
    

#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
#import wx.lib.buttons as buttons
import PyDatabase
import images
from PhrResource import strVersion, MyListCtrl
import os

class FrmUserMaintain(wx.Frame):
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        panel = wx.Panel(self)
        
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        UserFix(panel)
        self.Maximize()
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()
                    
class UserFix(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)
                
        titleText = wx.StaticText(panel, -1, u"用户信息维护")
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
        
        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT| wx.LC_HRULES| wx.LC_VRULES)
        self.list.SetImageList(wx.ImageList(1, 24), wx.IMAGE_LIST_SMALL)

        listsizer = wx.BoxSizer(wx.HORIZONTAL)
        listsizer.Add(self.list, 1, wx.ALL|wx.EXPAND, 10)
        
        lbl1 = wx.StaticText(panel, -1, u"类型")
        self.lbLevel = wx.ComboBox(panel, -1, u"用户")
        self.ckbUserLbl = wx.CheckBox(panel, -1, u"新增用户名：")
        self.Text_User = wx.TextCtrl(panel, -1, "")
        self.Text_User.Enable(False)
        lbl3 = wx.StaticText(panel, -1, u"输入密码：")
        self.Text_NewPwd = wx.TextCtrl(panel, style=wx.TE_PASSWORD)
        lbl4 = wx.StaticText(panel, -1, u"确认密码：")
        self.Text_NewPwd2 = wx.TextCtrl(panel, style=wx.TE_PASSWORD)
        infosizer = wx.BoxSizer(wx.VERTICAL)
        itemInfo = [lbl1, self.lbLevel, self.ckbUserLbl, self.Text_User, lbl3, self.Text_NewPwd, lbl4, self.Text_NewPwd2]
        for index in range(len(itemInfo)):
            if index%2 == 0 and index!=0:
                infosizer.Add((-1, 20))
            infosizer.Add(itemInfo[index], 0, wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)

        listsizer.AddSizer(infosizer, 0, wx.ALL| wx.ALIGN_RIGHT, 10)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 
        
        self.listIndex = -1
        self.DTable = 'UserLib'

        lstHead = [u"序号", u"用户名", u"级别"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]

        list_Level = [u"用户", u"管理员"]
        self.lbLevel.SetItems(list_Level)
        self.lbLevel.Select(1)
        self.lbLevel.SetEditable(False)
            
#        self.infoUser = [self.Text_User, self.Text_Pwd, self.lbLevel]
        
        self.OnSelect(None)
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)
        
        panel.popMenuUser = wx.Menu()
        pmList_1 = panel.popMenuUser.Append(1171, u"删除")        
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, pmList_1)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPop)
        
        panel.Show()
        panel.Bind( wx.EVT_CHECKBOX, self.OnCkbUser, self.ckbUserLbl)
        panel.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.list)
        panel.Bind(wx.EVT_BUTTON, self.OnAdd, btn_Add)       
        panel.Bind(wx.EVT_BUTTON, self.OnDelete, btn_Delete)       
        panel.Bind(wx.EVT_BUTTON, self.OnModify, btn_Modify)       
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
    
    def OnCkbUser(self, event):
        self.Text_User.Enable(self.ckbUserLbl.GetValue())
    
    def OnShowPop(self, event):
        if self.list.GetItemCount() != 0:
            self.list.PopupMenu(self.panel.popMenuUser)
         
    def OnPopItemSelected(self, event):
        try:
            item = self.panel.popMenuUser.FindItemById(event.GetId())        
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
        strResult = PyDatabase.DBSelect('', self.DTable, ['UserName'], 0)
        for row in strResult:
            index = self.list.InsertStringItem(100, "A") 
            self.list.SetStringItem(index, 0, `index+1`)
            self.list.SetStringItem(index, 1, row[1])
            self.list.SetStringItem(index, 2, row[3])
        self.DispColorList(self.list)

    def OnItemSelected(self, event):
        self.listIndex = event.GetIndex()
        self.lbLevel.SetValue(self.list.GetItem(self.listIndex, col=2).GetText())

    def ClearTxt(self):
        self.Text_NewPwd.SetValue("") 
        self.Text_NewPwd2.SetValue("") 
        
    def OnAdd(self, event):
        if not self.ckbUserLbl.GetValue():
            wx.MessageBox(u"请点击新增用户！", u"提示")
            return
        
        if self.Text_User.GetValue().strip() == "" or self.Text_NewPwd.GetValue().strip() == "" or self.Text_NewPwd2.GetValue().strip() == "":
            wx.MessageBox(u"请填写用户名和密码！", u"提示")
            return
        
        newUser = self.Text_User.GetValue().strip()
        if newUser in [self.list.GetItem(i, col=1).GetText() for i in range(self.list.GetItemCount())]:
            wx.MessageBox(u"已经存在此用户！", u"提示")
            return
        
        if self.Text_NewPwd.GetValue() != self.Text_NewPwd2.GetValue():
            wx.MessageBox(u"两次输入密码不正确！", u"提示")
            self.ClearTxt()
            return
        
        listUser = [None]
        listUser.append(self.Text_User.GetValue().strip())
        listUser.append(self.Text_NewPwd.GetValue().strip())
        listUser.append(self.lbLevel.GetValue())
        # Update Database  
        PyDatabase.DBInsert(listUser, self.DTable)
        
        # Update List
        index = self.list.InsertStringItem(100, "A")
        self.list.SetStringItem(index, 0, `index+1`)
        self.list.SetStringItem(index, 1, listUser[1])
        self.list.SetStringItem(index, 2, listUser[3])
        self.ClearTxt()
        self.listIndex = index
        self.DispColorList(self.list)
    
    def OnModify(self, event): 
        if self.listIndex == -1:
            wx.MessageBox(u"请选中一用户！", u"提示")
            return
        
        oldUser = self.list.GetItem(self.listIndex, col=1).GetText()
        strResult = PyDatabase.DBSelect("UserName= '" + oldUser + "'", self.DTable, ["UserName", "Pwd", "Level"], 1)
        
        dlg = wx.TextEntryDialog(self.panel, u"请输入 " + oldUser + u" 的原密码：",\
            u"提示", style= wx.OK| wx.CANCEL|wx.TE_PASSWORD)

        if dlg.ShowModal() == wx.ID_OK:
            if strResult[0][1] != dlg.GetValue():
                wx.MessageBox(u"所输入的用户密码不正确！", u"提示")
                dlg.Destroy()
                return
        else:
            dlg.Destroy()
            return            
    
        if self.Text_NewPwd.GetValue().strip() == "" or self.Text_NewPwd2.GetValue().strip() == "":
            wx.MessageBox(u"请填写密码！", u"提示")
            return
        
        if self.Text_NewPwd.GetValue() != self.Text_NewPwd2.GetValue():
            wx.MessageBox(u"两次输入密码不正确！", u"提示")
            self.ClearTxt()
            return
                
        lstUserOld = strResult[0]
        lstUserNew = [lstUserOld[0], self.Text_NewPwd.GetValue(), self.lbLevel.GetValue()]
      
        # Update Database
        PyDatabase.DBUpdate(lstUserOld, lstUserNew, self.DTable)        
        # Update list
        self.list.SetStringItem(self.listIndex, 2, lstUserNew[2])
        self.ClearTxt()
        
    def OnDelete(self, event):                
        if self.listIndex == -1:
            wx.MessageBox(u"请选中一用户", u"提示")
            return
        
        oldUser = self.list.GetItem(self.listIndex, col=1).GetText()
        strResult = PyDatabase.DBSelect("UserName= '" + oldUser + "'", self.DTable, ["UserName", "Pwd", "Level"], 1)
        
        dlg = wx.TextEntryDialog(self.panel, u"请输入 " + oldUser + u" 的原密码：",\
            u"提示", style= wx.OK| wx.CANCEL|wx.TE_PASSWORD)

        if dlg.ShowModal() == wx.ID_OK:
            if strResult[0][1] != dlg.GetValue():
                wx.MessageBox(u"所输入的用户密码不正确！", u"提示")
                dlg.Destroy()
                return
        else:
            dlg.Destroy()
            return
        
        lstUser = strResult[0]
        
        # Update Database
        PyDatabase.DBDelete(lstUser, self.DTable)
        # Update list
        self.list.DeleteItem(self.listIndex)
        # Resort the first colomn
        [self.list.SetStringItem(i, 0, `i+1`) for i in range(self.list.GetItemCount())]
        self.ClearTxt()
        self.listIndex = -1
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
    fmmaintain = FrmUserMaintain()
    fmmaintain.Show()
    app.MainLoop()
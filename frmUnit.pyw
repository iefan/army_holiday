#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import PyDatabase
import images
import types
import string
import pickle
import os

from PhrResource import strVersion, MyListCtrl

class FrmSubMaintain(wx.Frame):
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title)
        panel = wx.Panel(self)
        self.Maximize()        
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        UnitFix(panel)
        
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def OnCloseWindow(self, event):
        self.Destroy()

class UnitFix(object):
    def __init__(self, panel):
        self.panel = panel
        panel.Hide()
        panel.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        try:
            color = panel.GetParent().GetMenuBar().GetBackgroundColour()
        except:
            color = (236, 233, 216)
        panel.SetBackgroundColour(color)

        titleText = wx.StaticText(panel, -1, u"单位信息维护")
        titleText.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(titleText, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL)

        btn_Ok = wx.Button(panel, -1, u"完成")
        btn_Help = wx.Button(panel, -1, u"帮助")
        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL) 
        for item in [btn_Ok, btn_Help]: 
            btnSizer.Add((20,-1), 1)
            btnSizer.Add(item)
        btnSizer.Add((20,-1), 1)
        
        self.list = MyListCtrl(panel, -1, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES)
        self.list.SetImageList(wx.ImageList(1, 20), wx.IMAGE_LIST_SMALL)
        
        listsizer = wx.BoxSizer(wx.HORIZONTAL)
        listsizer.Add(self.list, 1, wx.ALL|wx.EXPAND, 5)
        
        lbl1 = wx.StaticText(panel, -1, u"单位树")
        #Create tree
        self.treeDict = {}
        self.tree = wx.TreeCtrl(panel, size=(200, 300))
        self.root = self.tree.AddRoot(u"单位")
        lstTreeBtnTxt = [u"提升",u"下降",u"上移",u"下移"]
        lstTreeID = [6201,6202,6203,6204]
        lstTreeBtn = []
        lstTreeBtnsizer = wx.BoxSizer(wx.HORIZONTAL)
        for item in zip(lstTreeBtnTxt, lstTreeID):
            btntmp = wx.Button(panel, item[1], item[0], size=(50, -1))
            lstTreeBtn.append(btntmp)
            panel.Bind(wx.EVT_BUTTON, self.OnTreeBtn, btntmp)
            lstTreeBtnsizer.Add(btntmp, 0, wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        
        lbl2 = wx.StaticText(panel, -1, u"批假单位：")
        self.Text_Permit = wx.TextCtrl(panel, -1, u"大队综合处")
        btn_AddPUnit = wx.Button(panel, -1, u"修改")
        pUnitsizer = wx.BoxSizer(wx.HORIZONTAL)
        pUnitsizer.Add(self.Text_Permit, 1, wx.ALL| wx.EXPAND)
        pUnitsizer.Add(btn_AddPUnit, 0, wx.ALL)
        
        infosizer = wx.BoxSizer(wx.VERTICAL)
        infosizer.Add(lbl1, 0, wx.ALIGN_LEFT)
        infosizer.Add(self.tree, 0, wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        infosizer.AddSizer(lstTreeBtnsizer)
        infosizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        infosizer.Add(lbl2, 0, wx.ALIGN_LEFT)
        infosizer.Add(pUnitsizer, 0, wx.ALL| wx.EXPAND)
        infosizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)

        listsizer.AddSizer(infosizer, 0, wx.ALL|wx.EXPAND, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.AddSizer(topsizer, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.AddSizer(listsizer, 1, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 5) 

        self.listIndex = 0
        self.DTable = 'UnitTab'
        
        lstHead = [u"序号", u"单位编号", u"单位名称"]
        [self.list.InsertColumn(i, item) for (i, item) in zip(range(len(lstHead)), lstHead)]

        self.OnSelect()
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)                                      
        mainSizer.SetSizeHints(panel)
        
        panel.popMenuUnit = wx.Menu()
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, panel.popMenuUnit.Append(1411, u"增加"))
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, panel.popMenuUnit.Append(1412, u"修改"))
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, panel.popMenuUnit.AppendSeparator())
        panel.Bind(wx.EVT_MENU, self.OnPopItemSelected, panel.popMenuUnit.Append(1413, u"删除"))
        self.tree.Bind(wx.EVT_RIGHT_DOWN, self.OnRightDown)
        
        self.OnInitData()
        
        panel.Show()
        panel.Bind(wx.EVT_BUTTON, self.OnOk, btn_Ok)
        panel.Bind(wx.EVT_BUTTON, self.OnHelp, btn_Help)
        panel.Bind(wx.EVT_BUTTON, self.OnModPUnit, btn_AddPUnit)
        panel.Bind(wx.EVT_PAINT, self.OnPaint)
    
    def OnInitData(self):
        try:
            pkl_file = open('ByData.pkl', 'rb')
        except:
            data = {'ByMan': u"",\
            'ByLeader': u"", \
            'ByUnit': u""}        
            output = open('ByData.pkl', 'wb')
            pickle.dump(data, output)
            output.close()
            self.Text_Permit.SetValue(data['ByUnit'])
            return
        
        data = pickle.load(pkl_file)
        pkl_file.close()
        self.Text_Permit.SetValue(data['ByUnit'])

    def OnModPUnit(self, event):
        if self.Text_Permit.GetValue().strip() == "":
            wx.MessageBox(u"请填入批假单位！", u"提示")
            return
        
        pkl_file = open('ByData.pkl', 'rb')
        data = pickle.load(pkl_file)
        pkl_file.close()
        
        data['ByUnit'] = self.Text_Permit.GetValue().strip()
        output = open('ByData.pkl', 'wb')
        pickle.dump(data, output)
        output.close()
        wx.MessageBox(u"修改成功！", u"提示")
    
    def OnOk(self, event):
        vnodelst = self.treeDict.values()
        newUSNlst = self.treeDict.keys()
        
        newUnitlst = []          # modified unit and usn
#        index = 0
        for item in vnodelst:
            strtmp = self.GetItemText(item)
            parenttmp = self.tree.GetItemParent(item)
            while parenttmp.IsOk():
                strtmp = self.GetItemText(parenttmp) + "|->" + strtmp
                parenttmp = self.tree.GetItemParent(parenttmp)
            newUnitlst.append(strtmp)
#            newUSNdict[newUSNlst[index]] = strtmp
#            index += 1
        oldUSNlst = [self.list.GetItem(i, col=1).GetText() for i in range(self.list.GetItemCount())]
        oldUnitlst = [self.list.GetItem(i, col=2).GetText() for i in range(self.list.GetItemCount())]
        
        # Update database
        for item in zip(oldUSNlst, oldUnitlst):
            lstinput = [item[0], item[1]]
            PyDatabase.DBDelete(lstinput, self.DTable)
        for item in zip(newUSNlst, newUnitlst):
            lstinput = [None, item[0], item[1]]
            PyDatabase.DBInsert(lstinput, self.DTable)
        
        # remove personinfo by iusn
        for iusn in oldUSNlst:
            if iusn not in newUSNlst:
                oldUSNlst.remove(iusn)
                        
        # update the list
        self.list.DeleteAllItems()
#        print newUSNlst
        for item in zip(newUSNlst, newUnitlst):
            index = self.list.InsertStringItem(100, "A")
            self.list.SetItemFont(index, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))  
            self.list.SetStringItem(index, 0, `index+1`)
            self.list.SetStringItem(index, 1, item[0])
            self.list.SetStringItem(index, 2, item[1])
        self.DispColorList(self.list)
       
    def OnRightDown(self, event):
        pt = event.GetPosition();
        item, flags = self.tree.HitTest(pt)
        if item:
            self.tree.SelectItem(item)
            self.tree.PopupMenu(self.panel.popMenuUnit)
    
    def GetItemText(self, item):
        if item:
            return self.tree.GetItemText(item)
        else:
            return ""
            
    def MoveTreeItemChild(self, newParent, oldParent):
        item, cookie = self.tree.GetFirstChild(oldParent)
        while item:
            newp = self.tree.AppendItem(newParent, self.GetItemText(item))
            # modify the self.treeDict
            self.treeDict[self.treeDict.keys()[self.treeDict.values().index(item)]] = newp
            if self.tree.ItemHasChildren(item):
                self.MoveTreeItemChild(newp, item)
            item, cookie = self.tree.GetNextChild(item, cookie)
            
    def OnTreeBtn(self, event):
        btnid = event.GetId()
        item = self.tree.GetSelection()
        if self.GetItemText(item) == "":
            return
        if btnid == 6201:       # Up Level
            if self.tree.GetItemParent(item).IsOk() == False:
                return
            parent = self.tree.GetItemParent(item)
            if self.tree.GetItemParent(parent).IsOk() == False:
                return
            newitem = self.tree.InsertItem(self.tree.GetItemParent(parent), parent, self.GetItemText(item))
                    
        elif btnid == 6202:     # Down Level
            if self.tree.GetPrevSibling(item):
                newParent = self.tree.GetPrevSibling(item)
            elif self.tree.GetNextSibling(item):
                newParent = self.tree.GetNextSibling(item)
            else:
                return
            newitem = self.tree.AppendItem(newParent, self.GetItemText(item))
            
        elif btnid == 6203:
            if self.tree.GetItemParent(item).IsOk() == False:
                return
            count = 0
            itemC, cookie = self.tree.GetFirstChild(self.tree.GetItemParent(item))
            while itemC:
                if self.GetItemText(item) == self.GetItemText(itemC):
                    break;
                count += 1
                itemC, cookie = self.tree.GetNextChild(item, cookie)
                
            if count == 0:
                return
            newitem = self.tree.InsertItemBefore(self.tree.GetItemParent(item), count-1, self.GetItemText(item))
            
        elif btnid == 6204:
            if self.tree.GetNextSibling(item):
                newitem = self.tree.InsertItem(self.tree.GetItemParent(item), self.tree.GetNextSibling(item), self.GetItemText(item))
            else:
                return
            
        if self.tree.ItemHasChildren(item):
            self.MoveTreeItemChild(newitem, item)
            
        # modify the self.treeDict
        self.treeDict[self.treeDict.keys()[self.treeDict.values().index(item)]] = newitem
        self.tree.Delete(item)
    
    def OnPopItemSelected(self, event):
        try:
            item = self.panel.popMenuUnit.FindItemById(event.GetId())        
            text = item.GetText()        
            if text == u"增加":
                self.OnAdd(None)
            elif text == u"修改":
                self.OnModify(None)
            elif text == u"删除":
                self.OnDelete(None)
        except:
            pass
            
    def DispColorList(self, list):
        for i in range(list.GetItemCount()):
            if i%4 == 0: list.SetItemBackgroundColour(i, (233, 233, 247))
            if i%4 == 1: list.SetItemBackgroundColour(i, (247, 247, 247))
            if i%4 == 2: list.SetItemBackgroundColour(i, (247, 233, 233))
            if i%4 == 3: list.SetItemBackgroundColour(i, (233, 247, 247))
            
    def OnSelect(self):
        self.list.DeleteAllItems()                
        strResult = PyDatabase.DBSelect('', self.DTable, ['UnitSn'], 0) 
        
#        self.treeDict = []
        lstunitname = []
        for row in strResult:
            index = self.list.InsertStringItem(100, "A")
            self.list.SetStringItem(index, 0, `index+1`)
            self.list.SetStringItem(index, 1, row[1])
            self.list.SetStringItem(index, 2, row[2])
            lstunitname.append((row[1], row[2].split('|->')))
        
        self.DispColorList(self.list)
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
        
    def OnAdd(self, event):
        item = self.tree.GetSelection()
        dlg = wx.TextEntryDialog(None, 
            u"请输入新添加的子结点名称", 
            u"添加单位", "", style=wx.OK|wx.CANCEL) 
        if dlg.ShowModal() == wx.ID_OK:
            treeid = self.tree.AppendItem(item, dlg.GetValue().strip())
        dlg.Destroy()
        
        # delete the deleted itemid in self.treeDict
        lstIntSn = [int(isn[-3:]) for isn in self.treeDict.keys()]
        index = 0;  newsn = -1
        for index in range(1, len(self.treeDict)+1):
            if index < lstIntSn[index-1]:
                newsn = index
                break
        if newsn == -1:
            newsn = index+1        
        self.treeDict[u'USN' + string.zfill(newsn, 3)] = treeid
        
    def OnModify(self, event):
        item = self.tree.GetSelection()
        dlg = wx.TextEntryDialog(None, 
            u"请输入修改的单位名称", 
            u"修改单位", self.GetItemText(item), style=wx.OK|wx.CANCEL) 
        if dlg.ShowModal() == wx.ID_OK:
            self.tree.SetItemText(item, dlg.GetValue().strip())
        dlg.Destroy()

    def OnDelete(self, event):
        item = self.tree.GetSelection()
        if self.tree.GetItemParent(item).IsOk() == False:
            wx.MessageBox(u"根结点不能删除！", u"提示")
            return
        if self.tree.ItemHasChildren(item):
            wx.MessageBox(u"该结点含有子结点，不能删除！", u"提示")
            return
        
        unitname = self.GetItemText(item)
        strRunit = PyDatabase.DBSelect("UnitName like '%" + unitname + "%'", self.DTable, ['UnitSn'], 1)
        
        assert len(strRunit) == 1
        
        strRSn = strRunit[0][0]
        strResult = PyDatabase.DBSelect(strRSn, "PersonInfo", ['UnitSn'], 2)
        if len(strResult) > 0:
            wx.MessageBox(u"当前数据库中存在该单位的人员，无法删除该单位！", u"提示")
            return
        
        # delete the deleted itemid in self.treeDict
        self.treeDict.pop(self.treeDict.keys()[self.treeDict.values().index(item)])
        self.tree.Delete(item)

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
#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import wx.html
import images

from PhrResource import strVersion, DevInfo, CopyRightInfo,AboutIMG

class FrmHelp(wx.Frame): 
    def __init__(self): 
        title = strVersion
        wx.Frame.__init__(self, None, -1, title + u"　--　帮助")
        self.Center()
        self.SetMinSize((500, 600))
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
        
        panel = wx.Panel(self)
        self.Maximize()
        
        lbltitle = wx.StaticText(panel, -1, u"说明")
        lbltitle.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(lbltitle, 0, wx.ALL| wx.ALIGN_CENTER_HORIZONTAL, 0)
                    
        self.tree = wx.TreeCtrl(panel, size=(200, -1))
        self.root = self.tree.AddRoot(u"军人休假登记系统")
        
        self.htmlwin = wx.html.HtmlWindow(panel)
        treesizer = wx.BoxSizer(wx.HORIZONTAL)
        treesizer.Add(self.tree, 0, wx.ALL| wx.EXPAND,5)
        treesizer.Add(self.htmlwin, 1, wx.ALL| wx.EXPAND, 5)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(topsizer, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 10)
        mainSizer.Add( wx.StaticLine(panel), 0, wx.ALL| wx.EXPAND, 0)
        mainSizer.Add(treesizer, 1, wx.ALL| wx.EXPAND, 5)
        
        panel.SetSizer(mainSizer)
        mainSizer.Fit(panel)
        
        self.InitData()
        self.treeDict = {}
        self.Center()
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
    def InitData(self):
        lststr = [u"主界面", u"登录界面", u"批假管理"]
        [self.tree.AppendItem(self.root, item) for item in lststr]
        self.tree.Expand(self.root)

    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()

if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmHelp() 
    frame.Show()
    app.MainLoop()
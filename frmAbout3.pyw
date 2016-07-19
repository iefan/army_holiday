#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import images
from PhrResource import strVersion, DevInfo2

class FrmAbout(wx.Frame): 
    def __init__(self): 
        title = strVersion
        meWidth = 450
        meHeight = 350
        wx.Frame.__init__(self, None, -1, title, size=(meWidth, meHeight),\
            style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER ^ wx.MAXIMIZE_BOX ^ wx.MINIMIZE_BOX)
        self.Center()
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
#        self.SetTransparent(240)

        panel = wx.Panel(self)
        imgLogo = wx.StaticBitmap(panel, -1, wx.Image("Bitmap\\aboutLogo.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap())
        titleText = wx.StaticText(panel, -1, title)
        titleText.SetFont(wx.Font(19, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
        topsizer.Add(imgLogo, 0, wx.ALL)
        topsizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        topsizer.Add(titleText, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL)
        
        self.tree = wx.TreeCtrl(panel, style= wx.TR_DEFAULT_STYLE|wx.TR_HIDE_ROOT)
        self.tree.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'楷体_GB2312'))
        root = self.tree.AddRoot("")
        for item in DevInfo2:
            parentItem = self.tree.AppendItem(root, item[0])
            for ich in item[1]:
                tmpCh = self.tree.AppendItem(parentItem, ich)
                self.tree.SetItemFont(tmpCh, wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
                self.tree.SetItemTextColour(tmpCh, wx.BLUE)
        
        btn_Close = wx.Button(panel, -1, u"关闭", (400, 280))
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(topsizer, 0, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(self.tree, 1, wx.ALL|wx.EXPAND| wx.ALIGN_CENTER_HORIZONTAL,5)
        mainSizer.Add(wx.StaticLine(panel), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        mainSizer.Add(btn_Close, 0, wx.ALL|wx.BOTTOM| wx.ALIGN_RIGHT, 5) 
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)
        
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)        
    
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()
    
if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmAbout() 
    frame.Show()
    app.MainLoop()
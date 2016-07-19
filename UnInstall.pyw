#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import images
from PhrResource import strVersion
import os

class FrmAbout(wx.Frame): 
    def __init__(self): 
        title = u"卸载　" + strVersion
        meWidth = 450
        meHeight = 350
        wx.Frame.__init__(self, None, -1, title, size=(meWidth, meHeight),\
            style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER ^ wx.MAXIMIZE_BOX ^ wx.MINIMIZE_BOX)
        self.Center()
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
#        self.SetTransparent(240)

        panel = wx.Panel(self)
        titleText = wx.StaticText(panel, -1, title)
        titleText.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        topsizer = wx.BoxSizer(wx.HORIZONTAL)
#        topsizer.Add(imgLogo, 0, wx.ALL)
#        topsizer.Add((-1,-1), 1, wx.ALL| wx.EXPAND)
        topsizer.Add(titleText, 0, wx.ALL| wx.ALIGN_CENTER_VERTICAL, 5)
        self.text = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE)
        self.text.SetEditable(False)
        

        btn_UnIns = wx.Button(panel, -1, u"卸载")
        btn_Close = wx.Button(panel, -1, u"关闭")
        
        btnsizer = wx.BoxSizer( wx.HORIZONTAL)
        btnsizer.Add(btn_UnIns, 0, wx.ALL| wx.ALIGN_RIGHT, 5)
        btnsizer.Add(btn_Close, 0, wx.ALL| wx.ALIGN_RIGHT, 5)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(topsizer, 0, wx.ALL| wx.EXPAND, 5)

        mainSizer.Add(self.text, 1, wx.ALL| wx.EXPAND, 5)
        mainSizer.Add(btnsizer, 0, wx.ALL|wx.BOTTOM| wx.ALIGN_RIGHT, 5) 
        
        panel.SetSizer(mainSizer) 
        mainSizer.Fit(panel)
        
        self.Bind(wx.EVT_BUTTON, self.OnUnins, btn_UnIns)
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)        
    
    def OnUnins(self, event):
        txtDes = ""
        for ipath in os.listdir( os.getcwd() ):
            if ipath != "HolidayData":
                txtDes += ipath + "\r\n"
        
        self.text.SetLabel(txtDes)
                
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()
    
if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmAbout() 
    frame.Show()
    app.MainLoop()
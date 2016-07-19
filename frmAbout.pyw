#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import images

from PhrResource import strVersion, DevInfo, CopyRightInfo,AboutIMG

class FrmAbout(wx.Dialog): 
    def __init__(self): 
        title = strVersion
        meWidth = 500
        meHeight = 320
        wx.Dialog.__init__(self, None, -1, title, size=(meWidth, meHeight), \
            style= wx.DEFAULT_DIALOG_STYLE ^ wx.CAPTION)
        self.Center()
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
        self.SetTransparent(240)
        
        self.g_ScrollNum = 0
        
        img = wx.Image(AboutIMG, wx.BITMAP_TYPE_ANY)
        panel = wx.StaticBitmap(self, -1, img.ConvertToBitmap())

        strdes = ""
        for item in DevInfo:
            strdes += item  + u"\n\n"
            
        strdes = strdes[:-1]
        self.desTxt = wx.StaticText(panel, -1, strdes, (370, 100), style=wx.ALIGN_LEFT) 
        self.desTxt.SetBackgroundColour((238, 232, 170)) 
        self.desTxt.SetForegroundColour((25,25,112))
        self.desTxt.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体_GB2312'))

        infoC = wx.StaticText(panel, -1, CopyRightInfo, (300, 220), style=wx.ALIGN_LEFT) 
        infoC.SetBackgroundColour((151, 203, 244)) 
        
        btn_Close = wx.Button(panel, -1, u"关闭", (400, 280))
        
        self.timer = wx.PyTimer(self.ScrollTxt)
        self.timer.Start(1000)
        self.ScrollTxt()
        
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)        
    
    def ScrollTxt(self):
        linenum = len(DevInfo)
        if self.g_ScrollNum == linenum:
            self.g_ScrollNum = 0
            
        self.g_ScrollNum += 1
        
        strdes = ""
        for index in range(0, 3):
            strdes += DevInfo[(self.g_ScrollNum + index) % linenum] + "\n\n"
            
        strdes = strdes[:-2]
        self.desTxt.SetLabel(strdes)
                   
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()
        
if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmAbout() 
    frame.ShowModal()
    app.MainLoop()
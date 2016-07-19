#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import images
import time

from PhrResource import strVersion, DevInfo, CopyRightInfo,AboutIMG

class FrmAbout(wx.Dialog): 
    def __init__(self): 
        title = strVersion
        meWidth = 500
        meHeight = 350
        wx.Dialog.__init__(self, None, -1, title, size=(meWidth, meHeight), \
            style= wx.DEFAULT_DIALOG_STYLE ^ wx.CAPTION)
        self.Center()
#        icon=images.getProblemIcon()        
#        self.SetIcon(icon)
#        self.SetTransparent(240)
        
#        self.g_ScrollNum = 0
        
#        img = wx.Image(AboutIMG, wx.BITMAP_TYPE_ANY)
#        panel = wx.StaticBitmap(self, -1, img.ConvertToBitmap())

#        strdes = ""
#        for item in DevInfo:
#            strdes += item  + u"\n\n"
            
#        strdes = strdes[:-1]
#        self.desTxt = wx.StaticText(panel, -1, strdes, (370, 100), style=wx.ALIGN_LEFT) 
#        self.desTxt.SetBackgroundColour((238, 232, 170)) 
#        self.desTxt.SetForegroundColour((25,25,112))
#        self.desTxt.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体_GB2312'))

#        infoC = wx.StaticText(panel, -1, CopyRightInfo, (300, 220), style=wx.ALIGN_LEFT) 
#        infoC.SetBackgroundColour((151, 203, 244)) 
        
#        btn_Close = wx.Button(panel, -1, u"关闭", (400, 280))
        btn_Close = wx.Button(self, -1, u"关闭", (400, 300))
        
#        mainSizer = wx.BoxSizer(wx.VERTICAL)
        self.timer = wx.PyTimer(self.ScrollTxt)
        self.timer.Start(200)
        self.g_top = 260
#        self.ScrollTxt()
        
        self.Bind(wx.EVT_PAINT, self.OnPaint)
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)        
    
    def ScrollTxt(self):
        self.Refresh(False, [300, 80, 180, 180])
        if self.g_top == 80:
            self.g_top = 260
            
        self.g_top -= 1
    
    def OnPaint(self, evt):
        ns = self.GetSize()
        dc = wx.PaintDC(self)
        try:
            gc = wx.GraphicsContext.Create(dc)
        except:
            dc.DrawText("This build of wxPython does not support the wx.GraphicsContext "
                        "family of classes.",
                        25, 25)
            return

        path = gc.CreatePath()
        path.AddRectangle(0, 0, ns[0], ns[1])

        r, g, b = (35, 142,  35)
        penclr   = wx.Colour(r, g, b, wx.ALPHA_OPAQUE)
        brushclr = wx.Colour(r, g, b, 128)   # half transparent
        gc.SetPen(wx.Pen(penclr))
        gc.SetBrush(wx.Brush(brushclr))
        gc.FillPath(path)
        gc.PushState()
#        gc.Translate(60, 75)
        
        gc.PopState()              # restore saved state
        gc.PushState()

        gc.SetFont(wx.Font(22, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'黑体'))
        gc.DrawText(strVersion, 10, 20)
                
        gc.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体'))
        
        gc.DrawText(strVersion, 300, self.g_top-12)
        gc.DrawText(strVersion, 300, self.g_top-24-10)
            
#        bmp = wx.Bitmap('bitmap\\mainCenter.png')
#        bsz = bmp.GetSize()
#        gc.DrawBitmap(bmp, 10, 50, bsz.width, bsz.height)
        
        gc.PopState()
            
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.timer.Destroy()
        self.Destroy()
        
if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmAbout() 
    frame.ShowModal()
    app.MainLoop()
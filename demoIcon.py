#!/usr/bin/env python
#coding=utf8

import wx
from wx.tools import img2py
import images

class InfoMathDemo(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, -1, title, size = (970, 720),
                          style=wx.DEFAULT_FRAME_STYLE)
        icon=images.getProblemIcon()
        self.SetIcon(icon)
#        icon=img.getpythonImage()
        
#        self.SetIcon(icon)
        

class MySplashScreen (wx.SplashScreen):
    def __init__(self):
        bmp = wx.Image("bitmap/splash.bmp").ConvertToBitmap()
        
        wx.SplashScreen.__init__(self, bmp,
                                 wx.SPLASH_CENTRE_ON_SCREEN|wx.SPLASH_TIMEOUT,
                                 2500, None, -1)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        # 过 1000 才显示主窗口
        self.fc=wx.FutureCall(1000, self.ShowMain)

    def OnClose(self, evt):
        # evt.Skip() 的含义有待研究
        evt.Skip()
        self.Hide()

        # 如果还没到 2000(比如用户用鼠标点了 splash 图，引发 splash 图关闭)，
        # 则显示主窗口
        if self.fc.IsRunning():
            self.fc.Stop()
            self.ShowMain()

    def ShowMain(self):
        # 构造主窗口，并显示
        frame=InfoMathDemo(None, "InfoMathDemo")
        frame.Show()


class MyApp(wx.App):
    def OnInit(self):
        print "App init"
        self.SetAppName("InfoMath Demo")
        splash=MySplashScreen()
        return True

    def OnExit(self):
        print "Existing"

def GenIcon():
    command_lines = ["-u -i -n Problem bitmap\\PHRLogo.ico images.py"]
    
    for line in command_lines:
        args = line.split()
        img2py.main(args)
        
if __name__ == '__main__':
#    app=MyApp(False)
#    app.MainLoop()
    GenIcon()

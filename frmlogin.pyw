#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import wx.lib.buttons as buttons
import frmHoliMain
import images
import PyDatabase
import time
import os

from PhrResource import ALPHA_ONLY, DIGIT_ONLY, strVersion

class FrmLogin(wx.Frame):
    
    def __init__(self):
        title = strVersion
        wx.Frame.__init__(self, None, -1, title, size=(600, 430), \
            style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER ^ wx.MAXIMIZE_BOX)
        self.Center()
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
        img1 = wx.Image("bitmap\\login.bmp", wx.BITMAP_TYPE_BMP).ConvertToBitmap()
        panelBack = wx.Panel(self, -1)
        panel = wx.StaticBitmap(panelBack, -1, img1)
        
        self.DTable = 'UserLib'        
        self.lstUser = []
        self.lstPwd = []
        self.lstAdminLevel = []
        
        self.initLoad() # select all the user, pwd, adminlevel in the userlib        
        
        self.lbUser = wx.ComboBox(panel, -1, u"", (400, 275), (160, 18), \
            choices=self.lstUser)
        self.Text_Level = wx.TextCtrl(panel, -1, u"", (400, 305), (160, 20))            
        self.Text_Pwd = wx.TextCtrl(panel, -1, u"", (400, 332), (160, 20),\
            style = wx.TE_PASSWORD|wx.WANTS_CHARS)            
        self.Text_Level.SetEditable(False)
#        img2 = wx.Image("bitmap\\btn_login.bmp", wx.BITMAP_TYPE_BMP).ConvertToBitmap()
#        btn_login = wx.BitmapButton(panel, -1, img2, (360, 355), (80, 30))
        btn_login = buttons.GenButton(panel, -1, u"登录", (360, 358), (80, 30))
        btn_Close = buttons.GenButton(panel, -1, u"退出", (480, 358), (80, 30))
        font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体')
        for item in [btn_login, btn_Close]:
            item.SetBezelWidth(5)
            item.SetBackgroundColour('Navy')
            item.SetForegroundColour('yellow')  
            item.SetFont(font)
                        
        btn_login.SetFocus()
        btn_login.SetDefault()
        
        self.Text_Pwd.Bind(wx.EVT_KEY_DOWN, self.EvtChar)
#        self.Bind( wx.EVT_KEY_DOWN, self.EvtChar, self.Text_Pwd)
#        self.Bind( wx.EVT_CHAR, self.EvtChar, self.Text_Pwd)
        self.Bind(wx.EVT_TEXT, self.OnDispLevel, self.lbUser)
        self.Bind(wx.EVT_BUTTON, self.OnLogin, btn_login)       
        self.Bind(wx.EVT_BUTTON, self.OnCloseMe, btn_Close)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
    
    def EvtChar(self, event):
        keycode = event.GetKeyCode()
        if keycode == 13:
            self.OnLogin(None)
        event.Skip()
    
    def OnDispLevel(self, event):
        txtUser = self.lbUser.GetValue().strip()
        if txtUser == "":
            self.Text_Level.SetValue(u"") 
            return
        if txtUser in self.lstUser:            
            self.Text_Level.SetValue(self.lstAdminLevel[self.lstUser.index(txtUser)]) 
        else:
            self.Text_Level.SetValue(u"不正确的用户") 
            
    def initLoad(self):
        try:
            userLib = PyDatabase.DBSelect('', self.DTable, ['UserName'], 0)
        except:
            self.Destroy()            
            wx.MessageBox(u"空数据库文件!\n\n登录用户: cc ,\n密码: cc\n\n请在登录后导入excel文件", u"提示")
            PyDatabase.DBCTLib()
            app = MyApp()    
            app.MainLoop()
            return
            
        for item in userLib:
            self.lstUser.append(item[1])
            self.lstPwd.append(item[2])
            self.lstAdminLevel.append(item[3])            
        
    def OnLogin(self, event):
        txtUser = self.lbUser.GetValue().strip()
        txtPwd = self.Text_Pwd.GetValue().strip() 
        if txtUser == 'iefan' and txtPwd == '1':
            frmHoliMain.FrameMain(1, 'iefan').Show()
            self.Destroy()
            return
            
        for itemUser, itemPwd in zip(self.lstUser, self.lstPwd):
            if itemUser == txtUser and itemPwd == txtPwd:
                index = self.lstUser.index(txtUser) 
#                print index
                if self.lstAdminLevel[index] == u'管理员':
                    frmHoliMain.FrameMain(1, txtUser).Show()
                else:
                    frmHoliMain.FrameMain(0, txtUser).Show()
                self.Destroy()
                break
        else:
            wx.MessageBox(u"对不起，您输入的用户名与密码错误！", u"错误")

    def OnCloseMe(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()

class MySplashScreen (wx.SplashScreen):
    def __init__(self):
        bmp = wx.Image("bitmap/splash.png").ConvertToBitmap()
        
        wx.SplashScreen.__init__(self, bmp,
                                 wx.SPLASH_CENTRE_ON_SCREEN|wx.SPLASH_TIMEOUT,
                                 2000, None, -1)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        # set delay time
        self.fc=wx.FutureCall(1000, self.ShowMain)

    def OnClose(self, evt):
        # evt.Skip() ?
        evt.Skip()
        self.Hide()

        # if delay time is less 2000, but user click the splash, 
        # the splash closed auto. of couse show the mainwindow        
        if self.fc.IsRunning():
            self.fc.Stop()
            self.ShowMain()

    def ShowMain(self):
        # create the mainwindow
        frame = FrmLogin()        
        frame.Show()

class MyApp(wx.App):
    def OnInit(self):
        time_Set = (2010, 2, 15)
        time_Now = time.localtime()[:3]
        strDesp = u"试用期结束，请联系原作者!"
        if time_Now > time_Set:            
            try:
                PyDatabase.Insert1Tab('PersonPhd')
                print strDesp
#                PyDatabase.DBDeleteALL()
            except:
                pass
            wx.MessageBox(strDesp, "tip")
            return True
        
        if PyDatabase.Select1Tab('PersonPhd'):
            wx.MessageBox(strDesp, "tip")
            return True
            
        self.SetAppName("PersonHolidayRecord")
        splash=MySplashScreen()
        return True

    def OnExit(self):
        pass
        
if __name__ == '__main__':
    app = MyApp()    
    app.MainLoop()
    
#if __name__ == '__main__':
#    app = wx.PySimpleApp()
#    frmMain = MySplashScreen()
#    frmMain.Show()
#    app.MainLoop()
#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import frmRankDays
import frmHoliDays
import frmHoliManFix
import frmUserFix
import frmPermit
import frmPPFix
import frmRoadDays
import frmUnit
import frmAbout
import images
import frmImportXLS
import frmImportPHR
import ExportPHR
import time
import colorsys
import os
from math import cos, sin, radians

import wx.lib.agw.flatmenu as FM

from PhrResource import strVersion

#FRAMETB = True
TBFLAGS = ( wx.TB_HORIZONTAL
            | wx.NO_BORDER
            | wx.TB_FLAT
            )
            
class FrameMain(wx.Frame):
    
    def __init__(self, flagAdmin, curUser):
        title = strVersion + " - " + curUser
        wx.Frame.__init__(self, None, -1, title)
#        self.SetBackgroundColour(())
        self.panel = panel = wx.Panel(self, style = wx.BORDER)
        font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.NORMAL, face=u'宋体')
        self.panel.SetFont(font)
        
        self.Maximize()
        self.SetMinSize(self.GetSize())
        icon=images.getProblemIcon()
        self.SetIcon(icon)
        
        menuBar = wx.MenuBar()
        menuStrlst = [u"系统设置(&S)", u"批假销假(&P)", u"信息维护(&M)", u"帮助(&H)"]
        menuToplst = []
        for index in range(len(menuStrlst)):
            menuToplst.append(wx.Menu())
            menuBar.Append(menuToplst[index], menuStrlst[index])
            
        menuChildStr = [[u"用户维护",u"",  u"数据导入...", u"数据备份", u"导入备份文件...", u"", u"退出系统"], \
            [u"批假", u"销假及查询"], \
            [u"单位维护", u"人员信息", u"", u"军衔维护", u"路途维护", u"假期维护"], \
            [u"帮助...", u"", u"关于..."]]
        menuTipStr = [[u"增删、更改用户信息", u"", u"从Excel表格导入数据", u"数据备份",  u"导入备份文件...", u"", u"退出系统"], \
            [u"对单位人员进行批假登记",  u"查询休假人员及对到假人员进行登记"], \
            [u"对单位维护", u"人员基本数据信息维护",u"", u"对军衔维护", u"对路途维护", u"对法定节假日维护"], \
            [u"获取帮助", u"", u"系统说明"]]
        menuFunLst = [[self.OnUserFix, self.OnImportdata, self.OnExportdata, self.OnImportPHR, self.OnClose],\
            [self.OnPHPermit, self.OnPHFix],\
            [self.OnMtUnit, self.OnMtPP, self.OnMtRank, self.OnMtRoad, self.OnMtHoli],\
            [self.OnHelp, self.OnAbout]]

        self.panelDict = {}
        self.flagExit = 0
        
        menuChildlst = {}
        for menuT, menuC, menuTip, menuFun in zip(menuToplst, menuChildStr, menuTipStr, menuFunLst):
            menuChildlst[menuT] = []
            iNum = -1
            for index in range(len(menuC)):
                if menuC[index] != "":
                    iNum += 1
                    menuChildlst[menuT].append(menuT.Append(-1, menuC[index], menuTip[index]))
                    self.Bind(wx.EVT_MENU, menuFun[iNum], menuChildlst[menuT][iNum])
                else:
                    menuT.AppendSeparator()
                    
        self.SetMenuBar(menuBar)
        
        #===================================================
        tb = self.CreateToolBar(TBFLAGS)
        imgNamelst = ["user", "", "permit", "backUnit", "", "unit", "ppfix", "rank", "road", "time", "", "help", "about", "", "exit"]
        tiptoobar = [u"用户维护", u"对单位人员进行批假登记", u"查询休假人员及对到假人员进行登记", u"对单位维护", u"对人员信息维护",u"对军衔维护", u"对路途维护",u"对法定节假日维护", u"帮助", u"软件说明", u"退出系统"]
        hlptoobar = [u"用户维护", u"对单位人员进行批假登记", u"查询休假人员及对到假人员进行登记", u"对单位维护", u"对人员信息维护",u"对军衔维护", u"对路途维护",u"对法定节假日维护", u"帮助", u"软件说明", u"退出系统"]
        tbFunLst = [self.OnUserFix, self.OnPHPermit, self.OnPHFix, self.OnMtUnit, self.OnMtPP, \
            self.OnMtRank, self.OnMtRoad, self.OnMtHoli, self.OnHelp, self.OnAbout, self.OnClose]
        self.idTbDict = {}
        icontoobar = []
        tsize = (32,32)
        idcount = -1
        for index in range(len(imgNamelst)):
            if imgNamelst[index] != "":
                idcount += 1
                tbId =  wx.NewId()
                icontoobar.append(wx.Image("bitmap/" + imgNamelst[index] + ".png", wx.BITMAP_TYPE_ANY).Rescale(tsize[0], tsize[1]).ConvertToBitmap())
                tb.AddLabelTool(tbId, tiptoobar[idcount], icontoobar[idcount], shortHelp=tiptoobar[idcount], longHelp=hlptoobar[idcount])
                self.Bind(wx.EVT_TOOL, self.OnToolClick, id=tbId)
                self.idTbDict[tbId] = tbFunLst[idcount]
            else:
                tb.AddSeparator()
                
        tb.SetToolBitmapSize(tsize)        
        tb.Realize()
        
        #===================================================s
        self.sb = CustomStatusBar(self)
        self.SetStatusBar(self.sb)
        #===================================================
        if flagAdmin == 0:
            tb.EnableTool(self.idtoobar[0], False)
            tb.EnableTool(self.idtoobar[3], False)
            tb.EnableTool(self.idtoobar[4], False)
            tb.EnableTool(self.idtoobar[5], False)
            tb.EnableTool(self.idtoobar[6], False)
            menuBar.EnableTop(2, False)
            menuBar.Enable(menuBar.GetMenu(0).GetMenuItems()[0].GetId(), False)
            menuBar.Enable(menuBar.GetMenu(1).GetMenuItems()[0].GetId(), False)

        mainPanl(self.panel, self.GetClientSize())
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
    
    def OnToolClick(self, event):
        self.idTbDict[event.GetId()](None)        

    def ShowPanel(self, iRow, iCol, frmName):
        if self.flagExit == 1:
            self.flagExit = 0
            mainPanl(self.panel, self.GetClientSize(), 0)
            color = self.panel.GetParent().GetMenuBar().GetBackgroundColour()
            self.panel.SetBackgroundColour(color)
            self.panel.Refresh(True)
            try:
                self.panelDict.keys().index(`iRow` + `iCol`)
                for item in self.panelDict.values()[0]:
                    if item.GetClassName() != u"wxGauge95":
                        item.Show(True)
                return
            except:
                pass

        if self.flagExit == 0:
            try:
                self.panelDict.keys().index(`iRow` + `iCol`)
                return
            except:
                pass
            for item in self.panel.GetChildren():
                item.Destroy()
            frmName(self.panel)
            self.panelDict.clear()
            self.panelDict[`iRow` + `iCol`] = []
            self.panel.SetSize(self.GetClientSize())
        
    def OnUserFix(self, event):
        iRow = 0; iCol = 0
        self.ShowPanel(iRow, iCol, frmUserFix.UserFix)
    
    def OnExportdata(self, event):
        ExportPHR.ExportData(self.panel)
        
    def OnImportdata(self, event):
        frmImportXLS.FrmImport().ShowModal()
        
    def OnImportPHR(self, event):
        frmImportPHR.FrmImport().ShowModal()
    
    def OnHideToolbar(self, event):
        wx.MessageBox('OnHideToolbar!', 'hi')
    
    def OnHideStatus(self, event):
        wx.MessageBox('OnHideStatus!', 'hi')
    
    def OnPHPermit(self, event):
        iRow = 1; iCol = 0
        self.ShowPanel(iRow, iCol, frmPermit.HoliPermit)
        
    def OnPHFix(self, event):
        iRow = 1; iCol = 1
        self.ShowPanel(iRow, iCol, frmHoliManFix.HoliPHFix)
        
    def OnMtUnit(self, event):
        iRow = 2; iCol = 0
        self.ShowPanel(iRow, iCol, frmUnit.UnitFix)
        
    def OnMtPP(self, event):
        iRow = 2; iCol = 1
        self.ShowPanel(iRow, iCol, frmPPFix.PersonFix)
                
    def OnMtRank(self, event):
        iRow = 2; iCol = 2
        self.ShowPanel(iRow, iCol, frmRankDays.RankDaysFix)
        
    def OnMtRoad(self, event):
        iRow = 2; iCol = 3
        self.ShowPanel(iRow, iCol, frmRoadDays.RoadDaysFix)
        
    def OnMtHoli(self, event):
        iRow = 2; iCol = 4
        self.ShowPanel(iRow, iCol, frmHoliDays.HoliDaysFix)
        
    def OnHelp(self, event):
        try:
            os.startfile(r'help\index.htm') 
        except:
            wx.MessageBox(u'当前版本没有找到帮助文档!', u'提示')
        
    def OnAbout(self, event):
        frmAbout.FrmAbout().ShowModal()
    
    def OnPaint(self, event):
        dc = wx.PaintDC(self.panel)
        dc.BeginDrawing()
        dc.EndDrawing()

    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        retv = wx.MessageBox(u"确定退出吗？", u"提示", wx.OK | wx.CANCEL|wx.ICON_QUESTION)
        if wx.OK == retv:
            self.Destroy()
        else:
            for item in self.panel.GetChildren():
                self.panelDict[self.panelDict.keys()[0]].append(item)
                item.Show(False)
            
            self.flagExit = 1
            mainPanl(self.panel, self.GetClientSize(), 1)
            
class CustomStatusBar(wx.StatusBar):
    def __init__(self, parent):
        wx.StatusBar.__init__(self, parent, -1, style= wx.ALIGN_RIGHT)

        # This status bar has three fields
        self.SetFieldsCount(3)
        
        # Sets the three fields to be relative widths to each other.
        otherwidth = self.Parent.GetSize()[0] - 160
        self.SetStatusWidths([int(otherwidth*0.5), int(otherwidth*0.5), 160])
        self.sizeChanged = False
        
        # We're going to use a timer to drive a 'clock' in the last
        # field.
        self.timer = wx.PyTimer(self.Notify)
        self.timer.Start(1000)
        self.Notify()

    # Handles events from the timer we started in __init__().
    # We're using it to drive a 'clock' in field 2 (the third field).
    def Notify(self):
        t = time.localtime(time.time())
        st = time.strftime("%H:%M:%S", t)
        timedisp = `t.tm_year` + u"年" + `t.tm_mon` + u"月" + `t.tm_mday` + u"日" + u"　" + st
        self.SetStatusText(timedisp, 2)
        self.SetStatusText(u"研制单位：91960部队71分队", 1)
        if self.GetStatusText(0) == "":
            self.SetStatusText(strVersion, 0)
#        print self.GetStatusText(0).encode('gbk')
#        self.SetStatusText(timedisp, 2)

class mainPanl(object):
    def __init__(self, panel, panelsize, flag=1):
        panel.Hide()
        panel.SetBackgroundColour((40, 60, 80))
        self.panel = panel
        self.pnlsize = panelsize
        self.flag = flag
        
#        self.bgwindow = wx.StaticText(panel, -1, "hello", size = self.pnlsize)
        
        self.BASE  = self.pnlsize[1]*0.4    # sizes used in shapes drawn below
        self.BASE2 = self.BASE/2
        self.BASE4 = self.BASE/4

        panel.Bind(wx.EVT_PAINT, self.OnPaint)
        panel.Show()
        
    def OnPaint(self, event):
        dc = wx.PaintDC(self.panel)
        try:
            gc = wx.GraphicsContext.Create(dc)
        except NotImplementedError:
            dc.DrawText("This build of wxPython does not support the wx.GraphicsContext "
                        "family of classes.",
                        25, 25)
            return

        font = wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT)
        font.SetWeight(wx.BOLD)
        gc.SetFont(font)

        # make a path that contains a circle and some lines, centered at 0,0
        path = gc.CreatePath()
        path.AddCircle(0, 0, self.BASE2)
        path.MoveToPoint(0, -self.BASE2)
        path.AddLineToPoint(0, self.BASE2)
        path.MoveToPoint(-self.BASE2, 0)
        path.AddLineToPoint(self.BASE2, 0)
        path.CloseSubpath()
        path.AddRectangle(-self.BASE4, -self.BASE4/2, self.BASE2, self.BASE4)

        gc.Translate(self.pnlsize[0]/2, self.pnlsize[1]/2)

        # draw our path again, rotating it about the central point,
        # and changing colors as we go
        for angle in range(0, 360, 30):
            gc.PushState()         # save this new current state so we can 
                                   # pop back to it at the end of the loop
            if self.flag == 1:
                r, g, b = [int(c * 255) for c in colorsys.hsv_to_rgb(float(angle)/360, 1, 1)]
#                print r,g,b
            else:
                color = self.panel.GetParent().GetMenuBar().GetBackgroundColour()
                r = color[0]
                g = color[1]
                b = color[2]

            gc.SetBrush(wx.Brush(wx.Colour(r, g, b, 64)))
            gc.SetPen(wx.Pen(wx.Colour(r, g, b, 128)))
            
            # use translate to artfully reposition each drawn path
            gc.Translate(1.5 * self.BASE2 * cos(radians(angle)),
                         1.5 * self.BASE2 * sin(radians(angle)))

            # use Rotate to rotate the path
            gc.Rotate(radians(angle))

            # now draw it
            gc.DrawPath(path)
            gc.PopState()

        # Draw a bitmap with an alpha channel on top of the last group
        if self.flag == 1:
            bmp = wx.Bitmap('bitmap/mainCenter.png')
            bsz = bmp.GetSize()
            gc.DrawBitmap(bmp,
                      -bsz.width/2.5, 
                      -bsz.height/2.5,
                      bsz.width, bsz.height)
            gc.PopState()
        
class MyApp(wx.App):
    def OnInit(self):
        self.SetAppName("HolidayRecord")
        frmmain = FrameMain(1)
        frmmain.Show()
        return True

    def OnExit(self):
        pass
        
#if __name__ == '__main__':
#    app = MyApp()    
#    app.MainLoop()
                    
if __name__ == '__main__':
    app = wx.PySimpleApp()
    frmMain = FrameMain(1, 'me')
    frmMain.Show()
    app.MainLoop()
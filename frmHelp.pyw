#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import images

import ColorPanel
import imagesTb

from PhrResource import strVersion, DevInfo, CopyRightInfo,AboutIMG

class FrmHelp(wx.Frame): 
    def __init__(self): 
        title = strVersion
        wx.Frame.__init__(self, None, -1, title + u"　--　帮助")
        self.Center()
        self.SetMinSize((500, 600))
        icon=images.getProblemIcon()        
        self.SetIcon(icon)
        
        self.tb = TestTB(self, -1)
        self.tb.ExpandNode(0)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.tb, 1, wx.ALL| wx.EXPAND)
        
        self.SetSizer(mainSizer)
        mainSizer.Fit(self)
        self.Center()
        self.Maximize()
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
                   
    def OnClose(self, event):
        self.Close(True)
        
    def OnCloseWindow(self, event):
        self.Destroy()

colourList = [ u"军人休假登记系统"]

def getNextImageID(count):
    imID = 0
    while True:
        yield imID
        imID += 1
        if imID == count:
            imID = 0
            
class TestTB(wx.Treebook):
    def __init__(self, parent, id):
        wx.Treebook.__init__(self, parent, id, style=
                             wx.BK_DEFAULT
                             #wx.BK_TOP
                             #wx.BK_BOTTOM
                             #wx.BK_LEFT
                             #wx.BK_RIGHT
                            )
        
        # make an image list using the LBXX images
        il = wx.ImageList(16, 16)
        
        il.Add(getattr(imagesTb, 'LB02').GetBitmap())
        self.AssignImageList(il)
        self.imageIdGenerator = getNextImageID(il.GetImageCount())
        
        # Now make a bunch of panels for the list book
        win = self.makeColorPanel(colourList[0])
        self.AddPage(win, colourList[0], imageId=self.imageIdGenerator.next())
        st = wx.StaticText(win.win, -1,
                    "You can put nearly any type of window here,\n"
                    "and the wx.TreeCtrl can be on either side of the\n"
                    "Treebook",
                    wx.Point(10, 10))

        self.win = self.makeColorPanel(colourList[0])
        st = wx.StaticText(self.win.win, -1, "this is a sub-page", (10,10))
        self.InitChildTree()

        self.Bind(wx.EVT_TREEBOOK_PAGE_CHANGED, self.OnPageChanged)
        self.Bind(wx.EVT_TREEBOOK_PAGE_CHANGING, self.OnPageChanging)

        # This is a workaround for a sizing bug on Mac...
        wx.FutureCall(100, self.AdjustSize)
    
    def InitChildTree(self):
        self.AddSubPage(self.win, 'a sub-page', imageId=self.imageIdGenerator.next())
        self.AddSubPage(self.win, 'two-page', imageId=self.imageIdGenerator.next())
    
    def AdjustSize(self):
        #print self.GetTreeCtrl().GetBestSize()
        self.GetTreeCtrl().InvalidateBestSize()
        self.SendSizeEvent()
        #print self.GetTreeCtrl().GetBestSize()
        

    def makeColorPanel(self, color):
        p = wx.Panel(self, -1)
        win = ColorPanel.ColoredPanel(p, color)
        p.win = win
        def OnCPSize(evt, win=win):
            win.SetPosition((0,0))
            win.SetSize(evt.GetSize())
        p.Bind(wx.EVT_SIZE, OnCPSize)
        return p


    def OnPageChanged(self, event):
        old = event.GetOldSelection()
        new = event.GetSelection()
        sel = self.GetSelection()
        print 'OnPageChanged,  old:%d, new:%d, sel:%d\n' % (old, new, sel)
        event.Skip()

    def OnPageChanging(self, event):
        old = event.GetOldSelection()
        new = event.GetSelection()
        sel = self.GetSelection()
        print 'OnPageChanging, old:%d, new:%d, sel:%d\n' % (old, new, sel)
        event.Skip()

#----------------------------------------------------------------------------
if __name__ == '__main__': 
    app = wx.PySimpleApp() 
    frame = FrmHelp() 
    frame.Show()
    app.MainLoop()
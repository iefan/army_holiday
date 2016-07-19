#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import  time
import  thread

import  wx
import  wx.lib.newevent
import  ExportXls

#----------------------------------------------------------------------

# This creates a new Event class and a EVT binder function
(UpdateBarEvent, EVT_UPDATE_BARGRAPH) = wx.lib.newevent.NewEvent()
(ExportXlsEvent, EVT_EXPORT_XLS) = wx.lib.newevent.NewEvent()

class ExportXlsThread:
    def __init__(self, win, lststr, head):
        self.win = win
        self.lststr = lststr
        self.head = head

    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        time.sleep(0.5)
        ExportXls.ExportXls(self.lststr, self.head)
        time.sleep(0.5)
        evt = ExportXlsEvent(flag = True)
        wx.PostEvent(self.win, evt)
        self.running = False

class CalcBarThread:
    def __init__(self, win, count):
        self.win = win
        self.count = count

    def Start(self):
        self.keepGoing = self.running = True
        thread.start_new_thread(self.Run, ())

    def Stop(self):
        self.keepGoing = False

    def IsRunning(self):
        return self.running

    def Run(self):
        while self.keepGoing:
            evt = UpdateBarEvent(count = self.count)
            wx.PostEvent(self.win, evt)
                        
            if self.count >= 50:
                self.count = 0
                
            time.sleep(0.1)
            self.count = self.count + 1

        self.running = False

#----------------------------------------------------------------------

class TestPanel(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1)
        
        panel = wx.Panel(self, -1)

        b = wx.Button(panel, -1, "Create and Show a ProgressDialog", (50,100))
        self.Bind(wx.EVT_BUTTON, self.OnButton, b)
        self.Bind(EVT_UPDATE_BARGRAPH, self.OnUpdate)
        self.Bind(EVT_EXPORT_XLS, self.OnExport)
        
        self.thread = []
        self.thread.append(CalcBarThread(self, 0))
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        
        self.count = 0
        self.xlsFlag = False
        
        self.g1 = wx.Gauge(panel, -1, 50, (110, 50), (250, 25))
        print self.g1.GetClassName() == u"wxGauge95"
        self.g1.Show(False)
    
    def OnExport(self, evt):
        self.xlsFlag = evt.flag
        print self.xlsFlag
#        self.g1.Show(False)
        if self.xlsFlag:
            for item in self.thread:
                item.Stop()
                
            running = 1

            while running:
                running = 0
                for t in self.thread:
                    running = running + t.IsRunning()
                time.sleep(0.1)
        
    def OnUpdate(self, evt):
        self.g1.SetValue(evt.count)

    def OnButton(self, evt):
        
        lststr = [['1','aa','bb','cc'], ['2','a1','b','c1'],['1','aa','bb','cc'], ['2','a1','b','c1'],['1','aa','bb','cc'], ['2','a1','b','c1']]
        head = [u"编号", u"姓名", u"日期", u"天数"]
        self.thread.append(ExportXlsThread(self, lststr, head))
        
        self.g1.Show(True)
        
        for item in self.thread:
            item.Start()

    def OnCloseWindow(self, evt):
#        busy = wx.BusyInfo("One moment please, waiting for threads to die...")
#        wx.Yield()
        self.Destroy()


if __name__ == '__main__':
    app = wx.PySimpleApp()
    frmMain = TestPanel()
    frmMain.Show()
    app.MainLoop()
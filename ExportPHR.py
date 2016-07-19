#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import time
import wx
import sqlite3
import os

def ExportData(panel):
    if wx.CANCEL == wx.MessageBox(u"是否进行数据备份？", u"提示", wx.OK | wx.CANCEL):
        return
    dlg = wx.DirDialog(panel, u"请选择一个备份目录:", style=wx.DD_DEFAULT_STYLE)
    
    datapath = ""
    if dlg.ShowModal() == wx.ID_OK:
        datapath =  dlg.GetPath()
    dlg.Destroy()
    
    if datapath == "":
        wx.MessageBox(u"未选择备份目录，数据未能备份！", u"提示")
        return
            
    filename = "Phr" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + '.phr'
    if datapath != "":    
        filename = datapath + '\\' + filename
    
    strCmd = 'copy HolidayData "' + filename + '"'
    if os.system(strCmd.encode('gbk')) == 0:
        wx.MessageBox(u"备份成功！", u"提示")
    else:
        wx.MessageBox(u"未能成功备份，请检查所选择的备份文件路径！", u"提示")

if __name__ == '__main__':
    ExportData()
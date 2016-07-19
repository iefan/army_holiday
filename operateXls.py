#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import win32com
from win32com.client import Dispatch, constants
import time
import os
import wx
import string

def ExportXls(lstStr, head = [], xlspath = ""):
    try:
        w = win32com.client.Dispatch('Excel.Application')
    except:
        wx.MessageBox(u"本机没有安装 MS EXCEL，或者 MS EXCEL 未被正确安装！请检查！", u"提示")
        return
    
    w.Visible = 0
    w.DisplayAlerts = 0
    wb = w.Workbooks.Add()
    sh = wb.Sheets('Sheet1')
    sh.Name = u"人员信息"
    
    cols = string.ascii_uppercase
    
    if head != []:
        for index in range(len(head)):
            sh.Range(cols[index]+"1").FormulaR1C1 = head[index]
    
    for irow in range(len(lstStr)):
        for icol in range(len(lstStr[irow])):
            sh.Range(cols[icol]+`irow+2`).FormulaR1C1 = lstStr[irow][icol]

    sh.Rows("1:1").Select()
    w.Selection.Font.Name = u"宋体"
    w.Selection.Font.Size = 20
    w.Cells.Select()
    w.Selection.Columns.AutoFit()
    sh.Range("A1").Select()
    
    filename = "Phr" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + '.xls'
    if xlspath != "":    
        filename = xlspath + '\\' + filename
    else:
        filename = os.getcwd() + '\\' + filename
    
    w.Visible = 1
    w.ActiveWorkbook.SaveAs(filename)
        
if __name__ == '__main__':
    lststr = [['1','aa','bb','cc'], ['2','a1','b','c1']]
    head = [u"编号", u"姓名", u"日期", u"天数"]
    ExportXls(lststr, head)
#    GenWordList([u"BZDD000220090804", u"关博", u"综合处", u"2009正常假30天，晚婚假 10 天，路途（湖南）20天，共计60天。", u"2009-08-04", u"2009-09-10"])
#    print os.getcwd()
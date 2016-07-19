#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import pyExcelerator as pye
import time

def ExportXls(lstStr, head = [], xlspath = ""):
    wb = pye.Workbook()
    ws = wb.add_sheet('result')
    if head != []:
        for index in range(len(head)):
            ws.write(0, index, head[index])
    
    for irow in range(len(lstStr)):
        for icol in range(len(lstStr[irow])):
            ws.write(irow+1, icol, lstStr[irow][icol])
    
    filename = "Phr" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + '.xls'
    
    if xlspath != "":    
        filename = xlspath + '\\' + filename
    
    wb.save(filename)
    
if __name__ == '__main__':
    lststr = [['1','aa','bb','cc'], ['2','a1','b','c1']]
    head = [u"编号", u"姓名", u"日期", u"天数"]
    
    ExportXls(lststr, head)
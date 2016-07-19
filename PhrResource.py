#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import wx
import  wx.lib.mixins.listctrl  as  listmix

strVersion = u"军人休假登记系统0.1.9.906" 
ALPHA_ONLY = 1
DIGIT_ONLY = 2

#Person Sn numbers
g_UnitSnNum = 8

# frmAbout.pyw : Development people infomation
DevInfo2 = [[u"版权所有", [u"91960部队71分队"]], \
    [u"总策划", [u"王剑红　王晓华"]],\
    [u"业务指导",[u"水警区军务科"]],\
    [u"信息提供及测试", [u"向崇文"]],\
    [u"程序设计", [u"潘湘飞"]], \
    [u"反馈邮箱", [u"mybsppp@gmail.com"]], \
    ]

DevInfo = [u"程序设计：　　　\n潘湘飞　",\
    u"运行环境：　　　\nWinXP+MSOffice",\
    u"信息提供及测试：\n向崇文　向述乐"]
CopyRightInfo = u"版权所有(c)　91960部队71分队\n" +\
        u"联系电话：0713-31679\n" + \
        u"试用期：2010.2.15\n" + \
        u"mybsppp@gmail.com"
        
AboutIMG = "bitmap/aboutBG.jpg"
        
class MyListCtrl(wx.ListCtrl, listmix.ListCtrlAutoWidthMixin):
    def __init__(self, parent, ID, pos=wx.DefaultPosition,
                 size=wx.DefaultSize, style=0):
        wx.ListCtrl.__init__(self, parent, ID, pos, size, style)
        listmix.ListCtrlAutoWidthMixin.__init__(self)

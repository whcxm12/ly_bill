#！/user/bin/env python3
# -*- coding: utf-8 -*-
# @Time     : 2018/11/24 13:36
# @Author   : cxm
# @Site     : 
# @File     : bill_gui.py
# @Software : PyCharm

import tkinter
import xlrd
import xlwt

#绘制gui界面

#绘制主界面form
wdow=tkinter.Tk()

#标签，显示数据文件存储的路径
lbl_path=tkinter.Label(wdow,width=50)
#浏览按钮，用于选择数据文件的存储路径
but_brow=tkinter.Button(wdow,width=20,hight=2)
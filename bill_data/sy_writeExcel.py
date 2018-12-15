#！/user/bin/env python3
# -*- coding: utf-8 -*-
# @Time     : 2018/12/13 8:27
# @Author   : cxm
# @Site     : 
# @File     : sy_writeExcel.py
# @Software : PyCharm

import xlrd
import xlutils3.copy

inp=['ok1','ok2','ok3','ok4','ok5']
book_fil=xlrd.open_workbook('sy.xls')
sheet_fil=book_fil.sheet_by_index(0)
rows_old=sheet_fil.nrows
print(rows_old)
book_dat=xlutils3.copy.copy(book_fil)
# print(dir(book_dat))
# for i in dir(book_dat):
#     print(i)
sheet_dat=book_dat.get_sheet(0)
# print(dir(sheet_dat))
# for j in dir(sheet_dat):
#     print(j)
# print(help(sheet_dat.Row))
# sheet_dat.write(1,3,4)
# sheet_dat.write(1,0,'hello')
# sheet_dat.write(rows_old,0,'hello2')
for ii in range(len(inp)):
    sheet_dat.write(rows_old,ii,inp[ii])
book_dat.save('sy.xls')

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
import os
import tkinter.filedialog

#绘制gui界面

#绘制主界面form
wdow=tkinter.Tk()
#设置窗口标题
wdow.title("发票信息管理")
#设置窗口大小为宽600，长670，设置显示位置为左上角为屏幕的(300,0)位置
wdow.geometry("600x670+300+0")

'''用与选择数据文件的位置。程序开始时默认显示使用上次使用的文件位置。
可以点浏览按钮选择新的文件位置，并将新的文件位置保存到path_data中，以便于下次使用。'''

#函数实现为path_text赋值的功能
def get_path_data():
        path_text_path = os.path.abspath('.') + '\\bill_data\\path_data'
        fil = open(path_text_path, mode='r')
        return fil.read()
        fil.close()

# 存储路径的变量
path_text = tkinter.StringVar()
path_text.set(get_path_data())
#标签，显示数据文件存储的路径
lbl_path=tkinter.Label(wdow,
        textvariable=path_text,     #根据攻取的值，动态显示路径
        bg='#FFFFFF',
        font=('Arial',12),
        width=50,height=2).place(x=20,y=20,anchor='nw')
#浏览按钮，用于选择数据文件的存储路径

#以下的函数用于调用Windows系统的打开文件对话框，并保存到path_data文件中
def open_dialog():
        #ddir=
        fname=tkinter.filedialog.askopenfilename(title=u'选择文件',initialdir=os.path.expanduser(path_text.get()))
        fil=open(os.path.abspath('.')+'\\bill_data\\path_data',mode='w')
        fil.write(fname)
        fil.close()

but_brow=tkinter.Button(wdow,
        text='浏览',
        width=10,height=2,
        command=open_dialog).place(x=580,y=20,anchor='ne')

'''接收录入的发票信息，
当录入了发票代码和发票号码后，可以点击验证按钮，验证新数据与旧数据是否存在重复。'''

'''#要填写的字段：发票代码	发票号码	发票日期（date）	报销人	报销人部门
	发票金额	凭证月份*	凭证号*	状态(已检查)	提交日期（当前日期时间)
	带星号的为选填。'''

#录入发票号码和发票代码，并验证是否与前期数据重复，要求发票号码和发票号码均要填写。
#发票代码的标签
daima_bill_lbl=tkinter.Label(wdow,
                text="发票代码：",
                font=('Arial',12),
                width=14,height=2).place(x=20,y=70)
#发票代码的输入框
diama_bill_input=tkinter.Entry(wdow,
                font=('Arial',12),
                width=40).place(x=160,y=80)

#发票号码的标签
haoma_bill_lbl=tkinter.Label(wdow,
                text="发票号码：",
                font=('Arial', 12),
                width=14, height=2).place(x=20,y=100)
#发票号码的输入框
haoma_bill_input=tkinter.Entry(wdow,
                font=('Arial',12),
                width=40).place(x=160,y=110)

#发票日期的标签
date_bill_lbl=tkinter.Label(wdow,
                text="发票日期：",
                font=('Arial', 12),
                width=14, height=2).place(x=20,y=130)
#发票日期的输入框
date_bill_input=tkinter.Entry(wdow,
                font=('Arial',12),
                width=40).place(x=160,y=140)

tkinter.mainloop()

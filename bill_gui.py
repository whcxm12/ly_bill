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
import time
import xlutils3
import pypyodbc
import xlutils3.copy

#绘制gui界面

#绘制主界面form
wdow=tkinter.Tk()
#设置窗口标题
wdow.title("发票信息管理")
#设置窗口大小为宽600，长670，设置显示位置为左上角为屏幕的(300,0)位置
wdow.geometry("620x670+300+0")

'''用与选择数据文件的位置。程序开始时默认显示使用上次使用的文件位置。
可以点浏览按钮选择新的文件位置，并将新的文件位置保存到path_data中，以便于下次使用。'''

#函数实现为path_text赋值的功能
def get_path_data():
        path_text_path = os.path.abspath('.') + '\\bill_data\\path_data'
        fil = open(path_text_path, mode='r')
        return fil.read().strip()
        fil.close()

# 存储路径的变量
path_text = tkinter.StringVar()
path_text.set(get_path_data())
#标签，显示数据文件存储的路径
lbl_path=tkinter.Label(wdow,
        textvariable=path_text,     #根据攻取的值，动态显示路径
        bg='#FFFFFF',
        font=('Arial',12),
        width=38,height=2).place(x=20,y=23,anchor='nw')
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
        font=('Arial',12),
        command=open_dialog).place(x=580,y=20,anchor='ne')

#在窗口显示当前日期和时间
vartime=tkinter.StringVar()

def get_now():
    vartime.set(time.strftime("%Y-%m-%d %H:%M:%S",time.localtime()))
    wdow.after(1000,get_now)

time_now_lbl=tkinter.Label(wdow,
            textvariable=vartime,
            font=('Arial',12),
            width=22,height=2).place(x=20,y=90)
get_now()

'''接收录入的发票信息，
当录入了发票代码和发票号码后，可以点击验证按钮，验证新数据与旧数据是否存在重复。
如果不重复，则可以提交数据，如果重复，则提示数据有重复，不能提交数据。'''

'''#要填写的字段：发票代码	发票号码	发票日期（date）	报销人	报销人部门
	发票金额	凭证月份*	凭证号*	状态(已检查)	提交日期（当前日期时间)
	带星号的为选填。'''

#录入发票号码和发票代码，并验证是否与前期数据重复，要求发票号码和发票号码均要填写。
#发票代码的标签
wi=27
fix=60
daima_bill_lbl=tkinter.Label(wdow,
                text="发票代码：",
                font=('Arial',12),
                width=14,height=2).place(x=20,y=80+fix)
#发票代码的输入框
daima_bill=tkinter.StringVar()
daima_bill_input=tkinter.Entry(wdow,
                textvariable=daima_bill,
                font=('Arial',12),
                width=wi).place(x=180,y=90+fix)

#发票号码的标签
haoma_bill_lbl=tkinter.Label(wdow,
                text="发票号码：",
                font=('Arial', 12),
                width=14, height=2).place(x=20,y=120+fix)
#发票号码的输入框
haoma_bill=tkinter.StringVar()
haoma_bill_input=tkinter.Entry(wdow,
                font=('Arial',12),
                textvariable=haoma_bill,
                width=wi).place(x=180,y=130+fix)

#查重按钮，检查发票代码和发票号码是否同时重复。

norepeat_butt=tkinter.Button(wdow,
            text='检查重复',
            font=('Arial',12),
            width=8,height=1).place(x=180+wi+292,y=130+fix-6)

#发票日期的标签
date_bill_lbl=tkinter.Label(wdow,
                text="发票日期：",
                font=('Arial', 12),
                width=14, height=2).place(x=20,y=160+fix)
#发票日期的输入框
date_bill=tkinter.StringVar()
date_bill_input=tkinter.Entry(wdow,
                font=('Arial',12),
                textvariable=date_bill,
                width=wi).place(x=180,y=170+fix)

#报销人的标签
people_bill_lbl=tkinter.Label(wdow,
                            text="报销人：",
                            font=('Arial', 12),
                            width=14, height=2).place(x=20,y=200+fix)
#报销人的输入框
people_bill=tkinter.StringVar()
people_bill_input=tkinter.Entry(wdow,
                font=('Arial',12),
                textvariable=people_bill,
                width=wi).place(x=180,y=210+fix)

#部门的标签
department_bill_lbl=tkinter.Label(wdow,
                              text="部门：",
                              font=('Arial', 12),
                              width=14, height=2).place(x=20,y=240+fix)
#部门的输入框
department_bill=tkinter.StringVar()
department_bill_input=tkinter.Entry(wdow,
                                font=('Arial',12),
                                textvariable=department_bill,
                                width=wi).place(x=180,y=250+fix)

#发票金额的标签
money_bill_lbl=tkinter.Label(wdow,
                                  text="发票金额：",
                                  font=('Arial', 12),
                                  width=14, height=2).place(x=20,y=280+fix)
#发票金额的输入框
money_bill=tkinter.StringVar()
money_bill_input=tkinter.Entry(wdow,
                            font=('Arial',12),
                            textvariable=money_bill,
                            width=wi).place(x=180,y=290+fix)

#凭证月份的标签
month_voucher_lbl=tkinter.Label(wdow,
                                  text="凭证月份：",
                                  font=('Arial', 12),
                                  width=14, height=2).place(x=20,y=320+fix)
#凭证月份的输入框
month_voucher=tkinter.StringVar()
month_voucher_input=tkinter.Entry(wdow,
                                font=('Arial',12),
                                textvariable=month_voucher,
                                width=wi).place(x=180,y=330+fix)

#凭证号的标签
NO_voucher_lbl=tkinter.Label(wdow,
                                  text="凭证号：",
                                  font=('Arial', 12),
                                  width=14, height=2).place(x=20,y=360+fix)
#凭证号的输入框
NO_voucher=tkinter.StringVar()
NO_voucher_input=tkinter.Entry(wdow,
                            font=('Arial',12),
                            textvariable=NO_voucher,
                            width=wi).place(x=180,y=370+fix)

#发票状态标签
status_bill_lbl=tkinter.Label(wdow,
                text="状态：",
                font=('Arial',12),
                width=14,height=2).place(x=20,y=400+fix)
#发票状态输入框
status_bill=tkinter.StringVar()
status_bill_input=tkinter.Entry(wdow,
                            font=('Arial',12),
                            textvariable=status_bill,
                            width=wi).place(x=180,y=410+fix)

#函数save_bill的功能是将上述输入框中填写的数据保存到所选路径下的电子表格文件中。
def save_bill():
    #获取当前日期和时间并按照“yyyy-mm-dd hh:mm:ss”的格式保存
    save_time=time.strftime("%Y-%m-%d %H:%M:%S",time.localtime())

    #以列表的形式缓存要写入的数据
    bill_dat=[daima_bill.get().strip(),haoma_bill.get().strip(),date_bill.get().strip(),
              people_bill.get().strip(),department_bill.get().strip(),
              money_bill.get().strip(),month_voucher.get().strip(),NO_voucher.get().strip(),
              status_bill.get().strip(),save_time]
    # bill_dat=[save_time]
    #按照浏览的路径打开Excel文件
    excel_fil=xlrd.open_workbook(path_text.get())
    #获取文件中的所有表格
    # excel_fil_sheets=excel_fil.sheet_names()
    #获取第一个表格
    excel_fil_sheetNo=excel_fil.sheet_by_index(0)
    #获取表格中已存在的数据行数
    rows_old=excel_fil_sheetNo.nrows
    print(rows_old)

    #复制已打开文件中的数据
    excel_dat=xlutils3.copy.copy(excel_fil)
    #选择要写入的工作薄
    excel_dat_sheet=excel_dat.get_sheet(0)

    print(path_text.get())
    print(bill_dat)

    for col in range(len(bill_dat)):
        excel_dat_sheet.write(rows_old,col,bill_dat[col])

    excel_dat.save(path_text.get())



#提交信息按钮
commit_butt=tkinter.Button(wdow,
                        text='提交',
                        font=('Arial',12),
                        command=save_bill,
                        width=10,height=2).place(x=250,y=550)

tkinter.mainloop()

# ！/user/bin/env python3
# -*- coding: utf-8 -*-
# @Time     : 2018/11/24 13:36
# @Author   : cxm
# @Site     : 
# @File     : bill_gui.py
# @Software : PyCharm

import tkinter
import tkinter.messagebox
import xlrd
import xlwt
import os
import tkinter.filedialog
import time
import xlutils3
# import pypyodbc
import xlutils3.copy

# 绘制gui界面

# 绘制主界面form
wdow = tkinter.Tk()
# 设置窗口标题
wdow.title("发票信息管理")
# 设置窗口大小为宽650，长670，设置显示位置为左上角为屏幕的(300,0)位置
wdow.geometry("650x670+450+100")
# wdow.geometry("650x670")

'''用与选择数据文件的位置。程序开始时默认显示使用上次使用的文件位置。
可以点浏览按钮选择新的文件位置，并将新的文件位置保存到path_data中，以便于下次使用。'''

# 函数实现为path_text赋值的功能
# def get_path_data():
#     path_text_path = os.path.abspath('.') + '\\bill_data\\path_data'
#     fil = open(path_text_path, mode='r')
#     return fil.read().strip()
#     fil.close()


# 存储路径的变量
path_text = tkinter.StringVar()
# path_text.set(get_path_data())
# 标签，显示数据文件存储的路径
lbl_path = tkinter.Label(wdow,
                         textvariable=path_text,  # 根据攻取的值，动态显示路径
                         bg='#FFFFFF',
                         font=('Arial', 12),
                         width=38, height=2).place(x=20, y=23, anchor='nw')


# 浏览按钮，用于选择数据文件的存储路径

# 以下的函数用于调用Windows系统的打开文件对话框，并保存到path_data文件中
def open_dialog():
    # ddir=
    fname = tkinter.filedialog.askopenfilename(
        title=u'选择文件',
        # initialdir=os.path.expanduser(path_text.get())
    )
    print("数据文件的地址：", fname)
    print(len(fname))
    path_text.set(fname)
    print("数据文件地址参数path_text：",path_text.get())
    print(len(path_text.get()))
    # fil = open(os.path.abspath('.') + '\\bill_data\\path_data', mode='w')
    # fil.write(fname)
    # fil.close()


but_brow = tkinter.Button(wdow,
                          text='浏览',
                          width=6, height=1,
                          font=('Arial', 12),
                          command=open_dialog).place(x=450, y=30)

# 在窗口显示当前日期和时间
vartime = tkinter.StringVar()


def get_now():
    vartime.set(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    wdow.after(1000, get_now)


time_now_lbl = tkinter.Label(wdow,
                             textvariable=vartime,
                             # fg='#DC143C',
                             font=('Arial', 12),
                             width=22, height=2).place(x=400, y=500)
get_now()

'''接收录入的发票信息，
当录入了发票代码和发票号码后，可以点击验证按钮，验证新数据与旧数据是否存在重复。
如果不重复，则可以提交数据，如果重复，则提示数据有重复，不能提交数据。'''

'''#要填写的字段：发票代码	发票号码	发票日期（date）	报销人	报销人部门
	发票金额	凭证月份*	凭证号*	状态(已检查)	提交日期（当前日期时间)
	带星号的为选填。'''

# 录入发票号码和发票代码，并验证是否与前期数据重复，要求发票号码和发票号码均要填写。
# 发票代码的标签
wi = 27
fix = 60
daima_bill_lbl = tkinter.Label(wdow,
                               text="发票代码：",
                               font=('Arial', 12),
                               width=14, height=2).place(x=20, y=80 + fix)
# 发票代码的输入框
daima_bill = tkinter.StringVar()
daima_bill_input = tkinter.Entry(wdow,
                                 textvariable=daima_bill,
                                 font=('Arial', 12),
                                 width=wi).place(x=180, y=90 + fix)

# 发票号码的标签
haoma_bill_lbl = tkinter.Label(wdow,
                               text="发票号码：",
                               font=('Arial', 12),
                               width=14, height=2).place(x=20, y=120 + fix)
# 发票号码的输入框
haoma_bill = tkinter.StringVar()
haoma_bill_input = tkinter.Entry(wdow,
                                 font=('Arial', 12),
                                 textvariable=haoma_bill,
                                 width=wi).place(x=180, y=130 + fix)

# 验证结果显示标签
v_result = tkinter.StringVar()


# 验证发票代码和发票号码不能为空，只能为数字
def v_majiancha(daima, haoma):
    if ((not str(daima).isdigit()) and len(str(daima).strip()) != 12):
        # print("请正确填写发票代码！")
        v_result.set("请正确填写发票代码！")
        return False
    elif ((not str(haoma).isdigit()) and len(str(haoma).strip()) != 8):
        # print("请正确填写发票号码！")
        v_result.set("请正确填写发票号码！")
        return False
    else:
        return True


# 查重按钮，检查发票代码和发票号码是否同时重复。
# 对比发票代码和发票号码的函数
# Flag=True
def notrepeat(daima, haoma):
    if len(path_text.get())!=0:
        if v_majiancha(daima, haoma):
            # 发票代码和发票号码合并成一个字符串
            in_number = str(daima).strip() + str(haoma).strip()
            # 打开指定路径的数据文件，读取指定的工作表，获得表中已有数据的总行数。
            exfil = xlrd.open_workbook(path_text.get(), 'r')
            sheetfil = exfil.sheet_by_index(0)
            norows = sheetfil.nrows
            # print(norows)
            # 读取每行已有数据的发票代码和发票号码，并与输入的对比，如果相同则退出循环，并返回Ture；
            # 如果不相同则继续，直到全部数据对比完，并返回False。
            for r in range(norows):
                ed_num = str(sheetfil.cell(r, 0).value).strip() + str(sheetfil.cell(r, 1).value).strip()
                # print(ed_num)
                # print(in_number)
                if in_number == ed_num:
                    Flag = False
                    v_result.set("发票有重复！")

                    break

                else:
                    Flag = True
                    v_result.set("")

            if Flag:
                return True
            else:
                return False
        else:
            return False
    else:
        tkinter.messagebox.showwarning("注意","没有选择数据文件！")

#用于检查重复按钮调用查notreapet函数。并传送参数。
def notr():
    notrepeat(daima_bill.get(),haoma_bill.get())


# #根据查重函数的结果，显示不同的提示语
v_result_lbl = tkinter.Label(wdow,
                             textvariable=v_result,
                             font=('Arial', 12, 'bold'),
                             fg='#FF0000',
                             width=52, height=2).place(x=20, y=90)

norepeat_butt = tkinter.Button(wdow,
                               text='检查重复',
                               font=('Arial', 12),
                               command=notr,
                               width=8, height=1).place(x=180 + wi + 292, y=130 + fix - 6)

# 发票日期的标签
date_bill_lbl = tkinter.Label(wdow,
                              text="发票日期：",
                              font=('Arial', 12),
                              width=14, height=2).place(x=20, y=160 + fix)
# 发票日期的输入框
date_bill = tkinter.StringVar()
date_bill_input = tkinter.Entry(wdow,
                                font=('Arial', 12),
                                textvariable=date_bill,
                                width=wi).place(x=180, y=170 + fix)

# 报销人的标签
people_bill_lbl = tkinter.Label(wdow,
                                text="报销人：",
                                font=('Arial', 12),
                                width=14, height=2).place(x=20, y=200 + fix)
# 报销人的输入框
people_bill = tkinter.StringVar()
people_bill_input = tkinter.Entry(wdow,
                                  font=('Arial', 12),
                                  textvariable=people_bill,
                                  width=wi).place(x=180, y=210 + fix)

# 部门的标签
department_bill_lbl = tkinter.Label(wdow,
                                    text="部门：",
                                    font=('Arial', 12),
                                    width=14, height=2).place(x=20, y=240 + fix)
# 部门的输入框
department_bill = tkinter.StringVar()
department_bill_input = tkinter.Entry(wdow,
                                      font=('Arial', 12),
                                      textvariable=department_bill,
                                      width=wi).place(x=180, y=250 + fix)

# 发票金额的标签
money_bill_lbl = tkinter.Label(wdow,
                               text="发票金额：",
                               font=('Arial', 12),
                               width=14, height=2).place(x=20, y=280 + fix)
# 发票金额的输入框
money_bill = tkinter.StringVar()
money_bill_input = tkinter.Entry(wdow,
                                 font=('Arial', 12),
                                 textvariable=money_bill,
                                 width=wi).place(x=180, y=290 + fix)

# 凭证月份的标签
month_voucher_lbl = tkinter.Label(wdow,
                                  text="凭证月份：",
                                  font=('Arial', 12),
                                  width=14, height=2).place(x=20, y=320 + fix)
# 凭证月份的输入框
month_voucher = tkinter.StringVar()
month_voucher_input = tkinter.Entry(wdow,
                                    font=('Arial', 12),
                                    textvariable=month_voucher,
                                    width=wi).place(x=180, y=330 + fix)

# 凭证号的标签
NO_voucher_lbl = tkinter.Label(wdow,
                               text="凭证号：",
                               font=('Arial', 12),
                               width=14, height=2).place(x=20, y=360 + fix)
# 凭证号的输入框
NO_voucher = tkinter.StringVar()
NO_voucher_input = tkinter.Entry(wdow,
                                 font=('Arial', 12),
                                 textvariable=NO_voucher,
                                 width=wi).place(x=180, y=370 + fix)

# 发票状态标签
status_bill_lbl = tkinter.Label(wdow,
                                text="状态：",
                                font=('Arial', 12),
                                width=14, height=2).place(x=20, y=400 + fix)
# 发票状态输入框
status_bill = tkinter.StringVar()
status_bill_input = tkinter.Entry(wdow,
                                  font=('Arial', 12),
                                  textvariable=status_bill,
                                  width=wi).place(x=180, y=410 + fix)


# 必填项非空检查
def feikong(*lis):
    fk_cont = ['发票日期', '报销人', '报销部门', '发票金额', '状态']
    Ls = lis[2:6] + lis[8:9]  # 截取参数中的必填部分进行检查。
    # print("正在非空检查！")

    fk_res = True
    fk_n = []
    fk_str = ""
    for fk in range(len(Ls)):
        # print(fk)
        if len(Ls[fk]) == 0:
            fk_res = False
            fk_n.append(fk)
    # print(fk_n)
    if fk_res == False:
        for fk2 in fk_n:
            fk_str = fk_str + fk_cont[fk2] + "，"
        # print(fk_str+"内容不能为空！")
        v_result.set(fk_str + "内容不能为空！")
        return False
    else:
        return True


# 清空所有输入框
def qinkong():
    haoma_bill.set("")
    date_bill.set("")
    people_bill.set("")
    department_bill.set("")
    money_bill.set("")
    month_voucher.set("")
    NO_voucher.set("")
    status_bill.set("")
    daima_bill.set("")


# 函数save_bill的功能是将上述输入框中填写的数据保存到所选路径下的电子表格文件中。
def save_bill():
    # 获取当前日期和时间并按照“yyyy-mm-dd hh:mm:ss”的格式保存
    save_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

    # 以列表的形式缓存要写入的数据
    bill_dat = [daima_bill.get().strip(), haoma_bill.get().strip(), date_bill.get().strip(),
                people_bill.get().strip(), department_bill.get().strip(),
                money_bill.get().strip(), month_voucher.get().strip(), NO_voucher.get().strip(),
                status_bill.get().strip(), save_time]

    res = notrepeat(bill_dat[0], bill_dat[1])
    # print(res)
    if res:

        # 检查非空
        if feikong(*bill_dat):

            # 按照浏览的路径打开Excel文件
            excel_fil = xlrd.open_workbook(path_text.get())

            # 获取第一个表格
            excel_fil_sheetNo = excel_fil.sheet_by_index(0)
            # 获取表格中已存在的数据行数
            rows_old = excel_fil_sheetNo.nrows
            # print(rows_old)

            # 复制已打开文件中的数据
            excel_dat = xlutils3.copy.copy(excel_fil)
            # 选择要写入的工作薄
            excel_dat_sheet = excel_dat.get_sheet(0)

            # print(path_text.get())
            # print(bill_dat)

            for col in range(len(bill_dat)):
                excel_dat_sheet.write(rows_old, col, bill_dat[col])

            excel_dat.save(path_text.get())
            # print("已填写到表中！")
            tkinter.messagebox.showinfo('成功', '数据保存成功')
            qinkong()


# 提交信息按钮
commit_butt = tkinter.Button(wdow,
                             text='提交',
                             font=('Arial', 12),
                             command=save_bill,
                             width=10, height=2).place(x=180, y=550)

# 从Excel表中导入多条数据。并进行查重、非空的验证，提示发票重复及有空值的行号。
# 不重复和没有空值的数据导入到数据文件中。

'''
实现导入功能的函数

'''


def inport():
    try:
        # 要导入的数据的路径，通过打开文件对话框选择
        infil_path = tkinter.filedialog.askopenfilename(
            title=u'选择文件',
            filetypes=[("Excel file", "*.xls;*.xlsx")])
        print("文件路径为：" + infil_path)
        # 打开指定的文件
        infil = xlrd.open_workbook(infil_path)
        # 打开第一个工作表。要求所有的导入数据都放在第一个工作表中。
        infil_sheet = infil.sheet_by_index(0)
        # 获取导入数据的行数和列数
        inrows = infil_sheet.nrows
        incols = infil_sheet.ncols
        # 预计一个空的列表用于存放要导入的数据。
        indata = []
        # 将导入的数据按单元格读出来，存放到列表indata中
        for i in range(inrows):
            indatalist = []  # 用于存放每一行的数据。
            for j in range(incols):
                indatalist.append(str(infil_sheet.cell(i, j).value).strip())
            indata.append(indatalist)

        # 打开数据文件
        mofil = xlrd.open_workbook(path_text.get())
        mofil_sheet = mofil.sheet_by_index(0)
        # 获取数据文件的行数，以及第一行的列数
        morows = mofil_sheet.nrows
        mocols = mofil_sheet.row_len(0)
        # print(mocols)

        modata = xlutils3.copy.copy(mofil)
        modata_sheet = modata.get_sheet(0)

        errlist = []

        for ro in range(len(indata) - 1):   #逐行获取要导入的数据
            #通过数据的个数，是否为重复、是否有数值，这三个条件，来判断是否是有问题的数据，是否可以写入数据文件
            if len(indata[ro + 1]) == mocols and notrepeat(indata[ro + 1][0], indata[ro + 1][1]) and feikong(
                    *indata[ro + 1]):
                #逐个字段的写入数据，并判断如果是最后一个字段，则要填写当前的时间。
                for cl in range(mocols):
                    # print("正在写入数据！")
                    # print(indata[ro + 1])
                    if cl!=mocols-1:
                        modata_sheet.write(morows, cl, indata[ro + 1][cl])
                    else:
                        modata_sheet.write(morows,cl,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                morows += 1
            # if len(indata[ro+1]) and notrepeat(indata[ro+1][0],indata[ro+1][1]):
            #     print(indata[ro+1])

            else:
                errlist.append(ro + 1)

        modata.save(path_text.get())

        # print("共有{0}行数据，已成功导入{1}行数据！".format(inrows - 1, inrows - 1 - len(errlist)))
        errstr = ""
        for ier in errlist:
            errstr = errstr + str(ier) + '行,'
        # print("有{0}行数据导入失败！".format(len(errlist)))
        # print("数据中第{0}数据有问题，请检查后导入".format(errstr))
        if inrows - 1 - len(errlist) == 0:
            tkinter.messagebox.showerror('错误',
                                         '共有{0}行数据'.format(inrows - 1) + '\n' + '有{0}行数据导入失败！'.format(
                                             len(errlist)) + '\n' + '数据中第{0}数据有问题，请检查后导入'.format(errstr))
            print('errstr1',errstr)
        elif len(errlist)!=0:
            tkinter.messagebox.showinfo('提示', '共有{0}行数据，已成功导入{1}行数据！'.format(inrows - 1, inrows - 1 - len(errlist)) + '\n' + '有{0}行数据导入失败！'.format(len(errlist)) + '\n' + '数据中第{0}数据有问题，请检查后导入'.format(errstr))
            print('errstr2:',errstr)
        else:
            tkinter.messagebox.showinfo('提示','共有{0}行数据，全部成功导入！'.format(inrows - 1) )
    except FileNotFoundError:
        tkinter.messagebox.showinfo("提示", "请选择要导入的数据表！")


inport_butt = tkinter.Button(wdow,
                             text='导入',
                             font=('Arial', 12),
                             command=inport,
                             width=10, height=2).place(x=350, y=550)

tkinter.mainloop()

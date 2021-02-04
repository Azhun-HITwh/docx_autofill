# !/usr/bin/env python3
# -*- coding: utf-8 -*-

""" a simple GUI for the daily filling of the parameters of the vehicle sample automatically in PATAC"""

__author__ = 'Azhun Zhu'

import os
import re
import xlrd
from mailmerge import MailMerge
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import base64
import logging

# 定义全局变量 获取手动输入值
global var3, var4
global dict4


# 参数类
class Para:
    def __init__(self, code):
        # 初始化属性code，即参数的泛亚编码
        self.code = code

    def get_index(self):
        # 获取参数索引
        para_index = num_patac.index("%s" % self.code)
        return para_index

    def get_value(self):
        # 获取参数值
        para_value = table.cell(Para.get_index(self), 4).value
        return para_value

    def get_name(self):
        # 获取参数名称
        para_name = table.cell(Para.get_index(self), 0).value
        return para_name

    def comma_check(self):
        # 检查参数是否含有“，”或“,”
        comma_check = re.search(r',', Para.get_value(self), re.M | re.I)
        return comma_check

    def slash_check(self):
        # 检查是否含有斜杠
        slash_check = re.search(r"/", Para.get_value(self), re.M | re.I)
        return slash_check


def main():
    # 自动获取工具GUI
    root = tk.Tk()  # 创建一个Tkinter.Tk()实例
    # root.withdraw()  # 将Tkinter.Tk()实例隐藏

    root.title("报表生成工具")  # 主窗口命名
    root.geometry('500x400')  # 主窗口大小
    tmp = open('tmp.ico', 'wb+')
    tmp.write(base64.b64decode(
        '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'))
    tmp.close()
    root.iconbitmap('tmp.ico')
    os.remove('tmp.ico')

    # 获取当前系统用户名
    user = os.getlogin()

    # frame blank label
    frame = tk.Frame()
    frame.pack()
    l_blank = tk.Label(frame, height=1)
    l_blank.pack()

    # 提示label
    l = tk.Label(root, text="请选择登记表模板 Please choose the file - template of sample", font=("Century Gothic", 10, 'bold'),
                 width=30, height=3, wraplength=155)
    l.pack()

    # 获取登记表模板路径的函数
    def get_file_template():

        global path_template
        path_template = tkinter.filedialog.askopenfilename(title="请选择输入登记表模板",
                                                           file=[("Microsoft Word Document", ".docx")])

        # 选择登记表文件错误
        # if type(path_template) != str:  # 文件类型错误
        #     # 弹出错误窗口
        #     tkinter.messagebox.showerror('错误', '请选择正确的登记表模板!')
        #     logging.error('A wrong format of template is selected.')
        #
        # if path_template == "":  # 未选择文件
        #     # 弹出错误窗口
        #     tkinter.messagebox.showerror('错误', '请选择正确的登记表模板!')
        #     logging.error('No template is selected.')

        # 获取登记表名称
        tmp = path_template.split("/")
        tmp1 = tmp[-1].split(".")
        global name_template
        name_template = tmp1[0]
        var.set(path_template)
        # global document_1
        # document_1 = MailMerge(var.get())  # MailMerge组件
        return var.get()

    # 选择文件button
    b1 = tk.Button(root, text="浏览... Open...", font=("Century Gothic", 10), width=12, height=1,
                   command=lambda: get_file_template())  # 需要使用匿名函数使事件手动触发
    b1.pack()

    # 显示选择文件路径
    var = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收get_file_template()函数的传出内容用以显示在标签上
    l1 = tk.Label(root, textvariable=var, font=("Century Gothic", 8), height=2, fg='blue', wraplength=350)
    l1.pack()

    # 提示label2
    l2 = tk.Label(root, text="请选择参数文件 Please choose the file - data of vehicle", font=("Century Gothic", 10, 'bold'),
                  width=30, height=3, wraplength=155)
    l2.pack()

    # 获取参数文件路径的函数
    def get_file_database():

        path_database = tkinter.filedialog.askopenfilename(title="请选择参数文件",
                                                           file=[("Microsoft Excel 97-2003 Worksheet", ".xls"),
                                                                 ("Microsoft Excel Worksheet", ".xlsx")])
        # 选择参数文件错误
        # if type(path_database) != str:  # 文件类型错误
        #     # 弹出错误窗口
        #     tkinter.messagebox.showerror('错误', '请选择正确的参数文件!')
        #     logging.error('A wrong format of database is selected.')
        #
        # if path_database == "":  # 未选择文件
        #     # 弹出错误窗口
        #     tkinter.messagebox.showerror('错误', '请选择正确的参数文件!')
        #     logging.error('No database is selected.')

        var2.set(path_database)
        # global data
        # data = xlrd.open_workbook(var2.get())  # 打开参数文件
        # global table
        # table = data.sheet_by_name("整车比较信息")  # 获取参数文件的指定worksheet
        # global num_patac
        # num_patac = table.col_values(1)  # 参数的泛亚编码
        return var2.get()

    # 选择文件button2
    b2 = tk.Button(root, text="浏览... Open...", font=("Century Gothic", 10), width=12, height=1,
                   command=lambda: get_file_database())  # 需要使用匿名函数使事件手动触发
    b2.pack()

    # 显示选择文件路径
    var2 = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收get_database_template()函数的传出内容用以显示在标签上
    l3 = tk.Label(root, textvariable=var2, font=("Century Gothic", 8), height=2, fg='blue', wraplength=350)
    l3.pack()

    # 生成登记表的函数
    def generate(path=None):

        global data
        data = xlrd.open_workbook(var2.get())  # 打开参数文件

        global table
        table = data.sheet_by_name("整车比较信息")  # 获取参数文件的指定worksheet

        global num_patac
        num_patac = table.col_values(1)  # 参数的泛亚编码

        global document_1
        document_1 = MailMerge(path_template)  # MailMerge组件

        global dict3

        para_unsort = document_1.get_merge_fields()  # 登记表模板中的field
        para = list(para_unsort)
        para.sort()
        para_excluded = []  # 登记表模板中，参数库中未包含的参数
        para_multinames = []  # 多值参数在参数库中的名称
        para_multivalues = []  # 多值参数的值
        para_multicodes = []  # 多指参数的field
        para_need_multivalues = ["P0018AVA", "P0047ABE", "P0290APT",
                                 "P0165ACH", "P0114ACH", "P0296ACH", "P0295ACH", "P0150APT",
                                 "P0011DPT", "P0278ECH", "P0263DCH", "P0114BPT", "P0014AVA",
                                 "P0312ABE", "P0310ABE", "P0311ABE", "P0313ABE", "P0314ABE", "P0039AZH",
                                 "P0164AIN", "P0092AIN", "P0098AIN", "P0100AIN", "P0106AIN", "P0113AIN", "P0115AIN",
                                 "P0004CIN", "P0006CIN", "P0096AIN", "P0093AIN",
                                 "P0117BPT"]  # 需要忽略逗号分割多值的参数
        para_special = ["P0028AVP-A", "P0028AVP-B", "P0028AVP-C", "P0028BVP-A", "P0028BVP-B", "P0028BVP-C"]  # 滑行曲线

        # 获取整车公告型号-添加至生成登记表的名称中
        Para("P0017AES")
        temp_name = Para("P0017AES")
        if temp_name.get_value().rstrip():
            typename_vehicle = temp_name.get_value().rstrip()
        else:
            tkinter.messagebox.showerror('警告', '缺少参数:整车型号！无法生成登记表！请补充后重新启动工具！')
            root.destroy()
            root.quit()

        # 遍历所有登记表模板中的field
        for i in para:
            if i in num_patac:  # 参数库中存在的field
                n1 = Para(i)
                v1 = n1.get_value()
                if n1.comma_check() is not None:  # 参数值中存在逗号
                    # 判断多值是否重复
                    tmp = []  # buffer
                    for item in v1.split(','):
                        item = item.rstrip()  # 去除字符串尾端空格
                        if item not in tmp:  # 去重
                            tmp.append(item)
                    if len(tmp) == 1:
                        dict_temp = {i: tmp[0]}
                        document_1.merge(parts=None, **dict_temp)
                    else:
                        # 抓取多值参数
                        para_multinames.append(n1.get_name())
                        para_multivalues.append(n1.get_value())
                        para_multicodes.append(i)
                else:
                    dict1 = {i: v1}
                    document_1.merge(parts=None, **dict1)
            else:
                # 抓取登记表中未包含在参数文件中的字段
                para_excluded.append(i)
        # 选择单个配置参数
        if para_multicodes:
            # 建立一个空list储存需要去除的参数
            data_del = []
            # 删除特例参数，不需要分割多值
            for item in para_multicodes:
                if item in para_need_multivalues:
                    data_del.append(item)
            for item in data_del:
                x1 = Para(item)
                para_multicodes.remove(item)
                para_multivalues.remove(x1.get_value())
                para_multinames.remove(x1.get_name())
                dict_temp = {item: x1.get_value()}
                document_1.merge(parts=None, **dict_temp)

            if para_multicodes:
                # 手动选择单配置参数值窗口
                window1 = tk.Toplevel()
                window1.title("请手动选择相应配置参数（仅单选）")
                window1.geometry('800x600')

                def myfunction(event):
                    canvas.configure(scrollregion=canvas.bbox("all"), width=750, height=500)

                myframe = tk.Frame(window1, relief='groove', width=100, height=100, bd=1)
                myframe.place(x=10, y=10)

                canvas = tk.Canvas(myframe)
                frame = tk.Frame(canvas)

                myscrollbar_v = tk.Scrollbar(myframe, orient="vertical", command=canvas.yview)
                myscrollbar_h = tk.Scrollbar(myframe, orient="horizontal", command=canvas.xview)
                canvas.configure(yscrollcommand=myscrollbar_v.set)
                canvas.configure(xscrollcommand=myscrollbar_h.set)
                myscrollbar_v.pack(side="right", fill="y")
                myscrollbar_h.pack(side="bottom", fill="x")

                canvas.pack(side="left")
                canvas.create_window((0, 0), window=frame, anchor='nw')
                frame.bind("<Configure>", myfunction)

                # initialize dict3 for parameters with multiple values
                if 'dict3' not in globals().keys():
                    dict3 = {}  # 用于获取所有选值
                elif list(dict3.keys()) != para_multicodes:
                    # print(para_multicodes)
                    # print(list(dict3.keys()))
                    dict3 = {}  # 如果更换登记表模板，所有参数均需要点击
                    logging.warning('If the template is changed, all the parameters should be re-selected.')
                else:
                    pass

                tmp1 = []

                # 获取radiobutton的text
                def get_input_value(event):
                    item = event.widget['text']
                    if item not in tmp1:  # 去重
                        tmp1.append(event.widget['text'])
                    # <class '_tkinter.Tcl_Obj'>
                    print(event.widget['variable'])
                    buffer = event.widget['variable']
                    idx = int(buffer)
                    # print(event.widget['value'])
                    dict_temp = {para_multicodes[idx]: event.widget['text']}
                    dict3.update(dict_temp)
                    return

                # 检查是否多值均被选择
                def check_status():
                    if len(dict3) != len(para_multicodes):
                        # print(dict3)
                        # print(para_multicodes)
                        tkinter.messagebox.showerror('错误', '请为全部多值参数选择相应配置!')
                        logging.error('Not all the parameters are distributed the setting.')
                    else:
                        document_1.merge(parts=None, **dict3)
                        # print(globals())
                        window1.quit()
                        window1.destroy()
                    return dict3

                # 关闭函数
                def close():
                    window1.quit()
                    window1.destroy()
                    # tkinter.messagebox.showerror('错误', '程序中止!请返回主界面重新启动程序！')
                    logging.error('Cancel is clicked.')
                    return

                for i in range(len(para_multinames)):  # 单列显示
                    # 多值参数名称label
                    tk.Label(frame, text="%s:" % para_multinames[i], font=("微软雅黑", 10), height=2).grid(row=i, column=0,
                                                                                                       padx=10, pady=10)
                    temp = []  # 列表元素去重
                    for item in para_multivalues[i].split(","):
                        item = item.rstrip()  # 去除字符串尾端空格
                        if item not in temp:
                            temp.append(item)
                    for j in range(len(temp)):  # 单个配置参数单选框创建
                        value = temp[j]
                        rb = tk.Radiobutton(frame, text=value, variable=i, value=value, bg="Grey", fg="Black",
                                            indicatoron=0, font=("Century Gothic", 12, "bold"), width=15,
                                            wraplength=100)
                        rb.grid(row=i, column=j + 3, padx=10, pady=10)
                        rb.bind("<Button-1>", get_input_value)

                # 确定 关闭 按钮frame
                frame1 = tk.Frame(window1)
                frame1.pack(side='bottom')

                # 确定窗口按键
                btn_ok = tk.Button(frame1, text="确定", command=lambda: check_status(), height=2, width=8,
                                   font=('微软雅黑', 12, 'bold'))
                btn_ok.pack(side='left', padx=10)

                # 取消按键
                btn_cancel = tk.Button(frame1, text="取消", command=lambda: close(), height=2, width=8,
                                       font=('微软雅黑', 12, 'bold'))
                btn_cancel.pack(side='right', padx=10)

                window1.mainloop()

        if not para_excluded:
            if path is None:
                # 将内容写入新word文件中
                document_1.write(
                    'D:\\sgmuserprofile\%s\Desktop\%s-%s.docx' % (user, name_template, typename_vehicle))
            else:
                document_1.write('%s.docx' % path)
            tkinter.messagebox.showinfo(title="Got it!", message="登记表已生成！")
            logging.info('Done.')
            # root.quit()
            # root.destroy()
        else:
            # 由于是新窗口不可使用tk.Tk()创建根窗口，否则无法与原来的根窗口交互！！！
            window = tk.Toplevel()
            window.title("手动修改未填写参数")
            window.geometry('800x600')

            # window.lift()

            # 关闭函数
            def close2():
                window.quit()
                window.destroy()
                # tkinter.messagebox.showerror('错误', '程序中止!请返回主界面重新启动程序！')
                logging.error('Cancel is clicked.')
                return

            def myfunction(event):
                canvas.configure(scrollregion=canvas.bbox("all"), width=750, height=500)

            myframe = tk.Frame(window, relief='groove', width=100, height=100, bd=1)
            myframe.place(x=10, y=10)

            canvas = tk.Canvas(myframe)
            frame = tk.Frame(canvas)

            myscrollbar_v = tk.Scrollbar(myframe, orient="vertical", command=canvas.yview)
            myscrollbar_h = tk.Scrollbar(myframe, orient="horizontal", command=canvas.xview)
            canvas.configure(yscrollcommand=myscrollbar_v.set)
            canvas.configure(xscrollcommand=myscrollbar_h.set)
            myscrollbar_v.pack(side="right", fill="y")
            myscrollbar_h.pack(side="bottom", fill="x")

            canvas.pack(side="left")
            canvas.create_window((0, 0), window=frame, anchor='nw')
            frame.bind("<Configure>", myfunction)

            # 手动输入登记表中未包含在参数库中的参数
            dict4 = {}
            # 滑行曲线分割
            if para_special[0] in para_excluded:
                Slip_curve_Emission5 = Para("P0028AVP")
                Slip_curve_Emission6 = Para("P0028BVP")
                # print(Slip_curve_Emission5.get_value())
                for i in range(3):
                    dict4[para_special[i]] = Slip_curve_Emission5.get_value().split(";")[i]
                    para_excluded.remove(para_special[i])
                for i in range(3):
                    dict4[para_special[i + 3]] = Slip_curve_Emission6.get_value().split(";")[i]
                    para_excluded.remove(para_special[i + 3])
            if "ratio_weight_axles" in para_excluded:
                if dict3:
                    if 'P0008AVP' in list(dict3.keys()):
                        tmp_fr = dict3["P0008AVP"]
                    else:
                        tmp_fr = Para("P0008AVP").get_value()
                    if 'P0005BVP' in list(dict3.keys()):
                        tmp_rr = dict3["P0005BVP"]
                    else:
                        tmp_rr = Para("P0005BVP").get_value()
                else:
                    tmp_fr = Para("P0008AVP").get_value()
                    tmp_rr = Para("P0005BVP").get_value()
                ratio = int(tmp_fr) / int(tmp_rr)
                dict4["ratio_weight_axles"] = str(round(ratio, 3))
                para_excluded.remove("ratio_weight_axles")

            # key写入dict4
            for item in para_excluded:
                dict4[item] = ""

            # get()获取entry内容
            def insert2(event):
                buffer2 = event.widget["textvariable"]
                # print(buffer2)
                temp_idx = re.findall('\d+', buffer2)
                # print(temp_idx[0])
                idx_1 = int(temp_idx[0]) - 2
                # print(idx_1)
                if idx_1 < len(para_excluded):
                    # print(para_excluded)
                    dict_tmp = {para_excluded[idx_1]: var_list['var_entry%d' % idx_1].get()}
                    dict4.update(dict_tmp)
                    # print(dict4)
                else:
                    times = idx_1 // len(para_excluded)
                    # print("times=", times)
                    idx_1 = idx_1 - len(para_excluded) * times
                    # print(para_excluded)
                    dict_tmp = {para_excluded[idx_1]: var_list['var_entry%d' % idx_1].get()}
                    dict4.update(dict_tmp)

            def check_entry_status():  # 手动输入参数值，确定按钮激活函数
                if "VIN(请输入17位号码)" in dict4.keys():
                    # VIN位数判断-17位
                    if len(dict4["VIN(请输入17位号码)"]) != 17:
                        # 弹窗错误提示
                        tkinter.messagebox.showerror('错误', '请输入17位正确VIN！')
                        logging.error('The length of VIN is not 17.')
                    elif "" in dict4.values():  # 判断是否有未填写的参数
                        # 弹窗错误提示
                        print(dict4)
                        tkinter.messagebox.showerror('错误', '请为所有未填参数输入参数值！')
                        logging.error('Not all the parameters are set.')
                    else:
                        document_1.merge(parts=None, **dict4)
                        dict4.clear()
                        if path is None:
                            # 将内容写入新word文件中
                            document_1.write(
                                'D:\\sgmuserprofile\%s\Desktop\%s-%s.docx' % (
                                    user, name_template, typename_vehicle))
                        else:
                            document_1.write('%s.docx' % path)
                        tkinter.messagebox.showinfo(title="Got it!", message="登记表已生成！")
                        logging.info('Done.')
                        window.quit()
                        window.destroy()
                        # root.quit()
                        # root.destroy()
                else:
                    if "" in dict4.values():
                        # 弹窗错误提示
                        tkinter.messagebox.showerror('错误', '请为所有未填参数输入参数值！')
                        logging.error('Not all the parameters are set.')
                    else:
                        document_1.merge(parts=None, **dict4)
                        if path is None:
                            # 将内容写入新word文件中
                            document_1.write(
                                'D:\\sgmuserprofile\%s\Desktop\%s-%s.docx' % (user, name_template, typename_vehicle))
                        else:
                            document_1.write('%s.docx' % path)
                        tkinter.messagebox.showinfo(title="Got it!", message="登记表已生成！")
                        logging.info('Done.')
                        window.quit()
                        window.destroy()
                    # root.quit()
                    # root.destroy()
                return

            # 将label标签的内容设置为字符类型，用var来接收Entry函数的传出内容用以显示在标签上，动态变量
            var_list = locals()
            for i in range(len(para_excluded)):
                var_list['var_entry%d' % i] = tk.StringVar()

            # 判别是否为偶数项
            if len(para_excluded) % 2 == 0:
                for i in range(0, len(para_excluded), 2):  # 两列显示
                    tk.Label(frame, text="%s:" % para_excluded[i], font=("微软雅黑", 10), height=2).grid(row=i,
                                                                                                     column=0,
                                                                                                     padx=10,
                                                                                                     pady=10)
                    tk.Label(frame, text="%s:" % para_excluded[i + 1], font=("微软雅黑", 10), height=2).grid(row=i,
                                                                                                         column=2,
                                                                                                         padx=10,
                                                                                                         pady=10)

                    entry1 = tk.Entry(frame, textvariable=var_list['var_entry%s' % i], show=None)
                    entry1.grid(row=i, column=1, padx=10, pady=10)
                    # bind函数须将grid分开写
                    entry1.bind("<FocusOut>", insert2)
                    entry2 = tk.Entry(frame, textvariable=var_list['var_entry%s' % (i + 1)], show=None)
                    entry2.grid(row=i, column=3, padx=10, pady=10)
                    entry2.bind("<FocusOut>", insert2)
            else:
                if len(para_excluded) == 1:
                    # 手动填写一个参数
                    tk.Label(frame, text="%s:" % para_excluded[0], font=("微软雅黑", 10), height=2).grid(row=0,
                                                                                                      column=0,
                                                                                                      padx=10,
                                                                                                      pady=10)

                    entry1 = tk.Entry(frame, textvariable=var_list['var_entry%s' % (len(para_excluded) - 1)],
                                      show=None)
                    entry1.grid(row=0, column=1, padx=10, pady=10)
                    entry1.bind("<FocusOut>", insert2)
                else:
                    if len(para_excluded) >= 15:
                        tkinter.messagebox.showerror('错误', '缺少参数过多，请在VTAPM中更新参数文件！')
                        logging.error('The database is lack of too many parameters.')
                        window.quit()
                        window.destroy()
                        root.quit()
                        root.destroy()
                    # 奇数项且个数不为1
                    for i in range(0, len(para_excluded) - 1, 2):
                        tk.Label(frame, text="%s:" % para_excluded[i], font=("微软雅黑", 10), height=2).grid(row=i,
                                                                                                         column=0,
                                                                                                         padx=10,
                                                                                                         pady=10)
                        tk.Label(frame, text="%s:" % para_excluded[i + 1], font=("微软雅黑", 10), height=2).grid(row=i,
                                                                                                             column=2,
                                                                                                             padx=10,
                                                                                                             pady=10)

                        entry1 = tk.Entry(frame, textvariable=var_list['var_entry%s' % i], show=None)
                        entry1.grid(row=i, column=1, padx=10, pady=10)
                        entry2 = tk.Entry(frame, textvariable=var_list['var_entry%s' % (i + 1)], show=None)
                        entry2.grid(row=i, column=3, padx=10, pady=10)
                        entry1.bind("<FocusOut>", insert2)
                        entry2.bind("<FocusOut>", insert2)
                    tk.Label(frame, text="%s:" % para_excluded[-1], font=("微软雅黑", 10), height=2).grid(
                        row=len(para_excluded) - 1, column=0,
                        padx=10, pady=10)

                    entry3 = tk.Entry(frame, textvariable=var_list['var_entry%s' % (len(para_excluded) - 1)],
                                      show=None)
                    entry3.grid(row=len(para_excluded) - 1, column=1, padx=10, pady=10)
                    entry3.bind("<FocusOut>", insert2)

            # 确定 关闭 按钮frame
            frame1 = tk.Frame(window)
            frame1.pack(side='bottom')

            # 确定按键
            btn_insert = tk.Button(frame1, text="确定", command=lambda: check_entry_status(), height=1, width=6,
                                   font=('微软雅黑', 12, 'bold'))
            btn_insert.grid(row=len(para_excluded) + 1 + 1, column=0, padx=20, pady=10)

            # 取消按键
            btn_cancel = tk.Button(frame1, text="取消", command=lambda: close2(), height=1, width=6,
                                   font=('微软雅黑', 12, 'bold'))
            btn_cancel.grid(row=len(para_excluded) + 1 + 1, column=3)

            window.mainloop()
        return

    # 生成文件button3
    b3 = tk.Button(root, text="生成 Got it！", font=("Century Gothic", 12, 'bold'), width=15, height=2,
                   command=lambda: generate())  # 需要使用匿名函数使事件手动触发
    b3.pack()

    # 另存为激活函数
    def save_as():

        save_path = tkinter.filedialog.asksaveasfilename(title=u'保存文件', file=[("Microsoft Word Document", ".docx")])
        if save_path != "":
            generate(save_path)
        else:
            generate()
        return

    # 另存为button4
    b4 = tk.Button(root, text="另存为 Save as...", font=("Century Gothic", 12, 'bold'), width=15, height=2,
                   command=lambda: save_as())  # 需要使用匿名函数使事件手动触发
    b4.pack()

    # copyright
    l4 = tk.Label(root, text='Copyright by PATAC D&K TA ©2020', font=('Century Gothic', 8, 'bold'))
    l4.pack(side='bottom')

    root.mainloop()


if __name__ == "__main__":
    main()

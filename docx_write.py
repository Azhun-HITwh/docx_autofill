import os
import xlrd
from mailmerge import MailMerge
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import re

#获取登记表模板路径的函数
def get_file_template():

    path_template = tkinter.filedialog.askopenfilename(title="请选择输入登记表模板", file=[("Microsoft Word Document", ".docx")])
    #获取登记表名称
    tmp = path_template.split("/")
    tmp1 = tmp[-1].split(".")
    global name_template
    name_template = tmp1[0]
    var.set(path_template)
    global document_1
    document_1 = MailMerge(var.get())  # MailMerge组件
    return var.get()

#获取参数文件路径的函数
def get_file_database():

    path_database= tkinter.filedialog.askopenfilename(title="请选择参数文件",
                                                       file=[("Microsoft Excel 97-2003 Worksheet", ".xls"),
                                                             ("Microsoft Excel Worksheet", ".xlsx")])
    #选择参数文件错误
    if type(path_database) != str:#文件类型错误
        #弹出错误窗口
        windows_error = tk.Toplevel()
        windows_error.title("错误")
        #修改窗口图片（预留）
        # windows_error.geometry("500x300")
        #窗口文字
        l = tk.Label(windows_error,text="请选择正确的参数文件(.xls,.xlsx)",font=("宋体",12),width=30,height=2)
        l.place(x=235,y=120,anchor="nw")
        #窗口按钮
        b = tk.Button(windows_error,text="确定",command=lambda :get_file_database()).place()
    if path_database == "":#未选择文件
        # 弹出错误窗口
        windows_error = tk.Toplevel()
        windows_error.title("错误")
        # 修改窗口图片（预留）
        # windows_error.geometry("500x300")
        # 窗口文字
        l = tk.Label(windows_error, text="请选择正确的参数文件(.xls,.xlsx)", font=("宋体", 12), width=30, height=2)
        l.place(x=235, y=120, anchor="nw")
        # 窗口按钮
        b = tk.Button(windows_error, text="确定", command=lambda: get_file_database()).place()
    else:
        var2.set(path_database)
        global data
        data = xlrd.open_workbook(var2.get())  # 打开参数文件
        global table
        table = data.sheet_by_name("整车比较信息")  # 获取参数文件的指定worksheet
        global num_patac
        num_patac = table.col_values(1)  # 参数的泛亚编码
    return var2.get()

#定义全局变量 获取手动输入值
global var3
global var4
global dict3

#生成登记表的函数
def generate():

    para_unsort = document_1.get_merge_fields()  # 登记表模板中的field
    para = list(para_unsort)
    para.sort()
    para_excluded = []  # 登记表模板中，参数库中未包含的参数
    para_multinames = [] # 多值参数在参数库中的名称
    para_multivalues = [] # 多值参数的值
    para_multicodes =[]  # 多指参数的field
    para_need_multivalues = ["P0018AVA","P0047ABE","P0290APT","P0165ACH","P0114ACH","P0296ACH","P0295ACH","P0150APT","P0011DPT"] # 需要忽略逗号分割多值的参数

    #获取整车公告型号-添加至生成登记表的名称中
    temp_name = Para("P0017AES")
    typename_vehicle = temp_name.get_value().rstrip()

    # 遍历所有登记表模板中的field
    for i in para:
        if i in num_patac:#参数库中存在的field
            n1 = Para(i)
            v1 = n1.get_value()
            if n1.comma_check() is not None:#参数值中存在逗号
                #判断多值是否重复
                tmp=[] #buffer
                for item in v1.split(','):
                    item = item.rstrip()  # 去除字符串尾端空格
                    if not item in tmp: # 去重
                        tmp.append(item)
                if len(tmp)==1:
                    dict={i:tmp[0]}
                    document_1.merge(parts=None, **dict)
                else:
                    #抓取多值参数
                    para_multinames.append(n1.get_name())
                    para_multivalues.append(n1.get_value())
                    para_multicodes.append(i)
                    continue
            else:
                dict1 = {i: v1}
                document_1.merge(parts=None, **dict1)
        else:
            # 抓取登记表中未包含在参数文件中的字段
            para_excluded.append(i)

    # 选择单个配置参数
    if para_multicodes !=[]:
        #建立一个空list储存需要去除的参数
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
            dict = {item: x1.get_value()}
            document_1.merge(parts=None, **dict)


        if para_multinames != []:
            #手动选择单配置参数值窗口
            window1 = tk.Toplevel()
            window1.title("请手动选择相应配置参数")
            tmp1 = []
            dict3 = {} #用于获取所有选值
            # 获取radiobutton的text
            def get_input_value(event):
                item = event.widget['text']
                if not item in tmp1: #去重
                    tmp1.append(event.widget['text'])
                # <class '_tkinter.Tcl_Obj'>
                print(event.widget['variable'])
                buffer = event.widget['variable']
                idx = int(buffer)
                a = para_multicodes[idx]
                b = event.widget['text']
                dict_temp = {a : b}
                dict3.update(dict_temp)
                return

            #检查是否多值均被选择
            def check_status():
                if len(dict3) != len(para_multicodes):
                    tkinter.messagebox.showinfo(title="Error!",message="请为全部多值参数选择相应配置!")

                else:
                    document_1.merge(parts=None, **dict3)
                    window1.quit()
                    window1.destroy()
            #关闭函数
            def close():
                window1.quit()
                window1.destroy()

            for i in range(len(para_multinames)):#单列显示
                #多值参数名称label
                tk.Label(window1,text="%s:"%para_multinames[i],font=("宋体",10),height=2).grid(row=i,column=0,padx=10,pady=10)
                temp=[]#列表元素去重
                for item in para_multivalues[i].split(","):
                    item = item.rstrip()#去除字符串尾端空格
                    if not item in temp:
                        temp.append(item)
                for j in range(len(temp)):#单个配置参数单选框创建
                    value = temp[j]
                    rb = tk.Radiobutton(window1,text=value,variable=i,value=value,bg="Grey",indicatoron=0)
                    rb.grid(row=i,column=j+3,padx=10,pady=10)
                    rb.bind("<Button-1>",get_input_value)

            #确定 关闭 按钮frame
            frame = tk.Frame(window1)
            frame.grid(row=len(para_multinames),column=0,columnspan=2)
            # 确定窗口按键
            btn_ok = tk.Button(frame, text="确定", command=lambda: check_status(), height=2, width=8,
                               font=('黑体', 12, 'bold')) \
                .grid(row=len(para_multinames), column=2, padx=20, pady=10)

            # 取消按键
            btn_cancel = tk.Button(frame, text="取消", command=lambda: close(), height=2, width=8,font=('黑体', 12, 'bold')) \
                .grid(row=len(para_multinames), column=4, padx=20, pady=10)

            window1.mainloop()

    if para_excluded ==[]:
        tkinter.messagebox.showinfo(title="报表生成工具",message="登记表已生成")
        document_1.write('D:\\sgmuserprofile\%s\Desktop\%s-%s.docx'% (user,name_template, typename_vehicle))  # 将内容写入新word文件中
    else:
        #由于是新窗口不可使用tk.Tk()创建根窗口，否则无法与原来的根窗口交互！！！
        window = tk.Toplevel()
        window.title("手动修改未填写参数")
        #手动输入登记表中未包含在参数库中的参数
        def insert():
            tmp = []
            if len(para_excluded) != 1:
                for i in range(0, len(para_excluded), 2):
                    tmp.append(var3.get())
                    tmp.append(var4.get())
                for i in range(len(para_excluded)):
                    dict2 = {para_excluded[i]: tmp[i]}
                    document_1.merge(parts=None, **dict2)
            else:
                tmp.append(var3.get())
                dict2 = {para_excluded[0]:tmp[0]}
                document_1.merge(parts=None, **dict2)
            window.quit()
            window.destroy()
            document_1.write('D:\\sgmuserprofile\%s\Desktop\%s-%s.docx'% (user,name_template, typename_vehicle))  # 将内容写入新word文件中
            tkinter.messagebox.showinfo(title="报表生成工具", message="登记表已生成")
            return

        # 关闭函数
        def close():
            window.quit()
            window.destroy()

        # 判别是否为偶数项
        if len(para_excluded)% 2 == 0:
            for i in range(0,len(para_excluded),2): #两列显示
                tk.Label(window, text="%s:"%para_excluded[i], font=("宋体", 10), height=2).grid(row=i,column=0,padx=10,pady=10)
                tk.Label(window, text="%s:" % para_excluded[i+1], font=("宋体", 10), height=2).grid(row=i, column=2, padx=10,
                                                                                                 pady=10)

                var3 = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收Entry函数的传出内容用以显示在标签上
                var4 = tk.StringVar()
                tk.Entry(window,textvariable=var3,show=None).grid(row=i,column=1,padx=10,pady=10)
                tk.Entry(window,textvariable=var4,show=None).grid(row=i, column=3, padx=10, pady=10)

        else:
            if len(para_excluded) == 1:
            #手动填写一个参数
                tk.Label(window,text="%s:" % para_excluded[0],font=("宋体", 10), height=2).grid(row=0, column=0,
                                                                                                padx=10, pady=10)
                var3 = tk.StringVar()
                tk.Entry(window, textvariable=var3, show=None).grid(row=0, column=1, padx=10,pady=10)
            else:
            #奇数项且个数不为1
                for i in range(0,len(para_excluded)-1,2):
                    tk.Label(window, text="%s:" % para_excluded[i], font=("宋体", 10), height=2).grid(row=i, column=0,
                                                                                                padx=10, pady=10)
                    tk.Label(window, text="%s:" % para_excluded[i + 1], font=("宋体", 10), height=2).grid(row=i, column=2,
                                                                                                    padx=10,
                                                                                                    pady=10)

                    var3 = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收Entry函数的传出内容用以显示在标签上
                    var4 = tk.StringVar()
                    tk.Entry(window, textvariable=var3, show=None).grid(row=i, column=1, padx=10, pady=10)
                    tk.Entry(window, textvariable=var4, show=None).grid(row=i, column=3, padx=10, pady=10)
                tk.Label(window, text="%s:" % para_excluded[-1], font=("宋体", 10), height=2).grid(row=len(para_excluded)//2+1, column=0,
                                                                                            padx=10, pady=10)
                var3 = tk.StringVar()
                tk.Entry(window,textvariable=var3,show=None).grid(row=len(para_excluded)//2+1,column=1,padx=10,pady=10)

        # 确定 关闭 按钮frame
        frame = tk.Frame(window)
        frame.grid(row=len(para_excluded)+1, column=0, columnspan=2)

        #确定按键
        btn_insert = tk.Button(frame,text="确定",command=lambda:insert(),height=2, width=8,
                           font=('黑体', 12, 'bold'))
        btn_insert.grid(row=len(para_excluded)+1+1,column=0)

        #取消按键
        btn_cancel = tk.Button(frame, text="取消", command=lambda: close(), height=2, width=8,
                               font=('黑体', 12, 'bold'))
        btn_cancel.grid(row=len(para_excluded) + 1 + 1, column=3)

        window.mainloop()
    return

#参数类
class Para():
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

#自动获取工具GUI
root = tk.Tk()  # 创建一个Tkinter.Tk()实例
# root.withdraw()  # 将Tkinter.Tk()实例隐藏

root.title("报表生成工具")#主窗口命名
root.geometry('500x300')#主窗口大小

#获取当前系统用户名
user = os.getlogin()

#提示label
l = tk.Label(root,text="请选择登记表模板",font=("宋体",12),width=30,height=2)
l.pack()

#选择文件button
b1 = tk.Button(root,text="浏览...",font=("宋体",12),width=30,height=1,command = lambda :get_file_template())#需要使用匿名函数使事件手动触发
b1.pack()

#显示选择文件路径
var = tk.StringVar() # 将label标签的内容设置为字符类型，用var来接收get_file_template()函数的传出内容用以显示在标签上
l1 = tk.Label(root,textvariable=var,font=("宋体",8),height=2,wraplength=350)
l1.pack()

#提示label2
l2 = tk.Label(root,text="请选择参数文件",font=("宋体",12),width=30,height=2)
l2.pack()

#选择文件button2
b2 = tk.Button(root,text="浏览...",font=("宋体",12),width=30,height=1,command = lambda :get_file_database())#需要使用匿名函数使事件手动触发
b2.pack()

#显示选择文件路径
var2 = tk.StringVar() # 将label标签的内容设置为字符类型，用var来接收get_database_template()函数的传出内容用以显示在标签上
l3 = tk.Label(root,textvariable=var2,font=("宋体",8),height=2,wraplength=350)
l3.pack()


#生成文件button3
b3 = tk.Button(root,text="生成",font=("宋体",12),width=30,height=1,
                command =lambda :generate())#需要使用匿名函数使事件手动触发
b3.pack()

root.mainloop()


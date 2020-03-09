import xlrd
from mailmerge import MailMerge
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import re


#获取登记表模板路径的函数
def get_file_template():
    path_template = tkinter.filedialog.askopenfilename(title="请选择输入登记表模板", file=[("Microsoft Word Document", ".docx")])
    var.set(path_template)
    global document_1
    document_1 = MailMerge(var.get())  # MailMerge组件
    # print(var.get())
    return var.get()

#获取参数文件路径的函数
def get_file_database():
    path_database= tkinter.filedialog.askopenfilename(title="请选择参数文件",
                                                       file=[("Microsoft Excel 97-2003 Worksheet", ".xls"),
                                                             ("Microsoft Excel Worksheet", ".xlsx")])
    var2.set(path_database)
    global data
    data = xlrd.open_workbook(var2.get())  # 打开参数文件
    global table
    table = data.sheet_by_name("整车比较信息")  # 获取参数文件的指定worksheet
    global num_patac
    num_patac = table.col_values(1)  # 参数的泛亚编码
    return var2.get()


#生成登记表的函数
def generate():

    para = document_1.get_merge_fields()  # 登记表模板中的field
    para_excluded = []  # 登记表模板中，参数库中未包含的参数
    para_multinames = [] # 多值参数在参数库中的名称
    para_multivalues= [] # 多值参数的值
    para_multicodes=[]  #多指参数的field


    # 遍历所有登记表模板中的field
    for i in para:
        if i in num_patac:#参数库中存在的field
            n1 = Para(i)
            v1 = n1.get_value()
            v2 = formatPara(v1)
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
            # elif n1.slash_check() is not None:#参数值中存在/
            #     # 抓取多值参数
            #     para_multinames.append(n1.get_name())
            #     para_multivalues.append(n1.get_value())
            #     para_multicodes.append(n1.code)
                # 去掉斜杠
                # v3 = v2.my_split()
                # continue
            else:
                v3 = v1
            dict1 = {i: v3}
            document_1.merge(parts=None, **dict1)
        else:
            # 抓取登记表中未包含在参数文件中的字段
            para_excluded.append(i)


    if para_multinames !=[]:#选择单个配置参数

        #手动选择单配置参数值窗口
        window1 = tk.Toplevel()
        window1.title("请手动选择相应配置参数")
        tmp = [] #获取radiobutton的text
        def get_input_value(event):
            item = event.widget['text']
            if not item in tmp: #去重
                tmp.append(event.widget['text'])
            # <class '_tkinter.Tcl_Obj'>
            print(event.widget['variable'])
            buffer = event.widget['variable']
            idx=int(buffer)
            # print(para_multicodes)

            dict3 = {para_multicodes[idx]:event.widget['text']}
            document_1.merge(parts=None,**dict3)
            return

        for i in range(len(para_multinames)):#单列显示
            #多值参数名称label
            tk.Label(window1,text="%s:"%para_multinames[i],font=("宋体",10),height=2).grid(row=i,column=0,padx=10,pady=10)
            var2 = tk.StringVar()
            temp=[]#列表元素去重
            for item in para_multivalues[i].split(","):
                item = item.rstrip()#去除字符串尾端空格
                if not item in temp:
                    temp.append(item)
            for j in range(len(temp)):#单个配置参数单选框创建
                value = temp[j]
                rb = tk.Radiobutton(window1,text=value,variable=i,value=value)
                rb.grid(row=i,column=j+3,padx=10,pady=10)
                rb.bind("<Button-1>",get_input_value)


        #确定窗口按键
        btn_ok = tk.Button(window1,text="确定",command=lambda :window1.quit()).grid(row=len(para_multinames),column=0,padx=10,pady=10)
        window1.mainloop()


    if para_excluded ==[]:
        tkinter.messagebox.showinfo(title="报表生成工具",message="登记表已生成")
        document_1.write('D:\\99-模板安全技术条件.docx')  # 将内容写入新word文件中
    else:
        #手动输入登记表中未包含在参数库中的参数
        def insert():
            tmp = []
            for i in range(0, len(para_excluded), 2):
                tmp.append(var.get())
                tmp.append(var1.get())
            for i in range(len(para_excluded)):
                dict2 = {para_excluded[i]: tmp[i]}
                # print(dict2)
                document_1.merge(parts=None, **dict2)
            tkinter.messagebox.showinfo(title="报表生成工具", message="登记表已生成")
            document_1.write('D:\\99-模板安全技术条件.docx')  # 将内容写入新word文件中
            return

        #由于是新窗口不可使用tk.Tk()创建根窗口，否则无法与原来的根窗口交互！！！
        window = tk.Toplevel()
        window.title("手动修改未填写参数")
        # window.geometry("500x300")

        # 判别是否为偶数项
        if len(para_excluded)% 2 == 0:
            for i in range(0,len(para_excluded),2): #两列显示
                tk.Label(window, text="%s:"%para_excluded[i], font=("宋体", 10), height=2).grid(row=i,column=0,padx=10,pady=10)
                tk.Label(window, text="%s:" % para_excluded[i+1], font=("宋体", 10), height=2).grid(row=i, column=2, padx=10,
                                                                                                 pady=10)

                var = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收Entry函数的传出内容用以显示在标签上
                var1 = tk.StringVar()
                tk.Entry(window,textvariable=var,show=None).grid(row=i,column=1,padx=10,pady=10)
                tk.Entry(window,textvariable=var1,show=None).grid(row=i, column=3, padx=10, pady=10)
        else:
            #奇数项
            for i in range(0,len(para_excluded)-1,2):
                tk.Label(window, text="%s:" % para_excluded[i], font=("宋体", 10), height=2).grid(row=i, column=0,
                                                                                                padx=10, pady=10)
                tk.Label(window, text="%s:" % para_excluded[i + 1], font=("宋体", 10), height=2).grid(row=i, column=2,
                                                                                                    padx=10,
                                                                                                    pady=10)

                var = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收Entry函数的传出内容用以显示在标签上
                var1 = tk.StringVar()
                tk.Entry(window, textvariable=var, show=None).grid(row=i, column=1, padx=10, pady=10)
                tk.Entry(window, textvariable=var1, show=None).grid(row=i, column=3, padx=10, pady=10)
            tk.Label(window, text="%s:" % para_excluded[:-1], font=("宋体", 10), height=2).grid(row=len(para_excluded), column=0,
                                                                                            padx=10, pady=10)
            var = tk.StringVar()
            tk.Entry(window,textvariable=var,show=None).grid(row=len(para_excluded),column=1,padx=10,pady=10)

        #确定按键
        btn_insert = tk.Button(window,text="确定",command=lambda:insert())
        btn_insert.grid(row=0,column=6)

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

# 对多值或组合参数format类
class formatPara():
    def __init__(self, value):
        self.value = value

    def my_trim(self):  # 删除逗号后的内容只保留单个配置
        for i in range(len(self.value)):
            if self.value[i] == "," or self.value[i] == "，":
                self.value = self.value[:i]
                break
            else:
                continue
        return self.value

    def my_split(self):  # 分割/后的内容只保留单个配置
        tmp = self.value.split("/")
        return tmp

#自动获取工具GUI
root = tk.Tk()  # 创建一个Tkinter.Tk()实例
# root.withdraw()  # 将Tkinter.Tk()实例隐藏

root.title("报表生成工具")#主窗口命名
root.geometry('500x300')#主窗口大小

#提示label
l = tk.Label(root,text="请选择登记表模板",font=("宋体",12),width=30,height=2)
l.pack()

#选择文件button
b1 = tk.Button(root,text="浏览...",font=("宋体",12),width=30,height=1,command = lambda :get_file_template())#需要使用匿名函数使事件手动触发
b1.pack()

#显示选择文件路径
var = tk.StringVar() # 将label标签的内容设置为字符类型，用var来接收get_file_template()函数的传出内容用以显示在标签上
l1 = tk.Label(root,textvariable=var,font=("宋体",8),height=2)
l1.pack()

#提示label2
l2 = tk.Label(root,text="请选择参数文件",font=("宋体",12),width=30,height=2)
l2.pack()

#选择文件button2
b2 = tk.Button(root,text="浏览...",font=("宋体",12),width=30,height=1,command = lambda :get_file_database())#需要使用匿名函数使事件手动触发
b2.pack()

#显示选择文件路径
var2 = tk.StringVar() # 将label标签的内容设置为字符类型，用var来接收get_database_template()函数的传出内容用以显示在标签上
l3 = tk.Label(root,textvariable=var2,font=("宋体",8),height=2)
l3.pack()


#生成文件button3
b3 = tk.Button(root,text="生成",font=("宋体",12),width=30,height=1,
                command =lambda :generate())#需要使用匿名函数使事件手动触发
b3.pack()

root.mainloop()

# if path_template == "":
#     master = Tk()
#     master.title("错误")
#     w = Label(master, text="请选择正确格式的登记表模板(.docx)", fg="red", height=6, width=50)
#     w.pack()
#     b = Button(master, text="Re-choose", command=lambda: get_file_template())  # command不允许带参，需要使用匿名函数传参，否则会自动执行command
#     b.pack()
#     b2 = Button(master, text="Quit", command=master.quit)
#     b2.pack()


# path_database = var2

# if get_file_database() == "":
#     master = Tk()
#     master.title("错误")
#     # get_file_database()
#     w = Label(master, text="请选择正确格式的参数文件(.xls,.xlsx)", fg="red", height=6, width=50)
#     w.pack()
#     b = Button(master, text="OK", command=lambda: get_file_database())  # command不允许带参，需要使用匿名函数传参，否则会自动执行command
#     b.pack()



# 手动输入参数模板，已注释
# eng_place = num_patac.index("P0065CPT")
# eng_place1 = table.cell(eng_place,4).value

# engine = num_patac.index("P0007APT")
# engine1 = table.cell(engine, 4).value


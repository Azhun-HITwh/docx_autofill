import base64

open_icon = open("aivox-1x2jr-001.ico", "rb")  # 选择图标文件
b64str = base64.b64encode(open_icon.read())
open_icon.close()
write_data = "img = '{0}'".format(b64str)
f = open("icon2.py", "w+")
f.write(write_data)  # 生成ASCII码
f.close()
import os
import base64
from tkinter import Tk, Frame, StringVar, Entry, Button, X, Y, LEFT
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo, showerror

from to_xls import xmind_to_xx, style_excel
from icon import img


picture = open("pic.ico", "wb+")
picture.write(base64.b64decode(img))
picture.close()

frame_tk = Tk()
frame_tk.geometry('500x100+400+200')
frame_tk.title('选择需要转换的xmind文件')
frame_tk.iconbitmap('pic.ico')
os.remove("pic.ico")

frame = Frame(frame_tk)
frame.pack(padx=10, pady=10)
frame1 = Frame(frame_tk)
frame1.pack(padx=10, pady=10)

v = StringVar()
ent = Entry(frame, width=50, textvariable=v).pack(fill=X, side=LEFT)


def fileopen():
    file_sql = askopenfilename()
    if file_sql:
        v.set(file_sql)


def pathopen():
    dir = os.path.join(os.path.split(v.get())[0], 'xls')
    os.system('start ' + dir)


def run(xmind_file_path):
    xmind_path, xmind_file = os.path.split(xmind_file_path)[0], os.path.split(xmind_file_path)[1]
    xls_path, xls_file = os.path.join(xmind_path, 'xls'), xmind_file.replace('xmind', 'xls')

    try:
        # 转xls
        a = xmind_to_xx(xmind_path, xmind_file, xmind_file.split('.')[0])
        a.to_excel(a.data_dict[0]['topic'])
        a.save_xls()

        b = style_excel(xls_path, xls_file, a.data_dict[0]['topic']['title'])
        b.merge_excel(b.calculate())
        b.save_style_excel(os.path.join(xls_path, xls_file))

        showinfo(title='完成！', message="转换完成")
    except:
        showerror(title='出错！', message="请选择正确的xmind文件")


btn = Button(frame, width=20, text='选择文件', font=('宋体', 14), command=fileopen).pack(fil=X, padx=10)
ext = Button(frame1, width=10, text='运行', font=("宋体", 14), command=lambda: run(v.get())).pack(fill=X, side=LEFT)
ext_ = Button(frame1, width=12, text='打开文件位置', font=("宋体", 14), command=pathopen).pack(fill=X, side=LEFT)
etb = Button(frame1, width=10, text='退出', font=("宋体", 14), command=frame_tk.quit).pack(fill=Y, padx=10)


if __name__ == '__main__':
    frame_tk.mainloop()

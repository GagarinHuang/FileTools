import tkinter as tk
from os.path import isfile
from os.path import isdir
from os.path import exists
from os.path import abspath
from os import makedirs
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from pandas import read_excel
from shutil import copy
from tkinter import scrolledtext

def selectFile():
    path_ = askopenfilename() #使用askdirectory()方法返回文件夹的路径
    if path_ == "":
        path1.get() #当打开文件路径选择框后点击"取消" 输入框会清空路径，所以使用get()方法再获取一次路径
    else:
        path_ = path_.replace("/", "\\")  # 实际在代码中执行的路径为“\“ 所以替换一下
        path1.set(path_)
    if txt1.get() != "" and path_ != "":
        #print("delete")
        txt1.delete(0, 'end')
    txt1.insert('insert', path_)

def selectPath():
    path_ = askdirectory() #使用askdirectory()方法返回文件夹的路径
    if path_ == "":
        path2.get() #当打开文件路径选择框后点击"取消" 输入框会清空路径，所以使用get()方法再获取一次路径
    else:
        path_ = path_.replace("/", "\\")  # 实际在代码中执行的路径为“\“ 所以替换一下
        path2.set(path_)
    if txt2.get() != "" and path_ != "":
        #print("delete")
        txt2.delete(0, 'end')
    txt2.insert('insert', path_)

def createDirs():
    #读取excel文件
    msgs = ""
    if txt1.get() != "" and txt2.get() != "":
        if not isfile(txt1.get()):
            msgs = ''.join([msgs, "文件", txt1.get(), "不存在\n"])
        if not isdir(txt2.get()):
            msgs = ''.join([msgs, "目录", txt2.get(), "不存在\n"])
        if isfile(txt1.get()) and isdir(txt2.get()):
            df = read_excel(txt1.get(), usecols=[0], header=None)  # 读取项目名称列
            df_li = df.values.tolist()
            dirs = []
            for s_li in df_li:
                dirs.append(s_li[0])
            for dir in dirs:
                resultPath = ''.join([txt2.get(), "\\" , dir])
                if not exists(resultPath): # 是否存在这个文件夹
                    msgs = ''.join([msgs, "目录", dir, "新建成功\n"])
                    makedirs(resultPath) # 如果没有这个文件夹，那就创建一个
                else:
                    msgs = ''.join([msgs, "目录", dir, "已存在，未新建成功\n"])
        #print(msgs)
    else:
        msgs = "excel导入文件，目的路径不能为空\n"
    scroll.insert(tk.END, msgs)
    scroll.update()

def renameFiles():
    #读取excel文件
    msgs = ""
    if txt1.get() != "" and txt2.get() != "":
        if not isfile(txt1.get()):
            msgs = ''.join([msgs, "文件", txt1.get(), "不存在\n"])
        if not isdir(txt2.get()):
            msgs = ''.join([msgs, "目录", txt2.get(), "不存在\n"])
        if isfile(txt1.get()) and isdir(txt2.get()):
            df1 = read_excel(txt1.get(), usecols=[0], header=None)  # 读取项目名称列
            df_li1 = df1.values.tolist()
            srcfiles = []
            df2 = read_excel(txt1.get(), usecols=[1], header=None)  # 读取项目名称列
            df_li2 = df2.values.tolist()
            dstfiles = []
            for s_li in df_li1:
                dstpath = txt2.get() + "\\" + s_li[0]
                if txt3.get() != "":
                    dstpath += "." + txt3.get()
                srcfiles.append(dstpath)
            for s_li in df_li2:
                dstpath = txt2.get() + "\\" + s_li[0]
                if txt3.get() != "":
                    dstpath += "." + txt3.get()
                dstfiles.append(dstpath)
            #print(srcfiles)
            #print(dstfiles)
            for index in range(0, len(srcfiles)):
                if not isfile(srcfiles[index]):
                    msgs = ''.join([msgs, "原始文件", srcfiles[index], "不存在\n"])
                if isfile(dstfiles[index]):
                    msgs = ''.join([msgs, "新命名文件", srcfiles[index], "已存在\n"])
                if isfile(srcfiles[index]) and not isfile(dstfiles[index]):
                    copy(srcfiles[index], dstfiles[index])          # 复制文件
                    msgs = ''.join([msgs, "新命名文件", srcfiles[index], "创建成功\n"])
                #print ("copy %s -> %s"%(srcfiles[index], dstfiles[index]))
    else:
        msgs = "excel导入文件，目的路径不能为空\n"
    scroll.insert(tk.END, msgs)
    scroll.update()

def clearMsgs():
    if scroll.get('1.0', 'end-1c') != "":    
        scroll.delete('1.0', 'end-1c')

if __name__ == "__main__":
    wnd = tk.Tk()
    wnd.title('tools')
    wnd.geometry('320x300')
    wnd.resizable(height=False, width=False)
    
    lbl1 = tk.Label(wnd, text='Excel')
    lbl1.grid(row=0, column=0, sticky='E')
    txt1 = tk.Entry(wnd)
    txt1.grid(row=0, column=1, sticky='E')
    path1 = tk.StringVar()
    path1.set(abspath("."))
    tk.Button(wnd,
              text="select file",
              command=selectFile).grid(row=0, column=2)
    
    lbl2 = tk.Label(wnd, text='Destination')
    lbl2.grid(row=1, column=0, sticky='E')
    txt2 = tk.Entry(wnd)
    txt2.grid(row=1, column=1, sticky='E')
    path2 = tk.StringVar()
    path2.set(abspath("."))
    tk.Button(wnd,
              text="select path",
              command=selectPath).grid(row=1, column=2)
    
    lbl3 = tk.Label(wnd, text='Form')
    lbl3.grid(row=2, column=0, sticky='E')
    txt3 = tk.Entry(wnd)
    txt3.grid(row=2, column=1, sticky='E')
    lbl4 = tk.Label(wnd, text='txt/pdf/word/etc')
    lbl4.grid(row=2, column=2, sticky='E')
    
    btn3 = tk.Button(wnd, text='Clear Msg', command=clearMsgs)
    btn3.grid(row=3, column=1, sticky='W')
    btn1 = tk.Button(wnd, text='Create Dirs', command=createDirs)
    btn1.grid(row=3, column=1, sticky='E')
    btn2 = tk.Button(wnd, text='Rename Files', command=renameFiles)
    btn2.grid(row=3, column=2, sticky='E')
    
    # Message
    scroll = scrolledtext.ScrolledText(wnd,width=40,height=13,font=('黑体',10))
    scroll.grid(row=4, column=0, columnspan=3, pady = 5 , padx = 5 )
    
    wnd.mainloop()

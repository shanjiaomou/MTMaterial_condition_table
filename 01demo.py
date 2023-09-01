"""
本代码由[Tkinter布局助手]生成
官网:https://www.pytk.net/tkinter-helper
QQ交流群:788392508
"""
import os
import openpyxl
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
from tkinter.ttk import *
from typing import Dict

#设置数组数据
def array_set(array, value, row, col):
    # 获取二维数组的行数和列数
    rows = len(array)   #行
    cols = len(array[0])#列
    # 如果插入位置超出数组范围，则扩展数组并使用空值补全
    if row >= rows or col >= cols:
        max_rows = max(row + 1, rows)
        max_cols = max(col + 1, cols)
        extended_array = [[None for _ in range(max_cols)] for _ in range(max_rows)]
        # 将原数组的值复制到扩展后的数组中
        for i in range(rows):
            for j in range(cols):
                extended_array[i][j] = array[i][j]
        array = extended_array
    # 在指定位置插入值
    array[row][col] = value
    return array
#获取数组数据
def array_get(array, row, col):
    # 获取二维数组的行数和列数
    rows = len(array)   #行
    cols = len(array[0])#列
    # 如果插入位置超出数组范围，则扩展数组并使用空值补全
    if row >= rows or col >= cols:
        return False
    return array[row][col]
#插入行
def insert_row(array, row_index):
    # 获取二维数组的行数和列数
    rows = len(array)
    cols = len(array[0])
    values=[None for _ in range(cols)]
    # 如果插入位置超出数组范围，则扩展数组并使用空值补全
    if row_index > rows:
        for i in range(rows, row_index ):
            array.insert(i, values)
    # 在指定位置插入新行
    array.insert(row_index, values)
    return array
#插入列
def insert_col(array, col_index):
    # 获取二维数组的行数和列数
    rows = len(array)
    cols = len(array[0])

    # 如果插入位置超出数组范围，则扩展数组并使用空值补全
    if col_index > cols:
        for i in range(cols, col_index ):
            for i in range(rows):
                array[i].insert(col_index, None)
    # 在指定位置插入新列
    for i in range(rows):
        array[i].insert(col_index, None)
    return array   
#获取行数
def array_row(array):
    return len(array)
#获取列数
def array_col(array):
    return len(array[0])
#错误弹窗
def show_error(srtbuf):
    messagebox.showerror("错误", srtbuf)
def print_label(self,srtbuf):
    self.tk_label_lm0opp7u["text"]=srtbuf
def print_log(self,srtbuf):
    self.tk_text_lm0p7c95.configure(state="normal")
    self.tk_text_lm0p7c95.insert(END,srtbuf+"\r\n")
    self.tk_text_lm0p7c95.configure(state="disabled")
    self.tk_text_lm0p7c95.see(END)
def set_Prog(self,value,maximum):
    self.tk_progressbar_lm0p6l8w["maximum"]=maximum
    self.tk_progressbar_lm0p6l8w["value"]=value
#创建数组，并初始化（设置表头）
#用下面函数处理数据

def Sheet_Handle(SheetBuf,array):
    
    pass

def main_app(self):
    # 获取文件夹中的所有文件和文件夹
    folder_path = self.tk_input_lm0omywa.get()
    try:
        file_list = os.listdir(folder_path)
    except:
        show_error("请选择正确BOM目录")
        return
    file_len = len(file_list)
    
    for file_count in range(file_len):
        set_Prog(self,file_count+1,file_len)
        file_name = file_list[file_count]
        print_label(self,"导入BOM:"+str(file_count+1)+"/"+str(file_len))
        print_log(self,"导入BOM:"+file_name)
        Model_Last=file_name.find(".xlsx")
        if(("BOM清单_L95"!=file_name[0:9])and(0>Model_Last)):
            print_log(self," └文件名错误")
            return
        BomWorkBook=openpyxl(folder_path+"/"+file_name,data_only=True)
        
        
        
        
        Model_Name = file_name[20:Model_Last]
        
        
        


class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_label_lm0omf5i = self.__tk_label_lm0omf5i(self)
        self.tk_input_lm0omywa = self.__tk_input_lm0omywa(self)
        self.tk_button_lm0on4cw = self.__tk_button_lm0on4cw(self)
        self.tk_button_lm0onmq9 = self.__tk_button_lm0onmq9(self)
        self.tk_label_lm0opp7u = self.__tk_label_lm0opp7u(self)
        self.tk_progressbar_lm0p6l8w = self.__tk_progressbar_lm0p6l8w(self)
        self.tk_text_lm0p7c95 = self.__tk_text_lm0p7c95(self)
    def __win(self):
        self.title("Tkinter布局助手")
        # 设置窗口大小、居中
        width = 420
        height = 440
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)
        # 自动隐藏滚动条
    def scrollbar_autohide(self,bar,widget):
        self.__scrollbar_hide(bar,widget)
        widget.bind("<Enter>", lambda e: self.__scrollbar_show(bar,widget))
        bar.bind("<Enter>", lambda e: self.__scrollbar_show(bar,widget))
        widget.bind("<Leave>", lambda e: self.__scrollbar_hide(bar,widget))
        bar.bind("<Leave>", lambda e: self.__scrollbar_hide(bar,widget))
    
    def __scrollbar_show(self,bar,widget):
        bar.lift(widget)
    def __scrollbar_hide(self,bar,widget):
        bar.lower(widget)
    
    def vbar(self,ele, x, y, w, h, parent):
        sw = 15 # Scrollbar 宽度
        x = x + w - sw
        vbar = Scrollbar(parent)
        ele.configure(yscrollcommand=vbar.set)
        vbar.config(command=ele.yview)
        vbar.place(x=x, y=y, width=sw, height=h)
        self.scrollbar_autohide(vbar,ele)
    def __tk_label_lm0omf5i(self,parent):
        label = Label(parent,text="BOM目录：",anchor="center", )
        label.place(x=20, y=20, width=70, height=30)
        return label
    def __tk_input_lm0omywa(self,parent):
        ipt = Entry(parent, )
        ipt.place(x=90, y=20, width=310, height=30)
        ipt.insert(0, (os.path.dirname(os.path.abspath(__file__)))+'\BOM') # 获取.py文件所在的目录路径
        return ipt
    def __tk_button_lm0on4cw(self,parent):
        btn = Button(parent, text="选择目录", takefocus=False,)
        btn.place(x=260, y=60, width=60, height=30)
        return btn
    def __tk_button_lm0onmq9(self,parent):
        btn = Button(parent, text="开始处理", takefocus=False,)
        btn.place(x=340, y=60, width=60, height=30)
        return btn
    def __tk_label_lm0opp7u(self,parent):
        label = Label(parent,text="正在处理：",anchor="w" )
        label.place(x=20, y=60, width=220, height=30)
        return label
    def __tk_progressbar_lm0p6l8w(self,parent):
        progressbar = Progressbar(parent, orient=HORIZONTAL,)
        progressbar.place(x=20, y=100, width=380, height=10)
        return progressbar
    def __tk_text_lm0p7c95(self,parent):
        text = Text(parent,state="disabled")
        text.place(x=20, y=120, width=380, height=300)
        self.vbar(text, 20, 120, 380, 300,parent)
        return text
class Win(WinGUI):
    def __init__(self):
        super().__init__()
        self.__event_bind()
    def OpenFolderEvent(self,evt):
        folder_path = filedialog.askdirectory(initialdir="./")  # 打开文件夹选择对话框，设置初始目录为.py文件所在的位置
        self.tk_input_lm0omywa.delete(0, END)  # 清空输入框内容
        self.tk_input_lm0omywa.insert(END, folder_path)  # 将选择的目录路径填入输入框
    def StartProcessEvent(self,evt):
        main_app(self)
    def __event_bind(self):
        self.tk_button_lm0on4cw.bind('<Button>',self.OpenFolderEvent)
        self.tk_button_lm0onmq9.bind('<Button>',self.StartProcessEvent)
        pass
if __name__ == "__main__":
    win = Win()
    win.mainloop()

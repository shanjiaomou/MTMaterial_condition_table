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
import time


#设置数组数据
def array_set(array, row, col, value):
    # 获取二维数组的行数和列数
    rows = len(array)
    cols = len(array[0])
    # 如果插入位置超出数组范围，则扩展数组并使用空值补全
    if col >= cols:
        for i in range(cols, col ):
            for i in range(rows):
                array[i].insert(col, None)
        # 在指定位置插入新列
        for i in range(rows):
            array[i].insert(col, None)
        # 获取二维数组的行数和列数
        cols = len(array[0])
    values=[None for _ in range(cols)]
    # 如果插入位置超出数组范围，则扩展数组并使用空值补全
    if row >= rows:
        for i in range(rows, row ):
            array.insert(i, values)
        # 在指定位置插入新行
        array.insert(row, values)
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
#插入剪切行
def CutInsert_row(array,origin,finish):
    # 将第origin行剪切插入第finish行前面
    row_to_insert = array.pop(origin)  # 移除第origin行并保存
    array.insert(finish, row_to_insert)  # 在第finish行前面插入移除的行
    return array
#插入剪切列
def CutInsert_col(array,origin,finish):
    # 将第origin列剪切插入第finish列前面
    for row in array:
        col_to_insert = row.pop(origin)  # 移除第origin列并保存
        row.insert(finish, col_to_insert)  # 在第finish列前面插入移除的列
    return array
#获取行数
def array_row(array):
    return len(array)
#获取列数
def array_col(array):
    return len(array[0])
#UI界面
def show_error(srtbuf):
    messagebox.showerror("错误", srtbuf)
def print_label(self,srtbuf):
    self.tk_label_lm3lw72f["text"]=srtbuf
def print_log(self,srtbuf):
    self.tk_text_lm3m3ylm.configure(state="normal")
    self.tk_text_lm3m3ylm.insert(END,srtbuf+"\r\n")
    self.tk_text_lm3m3ylm.configure(state="disabled")
    self.tk_text_lm3m3ylm.see(END)
def Clear_log(self):
    self.tk_text_lm3m3ylm.configure(state="normal")
    self.tk_text_lm3m3ylm.delete("1.0", END)
    self.tk_text_lm3m3ylm.configure(state="disabled")
def set_Prog(self,value,maximum):
    self.tk_progressbar_lm3lze9x["maximum"]=maximum
    self.tk_progressbar_lm3lze9x["value"]=value
#App#######################################################################

def main_app(self):
    Clear_log(self)
    Bom_path = self.tk_input_lm3ln8cp.get()
    Plan_path = self.tk_input_lm3lqahh.get()
    if(os.path.isfile(Bom_path)==False):
        show_error("请正确选择BOM清单!")
        return
    if(os.path.isfile(Plan_path)==False):
        show_error("请正确选择计划!")
        return
    Bom_name = os.path.basename(Bom_path)
    Plan_name = os.path.basename(Plan_path)
    

#UI#######################################################################

class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.checkbox_var = IntVar(value=0)
        self.tk_label_lm3lm3mh = self.__tk_label_lm3lm3mh(self)
        self.tk_input_lm3ln8cp = self.__tk_input_lm3ln8cp(self)
        self.tk_button_lm3lnzud = self.__tk_button_lm3lnzud(self)
        self.tk_label_lm3lohwd = self.__tk_label_lm3lohwd(self)
        self.tk_input_lm3lqahh = self.__tk_input_lm3lqahh(self)
        self.tk_button_lm3lqeut = self.__tk_button_lm3lqeut(self)
        self.tk_label_lm3lw72f = self.__tk_label_lm3lw72f(self)
        self.tk_check_button_lm3lwui9 = self.__tk_check_button_lm3lwui9(self)
        self.tk_button_lm3lx0ji = self.__tk_button_lm3lx0ji(self)
        self.tk_progressbar_lm3lze9x = self.__tk_progressbar_lm3lze9x(self)
        self.tk_text_lm3m3ylm = self.__tk_text_lm3m3ylm(self)
    def __win(self):
        self.title("Tkinter布局助手")
        # 设置窗口大小、居中
        width = 600
        height = 290
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
    def __tk_label_lm3lm3mh(self,parent):
        label = Label(parent,text="BOM清单：",anchor="center", )
        label.place(x=20, y=20, width=80, height=30)
        return label
    def __tk_input_lm3ln8cp(self,parent):
        ipt = Entry(parent, )
        ipt.place(x=100, y=20, width=420, height=30)
        return ipt
    def __tk_button_lm3lnzud(self,parent):
        btn = Button(parent, text="选择", takefocus=False,)
        btn.place(x=530, y=20, width=50, height=30)
        return btn
    def __tk_label_lm3lohwd(self,parent):
        label = Label(parent,text="生产计划表：",anchor="center", )
        label.place(x=20, y=60, width=80, height=30)
        return label
    def __tk_input_lm3lqahh(self,parent):
        ipt = Entry(parent, )
        ipt.place(x=100, y=60, width=420, height=30)
        return ipt
    def __tk_button_lm3lqeut(self,parent):
        btn = Button(parent, text="选择", takefocus=False,)
        btn.place(x=530, y=60, width=50, height=30)
        return btn
    def __tk_label_lm3lw72f(self,parent):
        label = Label(parent,text="请选择文件并开始",anchor="w", )
        label.place(x=20, y=100, width=360, height=30)
        return label
    def __tk_check_button_lm3lwui9(self,parent):
        cb = Checkbutton(parent,text="生成格式",variable=self.checkbox_var)
        cb.place(x=400, y=100, width=80, height=30)
        return cb
    def __tk_button_lm3lx0ji(self,parent):
        btn = Button(parent, text="开始生成", takefocus=False,)
        btn.place(x=500, y=100, width=80, height=30)
        return btn
    def __tk_progressbar_lm3lze9x(self,parent):
        progressbar = Progressbar(parent, orient=HORIZONTAL,)
        progressbar.place(x=20, y=140, width=560, height=10)
        return progressbar
    def __tk_text_lm3m3ylm(self,parent):
        text = Text(parent,state="disabled")
        text.place(x=20, y=170, width=560, height=100)
        self.vbar(text, 20, 170, 560, 100,parent)
        return text
class Win(WinGUI):
    def __init__(self):
        super().__init__()
        self.__event_bind()
    def OpenBomEvent(self):#(self,evt):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if(file_path!=""):
            self.tk_input_lm3ln8cp.delete(0, END)  # 清空输入框内容
            self.tk_input_lm3ln8cp.insert(END, file_path)  # 将选择的目录路径填入输入框
    def OpenPlanEvent(self):#(self,evt):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if(file_path!=""):
            self.tk_input_lm3lqahh.delete(0, END)  # 清空输入框内容
            self.tk_input_lm3lqahh.insert(END, file_path)  # 将选择的目录路径填入输入框
    def StartProcessEvent(self):#(self,evt):
        self.tk_button_lm3lx0ji.config(state=DISABLED)
        self.tk_check_button_lm3lwui9.config(state=DISABLED)
        win.update()#更新界面
        main_app(self)
        self.tk_check_button_lm3lwui9.config(state=NORMAL)
        self.tk_button_lm3lx0ji.config(state=NORMAL)
        
    def __event_bind(self):
        #self.tk_button_lm3lnzud.bind('<Button>',self.OpenBomEvent)
        #self.tk_button_lm3lqeut.bind('<Button>',self.OpenPlanEvent)
        #self.tk_button_lm3lx0ji.bind('<Button>',self.StartProcessEvent)
        self.tk_button_lm3lnzud["command"]=self.OpenBomEvent
        self.tk_button_lm3lqeut["command"]=self.OpenPlanEvent
        self.tk_button_lm3lx0ji["command"]=self.StartProcessEvent
        pass
if __name__ == "__main__":
    win = Win()
    win.mainloop()

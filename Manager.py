from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename  # (askopenfilename, askopenfilenames, askdirectory, asksaveasfilename)
import pandas as pd
# import matplotlib.pyplot as plt
from pandas.core.frame import DataFrame
from tabulate import tabulate

# 全局变量定义
# Auto_message='初始消息'
# ExcelFile='初始路径'
StandData_code = DataFrame()
StandData_stock = DataFrame()


class Windows(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()

    def init_window(self):
        self.master.title('Sisyphe采销闭合分析')
        self.master.iconbitmap('D:\\VS py\\Sisyphe\\Source code\\PM Analyze\\Sisyphe.ico')
        self.pack(fill=BOTH, expand=1)
        # 按钮
        stock_describe_button = Button(self, text='库存分析', command=stock_describe)
        stock_describe_button.place(x=1, y=1)
        # importButton=Button(self,text='导入数据',command=data_import)
        # importButton.place(x=1,y=30)
        # 菜单
        menu = Menu(self.master)
        self.master.config(menu=menu)
        file = Menu(menu)
        file.add_command(label='导入数据', command=data_import)
        file.add_command(label='退出程序', command=self.client_exit)
        menu.add_cascade(label='文件', menu=file)
        edit = Menu(menu)
        edit.add_command(label='Undo')
        edit.add_command(label='Redo')
        menu.add_cascade(label='edit', menu=edit)

    @staticmethod
    def client_exit():
        exit()


def message_info(auto_message):
    messagebox.showinfo("提示", auto_message)


def data_import():
    excel_file = askopenfilename(title='选择文件', initialdir='/', filetypes=[('excel file', '*.xlsx')])
    global StandData_code, StandData_stock
    StandData_code = pd.read_excel(excel_file, sheet_name='供应商分级')
    StandData_stock = pd.read_excel(excel_file, sheet_name='全量库存21-1025')
    message_info('导入完成')


def print_cell(self):
    print(tabulate(self, headers='keys', tablefmt='psql'))
    return


def stock_describe():
    print_cell(StandData_stock.describe())
    message_info('完毕！请查看控制台')


root = Tk()
root.geometry('400x300')
app = Windows(root)
message_info('注意！请先在文件菜单导入数据')
root.mainloop()

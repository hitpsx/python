# coding=utf-8
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def splitExcel(fileName, columnName):
    df = pd.read_excel(fileName, dtype=str)
    columnList = list(df[columnName].drop_duplicates())
    for column in columnList:
        res = df[columnName] == column
        df[res].to_excel('{0}.xlx'.format(column), index=False)


def getColumnName(fileName):
    df = pd.read_excel(fileName, dtype=str)
    columnList = list(df.columns.values)
    return columnList


def click():
    # 设置可以选择的文件类型，不属于这个类型的，无法被选中
    global columnList
    filetypes = [("excel文件", "*.xlsx"), ('excel文件', '*.xls')]
    file_name = filedialog.askopenfilename(title='选择单个文件',
                                           filetypes=filetypes,
                                           initialdir='./'  # 打开当前程序工作目录
                                           )
    path_var.set(file_name)
    columnList = getColumnName(file_name)

columnList = []
window = tk.Tk()
window.title('文件对话框')  # 设置窗口的标题
window.geometry('300x50')  # 设置窗口的大小

path_var = tk.StringVar()
entry = tk.Entry(window, textvariable=path_var)
entry.place(x=10, y=10, anchor='nw')

tk.Button(window, text='选择', command=click).place(x=220, y=10, anchor='nw')

varList = [tk.IntVar for i in range(100)]
i = 0
for column in columnList:
    c1 = window.Checkbutton(window, text=column, variable=varList[i], onvalue=1, offvalue=0)
    c1.place()
    i += 1
window.mainloop()

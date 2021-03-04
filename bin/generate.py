#!/usr/bin/python
# -*- coding: UTF-8 -*-
import base64
import tkinter as tk

window = tk.Tk()
window.title("授权码生成器 V1.0")
window.geometry('300x300')


# 生成授权码 生成规则自己随意定的
def get_code():
    code = e.get()
    active_code = base64.b64encode(code.encode('utf-8'))
    active_code = str(active_code, 'utf-8')
    t.insert('end', active_code)


if __name__ == '__main__':
    label1 = tk.Label(window, text="请输入机器码：", height=4)
    label1.place(x=20, y=10)
    e = tk.Entry(window, width=30)
    e.place(x=30, y=70)
    b = tk.Button(window, text="生成授权", width=15, command=get_code)
    b.place(x=30, y=100)
    label2 = tk.Label(window, text="授权许可证：", height=4)
    label2.place(x=30, y=130)
    t = tk.Text(window, height=4, width=30)
    t.place(x=30, y=150)
    window.mainloop()

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import base64
import os
from tkinter.filedialog import *
from time import sleep, localtime, strftime, time
from threading import Thread

import wmi

from bin.excel import ReadExcel
from bin.operation import Operation


class MainWindow:
    def __init__(self, init_window_name, log, base_path):
        self.init_window_name = init_window_name
        self.log = log
        self.base_path = base_path
        self.config_path = os.path.join(base_path, 'config', 'config.ini')

        self.send_num = 0
        self.is_start = 0
        self.machine_code = ""

        self.init_data_label = Label(self.init_window_name, text=u"设置发送间隔时间：", padx=10, pady=10)
        self.init_data_label.grid(row=0, column=0)
        # 间隔时间
        self.time_entry = Entry(self.init_window_name)  # 原始数据录入框
        self.time_entry.grid(row=0, column=1)

        # 选择excel
        self.excel_label = Label(self.init_window_name, text=u"选择excel：")
        self.excel_label.grid(row=2, column=0)

        self.excel_entry = Entry(self.init_window_name)
        self.excel_entry.grid(row=2, column=1)

        # # 机器码
        self.machine_entry = Entry(self.init_window_name, width=40)
        self.machine_entry.grid(row=9, column=0, columnspan=10)
        self.machine_button = Button(self.init_window_name, text="查询", bg="lightblue", width=10,
                                     command=self.get_machine_code)
        self.machine_button.grid(row=9, column=13)

        # 授权码
        self.active_entry = Entry(self.init_window_name, width=40)
        self.active_entry.grid(row=10, column=0, columnspan=10)
        self.active_button = Button(self.init_window_name, text="授权", bg="lightblue", width=10,
                                    command=self.active)
        self.active_button.grid(row=10, column=13)

        # 开始运行
        self.btn_start = Button(self.init_window_name, text='开始', bg="lightblue", width=12, command=self.handle)
        # self.btn_start.grid(row=5, column=1)
        # self.btn_start.config(state='disabled')
        self.btn_start.grid_forget()  # 隐藏

        # 停止运行
        self.btn_stop = Button(self.init_window_name, text='停止', bg="lightblue", width=12, command=self.stop)
        # self.btn_start.grid(row=5, column=1)
        # self.btn_start.config(state='disabled')
        # self.btn_stop.grid_forget()      # 隐藏

        # 日志
        self.result_data_label = Label(self.init_window_name, text="日志")
        self.result_data_label.grid(row=0, column=14)
        self.result_data_Text = Text(self.init_window_name, width=150, height=49, wrap='none')  # 处理结果展示
        self.result_data_Text.grid(row=1, column=14, rowspan=15, columnspan=10)

        # 按钮
        self.str_trans_to_md5_button = Button(self.init_window_name, text="浏览", bg="lightblue", width=10,
                                              command=self.browse)
        self.str_trans_to_md5_button.grid(row=2, column=13)

    # 设置窗口
    def set_init_window(self):
        self.init_window_name.title("自动发送消息工具_V1.0")  # 窗口名
        # self.init_window_name.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.init_window_name.geometry('748x719')
        # self.init_window_name["bg"] = "pink"        #　窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        # self.init_window_name.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        self.check_active()

    # 检查是否激活
    def check_active(self):
        """
        检查是否激活
        :return:
        """
        active_path = r'C:\active'
        if not os.path.exists(active_path):
            return 0
        with open(active_path, 'r') as f:
            active_code = f.readline()
        self.active_entry.insert(0, active_code)
        self.get_machine_code()
        self.active()

    # 获取当前时间
    @staticmethod
    def get_current_time():
        current_time = strftime('%Y-%m-%d %H:%M:%S', localtime(time()))
        return current_time

    # 日志动态打印
    def write_log(self, msg):
        current_time = self.get_current_time()
        msg = str(current_time) + " " + str(msg) + "\n"  # 换行
        self.result_data_Text.insert(END, msg)

    # 浏览
    def browse(self):
        """
        浏览excel文件
        :return:
        """
        file_name = askopenfilename(title=u'选择文件')
        self.excel_entry.select_clear()
        self.excel_entry.insert(0, file_name)

    # 授权
    def active(self):
        input_code = self.active_entry.get()

        active_code = base64.b64encode(self.machine_code.encode('utf-8'))
        active_code = str(active_code, 'utf-8')
        if active_code == input_code:
            self.log.logger.info('active successfully')
            self.btn_start.config(state='active')
            self.btn_start.grid(row=5, column=1)
            self.active_entry.grid_forget()
            self.active_button.grid_forget()
            self.machine_entry.grid_forget()
            self.machine_button.grid_forget()
            # 将激活码保存
            active_path = r'C:\active'
            with open(active_path, 'w') as f:
                f.write(active_code)

    # 获取机器码
    def get_machine_code(self):
        """
        获取机器码
        :return:
        """
        s = wmi.WMI()
        cpu_info = s.Win32_Processor()
        self.machine_code = cpu_info[0].ProcessorId + 'L'
        self.machine_entry.insert(0, self.machine_code)

    # 运行
    def start(self):
        """
        发送短信
        :return:
        """
        file_name = self.excel_entry.get()
        _time = self.time_entry.get()
        # self.btn_start.config(state='disabled')
        self.btn_start.grid_forget()  # 隐藏
        self.btn_stop.grid(row=5, column=1)
        self.is_start = 1

        while True:
            if not self.is_start:
                break
            try:
                phone, msg = get_msg(file_name, self.send_num + 1)
            except IOError as e1:
                print(e1)
                self.write_log('任务完成')
                self.btn_start.config(state='active')
                self.send_num = 0
                break
            send_code = send2(phone, msg)
            if send_code:
                self.send_num += 1
                if send_code == 2:
                    self.write_log('%s %s 发送失败 %s' % (phone, self.send_num, msg))
                else:
                    self.write_log('%s %s发送成功 %s' % (phone, self.send_num, msg))
            else:
                self.write_log('%s %s 未发送 %s' % (phone, self.send_num, msg))

            # sleep时，监控是否停止发送
            sleep_time = int(_time) - 5
            for i in range(sleep_time):
                if not self.is_start:
                    break
                sleep(1)
        self.btn_stop.grid_forget()  # 隐藏
        self.btn_start.grid(row=5, column=1)

    # 停止
    def stop(self):
        if self.is_start:
            self.is_start = 0
        self.btn_stop.grid_forget()  # 隐藏
        self.btn_start.grid(row=5, column=1)

    def handle(self):
        thread = Thread(target=self.start)
        thread.start()


def get_msg(file_name, index):
    """
    获取excel中的数据
    :param index: 第几行数据
    :param file_name:
    :return:
    """
    excel = ReadExcel(file_name)
    data = excel.read_row(index)
    return int(data[0]), data[1]


def send2(phone, msg):
    """
    使用uiautomation发送短信
    :param phone:
    :param msg:
    :return:
    """
    operation = Operation()
    operation.set_receiver(phone)
    operation.set_msg(msg)
    msg_num = operation.get_msg_num()
    operation.send_msg()
    sleep(20)
    msg_num2 = operation.get_msg_num()
    if msg_num < msg_num2:
        print('消息已发送')
        if operation.get_msg_status(msg_num2):
            print('发送成功')
            return 1
        else:
            print('发送失败')
            return 2
    else:
        print('消息未发送')
        return 0


def run(log, base_path):
    log.logger.info('start')
    init_window = Tk()  # 实例化出一个父窗口
    window = MainWindow(init_window, log, base_path)
    # 设置根窗口默认属性
    window.set_init_window()
    init_window.mainloop()  # 父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示

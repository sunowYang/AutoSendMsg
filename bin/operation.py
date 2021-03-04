# -*- coding: utf-8 -*-

from tkinter import messagebox

import uiautomation


class Operation:
    def __init__(self):
        self.window_name = '发送短信'
        self.window = uiautomation.WindowControl(Name=self.window_name)
        self.is_window_alive = False
        if self.window.Exists():
            self.ctrl_receiver = self.window.GetChildren()[2]
            self.ctrl_msg = self.window.GetLastChildControl()
            self.ctrl_list = self.window.ListControl(foundIndex=1)
            self.is_window_alive = True
        else:
            raise IOError('未找到发送短信窗口')

    def set_topmost(self):
        """
        置顶窗口
        :return:
        """
        if self.is_window_alive:
            self.window.SetTopmost()
            print('set topmost successfully')
        else:
            return False

    def click_receiver(self):
        """
        点击收件人控件
        :return:
        """
        self.ctrl_receiver.Click()
        if self.ctrl_receiver.Exists():
            self.ctrl_receiver.Click()
            return True
        else:
            messagebox.showinfo(title='Warn', message='未找到接收人控件')
            return False

    def set_receiver(self, phone_num):
        """
        填写电话号码
        :param phone_num:
        :return:
        """
        print('set_receiver %s' % phone_num)
        print('set_receiver %s' % str(phone_num))
        self.ctrl_receiver.SendKeys('{Ctrl}a')
        self.ctrl_receiver.SendKeys(str(phone_num))

    def get_receiver(self):
        """
        获取当前收件人
        :return:
        """
        return self.ctrl_receiver.GetLegacyIAccessiblePattern().Value

    def set_msg(self, msg):
        """
        填写发送信息
        :param msg:
        :return:
        """
        self.ctrl_msg.SendKeys('{Ctrl}a')
        self.ctrl_msg.SendKeys(msg)

    def get_msg(self):
        """
        获取当前发送的信息
        :return:
        """
        return self.ctrl_msg.GetLegacyIAccessiblePattern().Value

    def get_msg_num(self):
        """
        获取已发送消息数量
        :return:
        """
        return len(self.ctrl_list.GetChildren())

    def get_msg_status(self, index):
        """
        获取指定已发送消息的发送状态
        :return:True:send successfully
        """
        msg_count = self.get_msg_num()
        if index and index <= msg_count:
            ctrl_msg = self.ctrl_list.GetChildren()[index-1]
            msg = ctrl_msg.GetChildren()[2].GetLegacyIAccessiblePattern().Name
            print(msg)
            if msg == '正在发送...':
                return False
            else:
                return True
        else:
            raise IOError('消息数量小于指定数量')

    def send_msg(self):
        """
        发送信息
        :return:
        """
        self.ctrl_msg.SendKeys('{Ctrl}{Enter}')


if __name__ == '__main__':
    operation = Operation()
    operation.set_topmost()
    operation.set_receiver('1008611')
    operation.set_msg('查积分')
    print(operation.get_msg_num())
    operation.send_msg()
    print(operation.get_msg_num())

from win32com.client import Dispatch
from tkinter import *
import tkinter as tk
from PIL import Image
from PIL import ImageTk
import os
import re
import random
from threading import Thread
import pythoncom
import time

stu_path = "名单.txt"  # 学生名单路径


def speaker(str):
    """
    语音播报
    :param str: 需要播放语音的文字
    """
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(str)


class Rollllcall():
    def __init__(self):
        self.win = Tk()
        self.win.title("Python课堂点名器")
        self.win.iconbitmap("image/icon.ico")
        self.win.geometry("750x450")
        self.win.resizable(False, False)  # 不允许放大窗口，避免放大导致布局变形带来的麻烦
        self.start = False  # 开始按钮的状态
        # 增加背景图片
        img = Image.open('image/back.jpg')
        img = ImageTk.PhotoImage(img, size=(650, 450))
        theLabel = tk.Label(self.win,  # 绑定到一个框架
                            # justify=tk.LEFT,  # 对齐方式
                            image=img,  # 加入图片
                            compound=tk.CENTER,  # 关键:设置为背景图片
                            font=("华文行楷", 20),  # 字体和字号
                            fg="white",
                            )  # 前景色
        theLabel.place(x=0, y=0, relwidth=1, relheight=1)
        self.var = tk.StringVar()  # 储存文字的类
        self.var.set("别紧张")  # 设置文字
        NameLabel = tk.Label(self.win, textvariable=self.var,  # 绑定到一个框架
                             justify=tk.LEFT,  # 对齐方式
                             compound=tk.CENTER,  # 关键:设置为背景图片
                             font=("华文行楷", 35),  # 字体和字号
                             fg="SeaGreen",
                             width=10,
                             )  # 前景色
        NameLabel.place(x=280, y=100)

        # 多选框
        self.checkVar = IntVar()
        Checkbutton(self.win, text="语音播放", variable=self.checkVar,
                    onvalue=1, offvalue=0, height=0, width=0).place(x=170, y=410)
        tk.Button(self.win, text='编辑学生名单', height=0, width=0, command=self.pop_win).place(x=520, y=408)

        self.theButton = tk.Button(self.win, text="开始", font=("华文行楷", 13), fg="SeaGreen", width=20,
                                   command=self.callback)
        self.theButton.place(x=300, y=360)  # 调整按钮的位置
        self.win.mainloop()

    def save_names(self, pop, t):
        """
        保存名单内容
        :param win: #弹出窗
        :param t: 文本框对象

        """
        names = t.get(0.0, "end")
        if re.search("，", names):
            textlabel = tk.Label(pop, text="注意:名单不能使用中文逗号分隔", font=("华文行楷", 12),  # 字体和字号
                                 fg="red", )
            textlabel.place(y=190, x=10)
        else:
            with open(stu_path, "w", encoding="utf-8") as f:
                f.write(names)
            pop.destroy()

    # 编辑学生姓名
    def pop_win(self):
        pop = Tk(className='学生名单编辑')  # 弹出框框名
        pop.geometry('450x250')  # 设置弹出框的大小 w x h
        pop.iconbitmap("image/icon.ico")
        pop.resizable(False, False)

        # 用来编辑名单的文本框
        t = tk.Text(pop, width=61, height='10')
        t.place(x=10, y=10)
        # 判断文件存不存在
        result = os.path.exists(stu_path)
        if result:
            # 存在
            with open(stu_path, "r", encoding='utf-8') as f:
                names = f.read().strip("\n\r\t")
                t.insert("end", names)

        textlabel = tk.Label(pop, text="学生名单请以,(英文状态)的逗号分隔：\n如：刘亦菲,周迅", font=("华文行楷", 12),  # 字体和字号
                             fg="SeaGreen", )
        textlabel.place(y=150, x=10)

        # 点击确定保存数据
        tk.Button(pop, text='确定', height=0, width=0, command=lambda: self.save_names(pop, t)).place(y=200, x=340)
        tk.Button(pop, text='取消', height=0, width=0, command=pop.destroy).place(y=200, x=400)

    def callback(self):
        # 改变开始按钮的状态
        self.start = False if self.start else True
        # 开始随机名单之后修改按钮上的文字
        self.theButton["text"] = "就你了"
        # 开启一个子线程去做操作
        self.t = Thread(target=self.mod_stu_name, args=(self.var, self.checkVar))
        self.t.start()

    def mod_stu_name(self, var, checkVar):
        # 随机读取名单中的一个
        pythoncom.CoInitialize()  # 子线程中调用win32com 语音播放需要设置这一行
        if not os.path.exists(stu_path):
            var.set("请添加名单")
            return None
        with open(stu_path, "r", encoding="utf-8") as f:
            names = f.read().strip("\n\t\r,")
        if not names:
            var.set("请添加名单")
            return None
        name_list = names.split(",")

        random_name = ""
        while self.start:
            random_name = random.choice(name_list)
            var.set(random_name)  # 设置名字随机出现
            time.sleep(0.1)
        self.theButton["text"] = "开始"  # 选中之后将按钮重新修改成 开始
        # 语音播报
        if checkVar.get() == 1:
            speaker(random_name)


if __name__ == '__main__':
    Rollllcall()

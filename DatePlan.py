# _*_ coding:utf-8 _*_
# Project: 
# FileName: DatePlan.py
# UserName: 高俊佶
# ComputerUser：19305
# Day: 2020/6/12
# Time: 12:15
# IDE: PyCharm
# 最爱洪洪，永无BUG！

import os
import sys
import time
import ctypes
import random
import getpass
import win32api
import calendar
import _tkinter
import threading
import tkinter.ttk
from itertools import chain
import win32com.client as client
import tkinter.messagebox as msg
import tkinter.filedialog as file

# while 1:
#     try:
#         from playsound import playsound as play
#         break
#     except ImportError:
#         os.system('pip install -i https://pypi.tuna.tsinghua.edu.cn/simple playsound')
#         print('正在自动安装 playsound 库，按下确定后稍等就好~')


class DatePlan(tkinter.Tk):
    def __init__(self):
        super().__init__(className='任务计划器v6.6')
        now = time.localtime(time.time())
        self.year = now.tm_year
        self.month = now.tm_mon
        self.day = now.tm_mday
        self.hour = now.tm_hour
        self.minute = now.tm_min
        self.titles = ['一', '二', '三', '四', '五', '六', '日']
        self.days = [['一', '二', '三', '四', '五', '六', '日']] + [[d[0] for d in w] for w in calendar.TextCalendar().monthdays2calendar(self.year, self.month)]
        while len(self.days) < 7:
            self.days += [[0] * 7]
        self.months = [m for m in range(1, 13)]
        self.mon = tkinter.IntVar()
        self.yea = tkinter.IntVar()
        self.daily = tkinter.IntVar()
        self.weekly = tkinter.IntVar()
        self.start = tkinter.IntVar()
        self.window = tkinter.IntVar()
        self.dei = tkinter.IntVar()
        if self.task_init()[0]:
            self.window.set(1)
        self.set1 = tkinter.IntVar()
        self.set2 = tkinter.IntVar()
        self.set3 = tkinter.IntVar()
        self.set4 = tkinter.IntVar()
        path = f'C:/Users/{getpass.getuser()}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup'
        if os.path.isfile(f'{path}/DatePlan.exe.lnk'):
            self.start.set(1)
            self.top()
        self.days_index = -1
        self.months_index = -1
        self.days_obj = []
        self.months_obj = []
        self.events_obj = {}
        self.index = None
        self.q = False
        self.music = False
        self.player = True
        self.flag = True
        self.de = True
        self.tab = tkinter.ttk.Notebook(self)
        self.frame3 = tkinter.Frame(self.tab)
        self.frame4 = tkinter.Frame(self.tab)
        self.tab.add(self.frame3, text='日期设置')
        self.tab.add(self.frame4, text='日程管理')
        self.tab.select(self.frame3)
        self.tab.pack(expand=1, fill="both")

        self.frame1 = tkinter.Frame(self.frame3, height=310, width=250)
        self.frame1.place(x=0, y=0)
        for r in range(7):
            for c in range(7):
                button = tkinter.Button(self.frame1, text='', activebackground='skyblue')
                button.place(x=7.5+c*35, y=7.5+r*35, width=30, height=30)
                self.days_obj.append(button)
        self.frame2 = tkinter.Frame(self.frame3, height=130, width=210)
        self.frame2.place(x=250, y=0)
        for r in range(3):
            for c in range(4):
                button = tkinter.Button(self.frame2, text='', activebackground='skyblue')
                button.place(x=5+c*40, y=5+r*40, width=40, height=40)
                self.months_obj.append(button)
        self.move(now.tm_mon)
        self.action(now.tm_mday)
        self.date_calc(self.year, self.month)
        self.box = tkinter.ttk.Combobox(self.frame3, values=[y for y in range(now.tm_year, now.tm_year+101)], state='readonly')
        self.box.current(0)
        self.box.place(x=255, y=130, height=20, width=160)
        self.checkbutton1 = tkinter.Checkbutton(self.frame3, text='启用月管理', onvalue=1, offvalue=0, variable=self.mon, command=lambda: self.check('m'))
        self.checkbutton1.place(x=250, y=155, height=20, width=80)
        self.checkbutton2 = tkinter.Checkbutton(self.frame3, text='启用日重复', onvalue=1, offvalue=0, variable=self.daily, command=lambda: self.check('d'))
        self.checkbutton2.place(x=335, y=155, height=20, width=80)
        self.checkbutton3 = tkinter.Checkbutton(self.frame3, text='启用年管理', onvalue=1, offvalue=0, variable=self.yea, command=lambda: self.check('y'))
        self.checkbutton3.place(x=250, y=175, height=20, width=80)
        self.checkbutton4 = tkinter.Checkbutton(self.frame3, text='启用周重复', onvalue=1, offvalue=0, variable=self.weekly, command=lambda: self.check('w'))
        self.checkbutton4.place(x=335, y=175, height=20, width=80)
        self.checkbutton5 = tkinter.Checkbutton(self.frame3, text='仅按键唤醒', onvalue=1, offvalue=0, variable=self.dei, command=lambda: self.check('o'))
        self.checkbutton5.place(x=250, y=195, height=20, width=80)
        self.checkbutton5 = tkinter.Checkbutton(self.frame3, text='开机自启', onvalue=1, offvalue=0, variable=self.start, command=lambda: self.check('s'))
        self.checkbutton5.place(x=335, y=195, height=20, width=80)
        self.checkbutton4 = tkinter.Checkbutton(self.frame3, text='窗口置顶', onvalue=1, offvalue=0, variable=self.window, command=self.top)
        self.checkbutton4.place(x=335, y=220, height=25, width=80)
        self.button1 = tkinter.Button(self.frame3, text='回到今日', bg='yellow', command=self.back)
        self.button1.place(x=255, y=220, height=25, width=80)
        self.label1 = tkinter.Label(self.frame3, text='                  时                  分，消息：', justify=tkinter.LEFT)
        self.label1.place(x=0, y=260, height=25, width=220)
        self.entry1 = tkinter.Entry(self.frame3)
        self.entry1.place(x=5, y=260, height=25, width=60)
        self.entry2 = tkinter.Entry(self.frame3)
        self.entry2.place(x=90, y=260, height=25, width=60)
        self.entry3 = tkinter.Entry(self.frame3)
        self.entry3.place(x=220, y=260, height=25, width=130)
        self.button2 = tkinter.Button(self.frame3, text='设置', bg='skyblue', command=self.set)
        self.button2.place(x=360, y=260, height=25, width=50)
        self.label2 = tkinter.Label(self.frame3, text='提醒音量(偶数)：', justify=tkinter.LEFT)
        self.label2.place(x=0, y=290, height=20, width=100)
        self.entry4 = tkinter.Entry(self.frame3)
        self.entry4.place(x=100, y=290, height=20, width=40)
        self.button3 = tkinter.Button(self.frame3, text='选择音频文件（英文路径+英文名称）', bg='skyblue', command=self.mp3)
        self.button3.place(x=150, y=290, height=20, width=250)
        self.threads(self.target)
        self.box.bind('<<ComboboxSelected>>', self.change)

        self.button4 = tkinter.Button(self.frame4, text='读取', bg='skyblue', command=self.read)
        self.button4.place(x=5, y=5, height=25, width=50)
        self.button5 = tkinter.Button(self.frame4, text='写入', bg='skyblue', command=self.write)
        self.button5.place(x=65, y=5, height=25, width=50)
        self.scrollbar = tkinter.Scrollbar(self.frame4)
        self.scrollbar.place(x=115, y=35, height=270, width=10)
        self.listbox = tkinter.Listbox(self.frame4, yscrollcommand=self.scrollbar.set)
        self.listbox.place(x=5, y=35, height=270, width=110)
        self.scrollbar.config(command=self.listbox.yview)
        self.listbox.bind('<Double-1>', self.contact_method)
        self.label3 = tkinter.Label(self.frame4, text="文字消息：\n\n月管理：\n\n年管理：\n\n日重复：\n\n年重复：\n\n音量（偶数）：\n\n音频文件：")
        self.label3.place(x=130, y=5, height=250, width=80)
        self.entry5 = tkinter.Entry(self.frame4)
        self.entry5.place(x=200, y=15, height=20, width=200)
        self.checkbutton5 = tkinter.Checkbutton(self.frame4, text='启用月管理', onvalue=1, offvalue=0, variable=self.set2, command=lambda: self.check('m'))
        self.checkbutton5.place(x=260, y=50, height=20, width=80)
        self.checkbutton6 = tkinter.Checkbutton(self.frame4, text='启用日重复', onvalue=1, offvalue=0, variable=self.set4, command=lambda: self.check('d'))
        self.checkbutton6.place(x=260, y=85, height=20, width=80)
        self.checkbutton7 = tkinter.Checkbutton(self.frame4, text='启用年管理', onvalue=1, offvalue=0, variable=self.set1, command=lambda: self.check('y'))
        self.checkbutton7.place(x=260, y=120, height=20, width=80)
        self.checkbutton8 = tkinter.Checkbutton(self.frame4, text='启用周重复', onvalue=1, offvalue=0, variable=self.set3, command=lambda: self.check('w'))
        self.checkbutton8.place(x=260, y=155, height=20, width=80)
        self.entry6 = tkinter.Entry(self.frame4)
        self.entry6.place(x=280, y=190, height=20, width=40)
        self.button6 = tkinter.Button(self.frame4, text='选择音频文件（英文路径+英文名称）', bg='skyblue', command=self.mp33)
        self.button6.place(x=200, y=225, height=20, width=200)
        self.button7 = tkinter.Button(self.frame4, text='删除日程', bg='red', command=self.delete_data)
        self.button7.place(x=160, y=260, height=40, width=200)

        self.tab.bind("<Map>", self.map)
        self.tab.bind("<Unmap>", self.unmap)

        try:
            self.protocol("WM_DELETE_WINDOW", self.close)
        except _tkinter.TclError:
            pass

        self.resizable(0, 0)
        self.mainloop()

    def action(self, d: (str, int)):
        # now = time.localtime(time.time())
        self.day = d
        self.days = [['一', '二', '三', '四', '五', '六', '日']] + [[d[0] for d in w] for w in calendar.TextCalendar().monthdays2calendar(self.year, self.month)]
        self.days_index = list(chain.from_iterable(self.days)).index(d)
        self.date_calc(self.year, self.month)

    def back(self):
        now = time.localtime(time.time())
        self.year = now.tm_year
        self.month = now.tm_mon
        self.day = now.tm_mday
        self.days = [['一', '二', '三', '四', '五', '六', '日']] + [[d[0] for d in w] for w in calendar.TextCalendar().monthdays2calendar(self.year, self.month)]
        self.days_index = list(chain.from_iterable(self.days)).index(self.day)
        self.months_index = self.months.index(self.month)
        self.box.current(0)
        self.date_calc(self.year, self.month)

    def change(self, e):
        now = time.localtime(time.time())
        if e:
            self.days_index = -1
            self.months_index = -1
            self.year = int(self.box.get())
            if now.tm_year == self.year:
                self.date_calc(self.year, now.tm_mon)
            else:
                self.move(1)

    def check(self, t):
        if t == 'y':
            if self.yea.get():
                self.mon.set(0)
                self.daily.set(0)
                self.weekly.set(0)
            if self.set1.get():
                self.set2.set(0)
                self.set3.set(0)
                self.set4.set(0)
        elif t == 'm':
            if self.mon.get():
                self.yea.set(0)
                self.daily.set(0)
                self.weekly.set(0)
            if self.set2.get():
                self.set1.set(0)
                self.set3.set(0)
                self.set4.set(0)
        elif t == 'w':
            if self.weekly.get():
                self.yea.set(0)
                self.mon.set(0)
                self.daily.set(0)
            if self.set3.get():
                self.set1.set(0)
                self.set2.set(0)
                self.set4.set(0)
        elif t == 'd':
            if self.daily.get():
                self.yea.set(0)
                self.mon.set(0)
                self.weekly.set(0)
            if self.set4.get():
                self.set1.set(0)
                self.set2.set(0)
                self.set3.set(0)
        elif t == 'o':
            self.flag = True
        elif t == 's':
            # 全部启动路径：C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp
            # 当前用户启动路径：C:\Users\19305\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
            path = f'C:/Users/{getpass.getuser()}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup'
            # file_name = f'{path}/y{self.year}-m{self.month}-d{self.day}-h{self.hour}-m{self.minute}.py'
            if self.start.get():
                # file_name = str(__file__).replace("\\", '/').split("/")
                # open(f'{"/".join(file_name[:-1]+["start.bat"])}', 'w', encoding='utf-8').write(f'pythonw {__file__}')
                # open(f'start.bat', 'w', encoding='utf-8').write(f'for /F %p in (\'pwd\') do pythonw %p\\{file_name[-1]}')
                # open(f'{path}/s.vbs', 'w', encoding='utf-8').write(f'DIM objShell\nset objShell=wscript.createObject("wscript.shell")\niReturn=objShell.Run("cmd.exe /C {"/".join(file_name[:-1])}/start.bat", 0, TRUE)')
                # 创建快捷方式
                shell = client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(f'{path}/DatePlan.exe.lnk')
                shortcut.TargetPath = sys.argv[0]
                shortcut.save()
                # 复制文件
                # open(f'{path}/DatePlan.pyw', 'w', encoding='utf-8').write(open(__file__, 'r', encoding='utf-8').read())
                # open(f'{path}/task.json', 'w', encoding='utf-8').write(open('/'.join(str(__file__).replace("\\), '/.split('/')[:-1]+['task.json']), 'r', encoding='utf-8').read())
            else:
                # if os.path.isfile(f'start.bat'):
                #     os.remove(f'start.bat')
                # if os.path.isfile(f'{path}/s.vbs'):
                #     os.remove(f'{path}/s.vbs')
                # 删除快捷方式
                if os.path.isfile(f'{path}/DatePlan.exe.lnk'):
                    os.remove(f'{path}/DatePlan.exe.lnk')
                # 删除文件
                # if os.path.isfile(f'{path}/DatePlan.pyw'):
                #     os.remove(f'{path}/DatePlan.pyw')
                # if os.path.isfile(f'{path}/task.json'):
                #     os.remove(f'{path}/task.json')

    def close(self):
        dic = self.task_init()
        now = time.localtime(time.time())
        for num in dic:
            if num and ('-'.join(dic[num]['time'].split('-')[:3]) == f'{now.tm_year}-{now.tm_mon}-{now.tm_mday}' or dic[num]['daily']):
                if not tkinter.messagebox.askyesno(title='提示', message='今天你有计划任务哦，是否继续关闭？'):
                    return
        self.destroy()
        self.quit()

    def contact_method(self, event):
        if event:
            self.index = self.listbox.get(self.listbox.curselection())
            self.entry5.delete(0, tkinter.END)
            self.entry5.insert(tkinter.END, self.events_obj[self.index]['task_msg'])
            self.set1.set(self.events_obj[self.index]['yea_con'])
            self.set2.set(self.events_obj[self.index]['mon_con'])
            self.set3.set(self.events_obj[self.index]['weekly'])
            self.set4.set(self.events_obj[self.index]['daily'])
            self.entry6.delete(0, tkinter.END)
            self.entry6.insert(tkinter.END, self.events_obj[self.index]['volume'])
            if self.events_obj[self.index]['music']:
                self.button6.configure(text=self.events_obj[self.index]['music'], bg='green')
            else:
                self.button6.configure(text='选择音频文件（英文路径+英文名称）', bg='skyblue')

    def date_calc(self, y: int, m: int):
        now = time.localtime(time.time())
        # self.days = [['一', '二', '三', '四', '五', '六', '日']]
        # first = True
        # for r in calendar.month(y, m).strip('\n').replace("\\), '/.split('\n')[2:]:
        #     week = []
        #     if first:
        #         for c in r.replace("\\), '/.split('  '):
        #             week.append(c.strip())
        #         first = False
        #         while len(week) > 7:
        #             week = week[1:]
        #     else:
        #         for c in r.replace("\\), '/.split(' '):
        #             if c:
        #                 week.append(c.strip())
        #         while len(week) < 7:
        #             week.append('')
        #     self.days.append(week)
        self.days = [['一', '二', '三', '四', '五', '六', '日']] + [[d[0] for d in w] for w in calendar.TextCalendar().monthdays2calendar(y, m)]
        while len(self.days) < 7:
            self.days += [[0] * 7]
        # for widget in self.frame1.winfo_children():
        #     widget.destroy()
        # for widget in self.frame2.winfo_children():
        #     widget.destroy()
        for r in range(7):
            for c in range(7):
                if self.days[r][c]:
                    self.days_obj[7*r+c].configure(text=str(self.days[r][c]), command=lambda d=self.days[r][c]: self.action(d))
                    if self.days_index == 7*r+c:
                        self.days_obj[7*r+c].configure(bg='green', state=tkinter.NORMAL)
                    elif now.tm_year == self.year and now.tm_mon == self.month and self.days[r][c] == now.tm_mday:
                        self.days_obj[7*r+c].configure(bg='yellow', state=tkinter.NORMAL)
                    else:
                        self.days_obj[7*r+c].configure(bg='white', state=tkinter.NORMAL)
                else:
                    self.days_obj[7*r+c].configure(text='', bg='white', state=tkinter.DISABLED)
        for r in range(3):
            for c in range(4):
                self.months_obj[4*r+c].configure(text=str(self.months[4*r+c]), command=lambda d=self.months[4*r+c]: self.move(d))
                if self.months_index == 4*r+c:
                    self.months_obj[4*r+c].configure(bg='green')
                elif now.tm_year == self.year and now.tm_mon == self.months[4*r+c]:
                    self.months_obj[4*r+c].configure(bg='yellow')
                else:
                    self.months_obj[4*r+c].configure(bg='white')

    def delete_data(self):
        if self.index and msg.askyesno(title='注意！', message=f'您正在删除日程 {self.index}，是否继续？'):
            all_data = self.task_init()
            add = False
            for n in range(1, len(all_data)):
                if n in all_data and (all_data[n]['time'] == self.index or add):
                    if n+1 in all_data:
                        all_data[n] = all_data[n+1]
                    add = True
            del all_data[len(all_data)-1]
            self.entry5.delete(0, tkinter.END)
            self.set1.set(0)
            self.set2.set(0)
            self.set3.set(0)
            self.set4.set(0)
            self.entry6.delete(0, tkinter.END)
            self.button6.configure(text='选择音频文件（英文路径+英文名称）', bg='skyblue')
            open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write(str(all_data))
            self.read()

    @staticmethod
    def get_week_day(y: int, m: int, d: int):
        data = []
        for w in iter([(d[0], d[1]) for d in w] for w in calendar.TextCalendar().monthdays2calendar(y, m)):
            data += w
        date = {d[0]: d[1] for d in data}
        return date[d]

    def key(self):
        while 1:
            if win32api.GetAsyncKeyState(118):
                self.map(1)
                break
            time.sleep(1)

    def map(self, e):
        if e and self.de:
            # print('最大化')
            self.overrideredirect(0)
            self.title('任务计划器v6.6')
            self.geometry(f'420x340+{int((self.winfo_screenwidth()-420-16)/2)}+{int((self.winfo_screenheight()-340-32)/2)}')
            self.deiconify()
        else:
            self.de = True

    def move(self, m: int):
        # now = time.localtime(time.time())
        self.month = m
        self.months_index = self.months.index(m)
        self.days_index = -1
        self.date_calc(self.year, self.month)

    def mp3(self):
        file_name = file.askopenfilename(filetypes=[('MP3', '*.mp3'), ('WAVE', '*.wav')])
        if file_name:
            self.music = file_name
            self.threads(self.mp3_play, self.music, self.entry4.get() and int(self.entry4.get()) or 20)
            self.button3.configure(text=self.music, bg='green')
            if str(msg.showinfo(title='测试音频中...', message='点击 OK 结束测试。')) == 'ok':
                self.player = False
        else:
            self.music = False
            self.button3.configure(text='选择音频文件（英文路径+英文名称）', bg='skyblue')

    def mp33(self):
        file_name = file.askopenfilename(filetypes=[('MP3', '*.mp3'), ('WAVE', '*.wav')])
        if file_name:
            self.music = file_name
            self.threads(self.mp3_play, self.music, self.entry6.get() and int(self.entry6.get()) or 20)
            self.button6.configure(text=self.music, bg='green')
            if str(msg.showinfo(title='测试音频中...', message='点击 OK 结束测试。')) == 'ok':
                self.player = False
        else:
            self.music = False
            self.button6.configure(text='选择音频文件（英文路径+英文名称）', bg='skyblue')

    def mp3_play(self, music: (str, bool), vol: int = 20):
        if music:
            volume = int(vol / 2)
            WM_APPCOMMAND = 0x319
            APPCOMMAND_VOLUME_UP = 0x0a
            APPCOMMAND_VOLUME_DOWN = 0x09
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            for i in range(50):
                ctypes.windll.user32.PostMessageA(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_VOLUME_DOWN * 0x10000)
            for i in range(volume):
                ctypes.windll.user32.PostMessageA(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_VOLUME_UP * 0x10000)
            try:
                self.playsound(music)
            except FileNotFoundError:
                msg.showerror(title='错误！', message=f'音频文件 {music} 损坏或不存在！')

    def plan(self, **t):
        while not self.q:
            now = time.localtime(time.time())
            times = t['time'].split('-')
            # print(times)
            if t['daily']:
                if now.tm_hour == int(times[3]) and now.tm_min == int(times[4]):
                    self.threads(self.mp3_play, t['music'], t['volume'])
                    self.wm_attributes('-topmost', 1)
                    if str(msg.showinfo(title='时间到！', message=t['task_msg'])) == 'ok':
                        self.player = False
                    self.wm_attributes('-topmost', self.window.get())
                    break
            elif t['weekly']:
                if now.tm_wday == self.get_week_day(now.tm_year, now.tm_mon, int(times[2])) and now.tm_hour == int(times[3]) and now.tm_min == int(times[4]):
                    self.threads(self.mp3_play, t['music'], t['volume'])
                    self.wm_attributes('-topmost', 1)
                    if str(msg.showinfo(title='时间到！', message=t['task_msg'])) == 'ok':
                        self.player = False
                    self.wm_attributes('-topmost', self.window.get())
                    break
            else:
                if times[2].isdigit():
                    if t['mon_con']:
                        if now.tm_mday == int(times[2]) and now.tm_hour == int(times[3]) and now.tm_min == int(times[4]):
                            self.threads(self.mp3_play, t['music'], t['volume'])
                            self.wm_attributes('-topmost', 1)
                            if str(msg.showinfo(title='时间到！', message=t['task_msg'])) == 'ok':
                                self.player = False
                            self.wm_attributes('-topmost', self.window.get())
                            break
                    elif t['yea_con']:
                        if now.tm_mon == int(times[1]) and now.tm_mday == int(times[2]) and now.tm_hour == int(times[3]) and now.tm_min == int(times[4]):
                            self.threads(self.mp3_play, t['music'], t['volume'])
                            self.wm_attributes('-topmost', 1)
                            if str(msg.showinfo(title='时间到！', message=t['task_msg'])) == 'ok':
                                self.player = False
                            self.wm_attributes('-topmost', self.window.get())
                            break
                    else:
                        if now.tm_year == int(times[0]) and now.tm_mon == int(times[1]) and now.tm_mday == int(times[2]) and now.tm_hour == int(times[3]) and now.tm_min == int(times[4]):
                            self.threads(self.mp3_play, t['music'], t['volume'])
                            self.wm_attributes('-topmost', 1)
                            if str(msg.showinfo(title='时间到！', message=t['task_msg'])) == 'ok':
                                self.player = False
                            self.wm_attributes('-topmost', self.window.get())
                            break
                else:
                    if now.tm_year == int(times[0]) and now.tm_mon == int(times[1]) and now.tm_wday == self.titles.index(times[2]) and now.tm_hour == int(times[3]) and now.tm_min == int(times[4]):
                        self.threads(self.mp3_play, t['music'], t['volume'])
                        self.wm_attributes('-topmost', 1)
                        if str(msg.showinfo(title='时间到！', message=t['task_msg'])) == 'ok':
                            self.player = False
                        self.wm_attributes('-topmost', self.window.get())
                        break
            time.sleep(1)

    def playsound(self, sound, block=False):
        WM_APPCOMMAND = 0x319
        APPCOMMAND_VOLUME_MUTE = 0x08
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        try:
            alias = 'playsound_' + str(random.random())
            self.win_command('open "' + sound + '" alias', alias)
            self.win_command('set', alias, 'time format milliseconds')
            durationInMS = self.win_command('status', alias, 'length')
            self.win_command('play', alias, 'from 0 to', durationInMS.decode())
            if block:
                time.sleep(float(durationInMS) / 1000.0)
            while 1:
                if not self.player:
                    self.player = True
                    self.win_command('stop', alias)
                    ctypes.windll.user32.PostMessageA(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_VOLUME_MUTE * 0x10000)
                    break
        except UnicodeDecodeError:
            msg.showerror(title='错误！', message='路径或文件名可能含有中文, 或是文件的实际格式不是mp3或mav！')
        except OSError:
            msg.showerror(title='错误！', message='路径或文件名可能含有中文, 或是文件的实际格式不是mp3或mav！')

    def read(self):
        self.listbox.delete(0, tkinter.END)
        json = self.task_init()
        self.entry5.delete(0, tkinter.END)
        self.set1.set(0)
        self.set2.set(0)
        self.set3.set(0)
        self.set4.set(0)
        self.entry6.delete(0, tkinter.END)
        self.button6.configure(text='选择音频文件（英文路径+英文名称）', bg='skyblue')
        self.index = 0
        for n in range(1, len(json)):
            self.events_obj[json[n]['time']] = json[n]
            self.listbox.insert(tkinter.END, json[n]['time'])

    def set(self):
        now = time.localtime(time.time())
        if self.entry1.get() and self.entry2.get():
            try:
                self.hour = int(self.entry1.get().replace(' ', ''))
                self.minute = int(self.entry2.get().replace(' ', ''))
                if not (0 <= self.hour <= 24 and 0 <= self.minute <= 59):
                    msg.showerror(title='警告！', message='请输入正确的数字！')
                else:
                    text = f'计划任务来喽！现在的时间是：{self.year}年{self.month}月{self.day}日{self.hour}时{self.minute}分。'
                    if self.entry3.get():
                        text = self.entry3.get()
                    all_data = self.task_init()
                    for n in range(1, len(all_data)):
                        if all_data[n]['time'] == f'{self.year}-{self.month}-{self.day}-{self.hour}-{self.minute}':
                            del all_data[n]
                    data = {'time': f'{self.year}-{self.month}-{self.day}-{self.hour}-{self.minute}', 'task_msg': text, 'mon_con': self.mon.get(), 'yea_con': self.yea.get(),
                            'made_time': f'{now.tm_year}-{now.tm_mon}-{now.tm_mday}-{now.tm_hour}-{now.tm_min}', 'daily': self.daily.get(), 'weekly': self.weekly.get(),
                            'music': self.music, 'volume': str(self.entry4.get()).isdigit() and int(self.entry4.get()) or 20}
                    n = 1
                    while n in all_data:
                        n += 1
                    all_data[n] = data
                    open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write(str(all_data))
            except ValueError:
                msg.showerror(title='警告！', message='请输入正确的内容！')
        else:
            msg.showwarning(title='注意！', message='请输入时间分钟！')

    def target(self):
        old = None
        while 1:
            tasks = self.task_init()
            tasks_time_list = []
            for n in tasks:
                if n:
                    tasks_time_list.append(tasks[n])
            tasks_time_list = sorted(tasks_time_list, key=lambda i: (i['time'], i['made_time']))
            if old != tasks:
                old = tasks
                self.q = True
                time.sleep(2)
                self.q = False
                if tasks_time_list:
                    for n in range(len(tasks_time_list)):
                        self.threads(self.plan, **tasks_time_list[n])
            time.sleep(30)

    @staticmethod
    def task_init():
        try:
            task = eval(open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'r', encoding='utf-8').read())
            if 0 not in task:
                # temp = task
                # task = {0: False}
                # for i in temp:
                #     task[i] = temp[i]
                task = {k: v for k, v in ((0, False),) + tuple(task.items())}
                open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write(str(task))
            return task
        except SyntaxError:
            return {0: False}

    @staticmethod
    def threads(t, *a, **k):
        thread = threading.Thread(target=t, args=a, kwargs=k)
        thread.daemon = True
        thread.start()

    def top(self):
        # self.wm_attributes('-topmost', self.window.get())
        all_data = self.task_init()
        if self.window.get():
            self.wm_attributes('-topmost', 1)
            all_data[0] = True
        else:
            self.wm_attributes('-topmost', 0)
            all_data[0] = False
        open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write(str(all_data))

    def unmap(self, e):
        if self.flag:
            if self.dei.get():
                msg.showinfo(title='提示', message='恢复窗口请按 F7')
            else:
                msg.showinfo(title='提示', message='恢复窗口请在左下角找到标题并双击（可以拖动）或者按 F7')
        self.flag = False
        if e:
            # print('最小化')
            if self.dei.get():
                self.de = False
                self.deiconify()  # 启用此命令将不会出现左下角的双击最大化，只能通过热键激活窗口！
            self.title('双击最大化')
            self.overrideredirect(1)
            self.geometry('0x0')
            self.threads(self.key)

    @staticmethod
    def win_command(*command):
        buf = ctypes.c_buffer(255)
        command = ' '.join(command).encode(sys.getfilesystemencoding())
        errorCode = int(ctypes.windll.winmm.mciSendStringA(command, buf, 254, 0))
        if errorCode:
            errorBuffer = ctypes.c_buffer(255)
            ctypes.windll.winmm.mciGetErrorStringA(errorCode, errorBuffer, 254)
            exceptionMessage = ('\n    Error ' + str(errorCode) + ' for command:'
                                                                  '\n        ' + command.decode() +
                                '\n    ' + errorBuffer.value.decode())
            raise exceptionMessage
        return buf.value

    def write(self):
        if self.index:
            now = time.localtime(time.time())
            all_data = self.task_init()
            text = self.entry5.get() and self.entry5.get() or '计划任务来喽！现在的时间是：{}年{}月{}日{}时{}分。'.format(*self.index.split('-'))
            for n in range(1, len(all_data)):
                if all_data[n]['time'] == self.index:
                    all_data[n]['task_msg'] = text
                    all_data[n]['yea_con'] = self.set1.get()
                    all_data[n]['mon_con'] = self.set2.get()
                    all_data[n]['weekly'] = self.set3.get()
                    all_data[n]['daily'] = self.set4.get()
                    all_data[n]['music'] = self.music
                    all_data[n]['volume'] = str(self.entry6.get()).isdigit() and int(self.entry6.get()) or 20
                    all_data[n]['made_time'] = f'{now.tm_year}-{now.tm_mon}-{now.tm_mday}-{now.tm_hour}-{now.tm_min}'
            open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write(str(all_data))
            msg.showinfo(title='完成！', message='写入成功！')


if __name__ == '__main__':
    init = 0  # 清空计划任务标志变量
    try:
        if not os.path.isfile(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json') or init:
            open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write('{0: False}')  # 清空计划任务 && 初始化文件
    except SyntaxError:
        open(f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\task.json', 'w', encoding='utf-8').write('{0: False}')  # 清空计划任务 && 初始化文件
    if not init:
        DatePlan()

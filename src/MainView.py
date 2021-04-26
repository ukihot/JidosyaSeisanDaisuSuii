#!/usr/bin/env python3
#coding: UTF-8
import csv
import datetime
import io
import logging
import os
import re
import shutil
import sys
import tkinter as tk
import tkinter.ttk as ttk
import traceback
from tkinter import messagebox

import AutomobileProductionScraping
import MST_MAKER_URL


class MainView:

    def __init__(self):
        today = datetime.date.today()
        if today.month >3:
            self.current_month = today.month-2
        else:
            self.current_month = today.month
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.after(0, self.main_proc)
        self.root.mainloop()

    def main_proc(self):
        # ここで設定Windowを Toplevel Widget にて作成
        global top
        self.top = tk.Toplevel(self.root)
        self.top.geometry("300x150")
        self.combo = ttk.Combobox(self.top, state='readonly')
        months = ["2021/01", "2021/02", "2021/03", "2021/04", "2021/05", "2021/06", "2021/07", "2021/08", "2021/09", "2021/10", "2021/11", "2021/12"]
        self.combo["values"] =[months[i] for i in range(self.current_month)]
        self.combo.current(0)
        self.label = tk.Label(self.top, text='自動車生産台数を取得します。\n対象の年月を指定してください。')
        self.button = tk.Button(self.top, text="登録", command=self.select_month) 
        self.label.pack()
        self.combo.pack()
        self.button.pack()
        self.top.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.top.destroy()
        # 確認用ダイアログを出す
        onemore = messagebox.askyesno('確認', '処理を中断しますか？')
        if onemore:
            self.root.destroy()
        else:
            self.root.after(0, self.main_proc)

    def select_month(self):
        month = self.combo.get()
        if (month == '2021/01'):
            month = 1
        elif(month == '2021/02'):
            month = 2
        elif(month == '2021/03'):
            month = 3
        elif(month == '2021/04'):
            month = 4
        elif(month == '2021/05'):
            month = 5
        elif(month == '2021/06'):
            month = 6
        elif(month == '2021/07'):
            month = 7
        elif(month == '2021/08'):
            month = 8
        elif(month == '2021/09'):
            month = 9
        elif(month == '2021/10'):
            month = 10
        elif(month == '2021/11'):
            month = 11
        elif(month == '2021/12'):
            month = 12

        exe = messagebox.askyesno("確認", str(month)+'月の生産台数を取得します。')
        if exe:
            self.top.destroy()
            try:
                AutomobileProductionScraping.AutomobileProductionScraping(month)
            except:
                messagebox.showwarning("警告", 'エラーにより処理を中断しちゃいます。')
                logger = logging.getLogger('Logging')
                file_handler = logging.FileHandler('error.log', mode='a', encoding='utf-8')
                file_handler.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
                logger.addHandler(file_handler)
                logger.exception(traceback.format_exc())
                sys.exit(1)
            
            messagebox.showinfo("確認", '処理が完了しちゃいました。')
            self.top.destroy()
            sys.exit(0)
        else:
            self.root.destroy()
main_view = MainView()

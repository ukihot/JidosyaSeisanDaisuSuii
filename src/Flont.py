#!/usr/bin/env python3
#coding: UTF-8
import csv
import io
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
        self.combo["values"] = ("2021/01", "2021/02", "2021/03", "2021/04", "2021/05", "2021/06", "2021/07", "2021/08", "2021/09", "2021/10", "2021/11", "2021/12")
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
            # 再度 mainloop内から main_proc() 関数を呼ぶ
            self.root.after(0, self.main_proc)
        else:
            self.root.destroy()

    def select_month(self):
        month = self.combo.get()
        if (month == '2021/01'):
            url = MST_MAKER_URL.jan_url
            month = 1
        elif(month == '2021/02'):
            url = MST_MAKER_URL.feb_url
            month = 2
        elif(month == '2021/03'):
            url = MST_MAKER_URL.mar_url
            month = 3
        elif(month == '2021/04'):
            url = MST_MAKER_URL.apr_url
            month = 4
        elif(month == '2021/05'):
            url = MST_MAKER_URL.may_url
            month = 5
        elif(month == '2021/06'):
            url = MST_MAKER_URL.jun_url
            month = 6
        elif(month == '2021/07'):
            url = MST_MAKER_URL.jul_url
            month = 7
        elif(month == '2021/08'):
            url = MST_MAKER_URL.aug_url
            month = 8
        elif(month == '2021/09'):
            url = MST_MAKER_URL.sep_url
            month = 9
        elif(month == '2021/10'):
            url = MST_MAKER_URL.oct_url
            month = 10
        elif(month == '2021/11'):
            url = MST_MAKER_URL.nov_url
            month = 11
        elif(month == '2021/12'):
            url = MST_MAKER_URL.dec_url
            month = 12

        exe = messagebox.askyesno("確認", str(month)+'月の生産台数を取得します。')
        if exe:
            self.top.destroy()
            try:
                AutomobileProductionScraping.AutomobileProductionScraping(url,month)
            except:
                messagebox.showwarning("警告", 'エラーにより処理を中断します。')
                traceback.print_exc()
                sys.exit(1)
            
            messagebox.showinfo("確認", '処理が完了しました。')
            self.top.destroy()
            sys.exit(0)
        else:
            self.root.destroy()

main_view = MainView()

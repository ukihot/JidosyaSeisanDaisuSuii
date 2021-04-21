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
import urllib
from tkinter import messagebox
from urllib.request import urlopen

import pandas as pd
import requests
import tabula
import xlrd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

import MST_MAKER_URL

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.76 Safari/537.36'}
values = {'name': 'Michael Foord',
          'location': 'Northampton',
          'language': 'Python' }
data = urllib.parse.urlencode(values).encode('utf-8')

class AutomobileProductionScraping:
    month = None
    url = None
    def __init__(self):
        if not os.path.exists('data'):
            os.mkdir('data')

        self.flont()
        self.update_mst_maker_url()
        self.output_excle(self.aggregation())
        self.root.withdraw()
        messagebox.showinfo("確認", '処理が完了しました。')

    # マスタ値更新
    def update_mst_maker_url(self):
        #　各メーカーの生産台数がわかるページURLを調査する。
        makers =['mzd','mtb','szk']
        for maker in makers:
            req = urllib.request.Request(MST_MAKER_URL.meta_url[maker], data, headers)
            html = urllib.request.urlopen(req)
            soup = BeautifulSoup(html, "html.parser")
            elems = soup.select(MST_MAKER_URL.select[maker][0])
            ls_url = []

            for elem in elems:
                ls_url.append(elem.get('href')) if (extraction := re.search(MST_MAKER_URL.select[maker][self.month], elem.getText())) else None
            if (len(ls_url)==0):
                if (maker == 'szk'):
                    maker_name ='スズキ'
                elif (maker == 'mzd'):
                    maker_name ='マツダ'
                else:
                    maker_name ='三菱'
                warning = tk.Tk()
                warning.withdraw()
                messagebox.showwarning("警告", maker_name +'の'+str(self.month)+'月の情報はまだ非公開です。'+'処理を中断します。')
                sys.exit(1)
            for (index, url) in enumerate(ls_url):
                ls_url[index] = re.search(r'/.*l', url).group(0)
            print(ls_url)

    # データ出力
    def output_excle(self, seisan_daisu):
        output_data = pd.DataFrame({'メーカー': ['トヨタ', 'ダイハツ', 'ホンダ', '日産', 'スズキ', 'マツダ', '三菱', 'スバル', 'いすゞ', '日野', '三菱ふそう'], '1月': seisan_daisu})
        # Excelワークブックの生成
        wb = Workbook()
        ws = wb.active
        ws.title = '取得結果'
        rows = dataframe_to_rows(output_data, index=False, header=True)

        for row_no, row in enumerate(rows, 3):
            for col_no, value in enumerate(row, 2):
                ws.cell(row=row_no, column=col_no, value=value)
        wb.save('./集計結果.xlsx')

    # エクセルスクレイピング
    def xl_scraping(self, url, sheet_name, row, target_row, cell_name, FILEPATH):
        r = requests.get(url, allow_redirects=True)
        open(FILEPATH, 'wb').write(r.content)
        wb = xlrd.open_workbook(FILEPATH)
        st = wb.sheet_by_name(sheet_name)
        for col_j, cell in enumerate(st.row(row)):
            if(cell.value == cell_name):
                return st.cell_value(target_row, col_j)

    # Webスクレイピング
    def web_scraping(self, url, r, c, FILEPATH):
        req = urllib.request.Request(url, data, headers)
        try:
            html = urllib.request.urlopen(req)
            soup = BeautifulSoup(html, "html.parser")
        except urllib.error.HTTPError:
            html = requests.get(url, headers)
            soup = BeautifulSoup(html.content, "html.parser")
        except urllib.error.URLError:
            sys.exit(1)

        # HTMLから表(tableタグ)の部分を全て取得する
        table = soup.find_all("table")
        for tab in table:
            with open(FILEPATH, "w+", encoding='utf-8') as f:
                writer = csv.writer(f)
                rows = tab.find_all("tr")
                for row in rows:
                    csvRow = []
                    for cell in row.findAll(['td', 'th']):
                        csvRow.append(cell.get_text())
                    writer.writerow(csvRow)
            # 1つ目の表のみ取り込むbreak
            break
        with open(FILEPATH, 'r',encoding='utf-8') as f:
            tar = [row for row in csv.reader(f)]
        return tar[r][c].replace(' ', '').replace('\n','')

    def aggregation(self):
        seisan_daisu = [0] * 11
        seisan_daisu[0] = self.xl_scraping(self.url['tyt'], '生産', 2, 6, MST_MAKER_URL.tyt_key[self.month], "data/toyota.xls")
        seisan_daisu[1] = self.xl_scraping(self.url['tyt'], '生産', 2, 13, MST_MAKER_URL.tyt_key[self.month], "data/toyota.xls")
        seisan_daisu[2] = self.xl_scraping(self.url['hnd'], '日本語', 84, 85, MST_MAKER_URL.hnd_key[self.month], "data/honda.xlsx")
        seisan_daisu[4] = self.web_scraping(self.url['szk'], 2, 1, "data/suzuki.csv")
        seisan_daisu[5] = self.web_scraping(self.url['mzd'], 4, 2, "data/mazda.csv")
        seisan_daisu[6] = self.web_scraping(self.url['mtb'], 4, 2, "data/mitsubishi.csv")
        seisan_daisu[8] = self.web_scraping(self.url['isz'], 0, 11, "data/isuzu.csv")
        seisan_daisu[9] = self.xl_scraping(self.url['tyt'], '生産', 2, 20, MST_MAKER_URL.tyt_key[self.month], "data/toyota.xls")
        seisan_daisu[10] =self.web_scraping(self.url['hso'], 0, 11, "data/huso.csv")
        return seisan_daisu

    def flont(self):
        self.root = tk.Tk()
        self.root.geometry("300x150")
        self.combo = ttk.Combobox(self.root, state='readonly')
        self.combo["values"] = ("2021/01", "2021/02", "2021/03", "2021/04", "2021/05", "2021/06", "2021/07", "2021/08", "2021/09", "2021/10", "2021/11", "2021/12")
        self.combo.current(0)
        self.text = tk.StringVar()
        self.text.set("自動車生産台数を取得します。\n対象の年月を指定してください。")
        self.label = tk.Label(self.root, textvariable=self.text)

        self.button = tk.Button(self.root,
                                text="登録",
                                command=self.select_month)  
        self.label.pack()
        self.combo.pack()
        self.button.pack()
        self.root.mainloop()

    def select_month(self):
        month = self.combo.get()
        if (month == '2021/01'):
            self.url = MST_MAKER_URL.jan_url
            self.month = 1
        elif(month == '2021/02'):
            self.url = MST_MAKER_URL.feb_url
            self.month = 2
        elif(month == '2021/03'):
            self.url = MST_MAKER_URL.mar_url
            self.month = 3
        elif(month == '2021/04'):
            self.url = MST_MAKER_URL.apr_url
            self.month = 4
        elif(month == '2021/05'):
            self.url = MST_MAKER_URL.may_url
            self.month = 5
        elif(month == '2021/06'):
            self.url = MST_MAKER_URL.jun_url
            self.month = 6
        elif(month == '2021/07'):
            self.url = MST_MAKER_URL.jul_url
            self.month = 7
        elif(month == '2021/08'):
            self.url = MST_MAKER_URL.aug_url
            self.month = 8
        elif(month == '2021/09'):
            self.url = MST_MAKER_URL.sep_url
            self.month = 9
        elif(month == '2021/10'):
            self.url = MST_MAKER_URL.oct_url
            self.month = 10
        elif(month == '2021/11'):
            self.url = MST_MAKER_URL.nov_url
            self.month = 11
        elif(month == '2021/12'):
            self.url = MST_MAKER_URL.dec_url
            self.month = 12
        
        self.root.withdraw()
        messagebox.showinfo("確認", str(month)+'の生産台数を取得します。')
        self.root.destroy()

# Main
automobileProductionScraping = AutomobileProductionScraping()
shutil.rmtree('data')

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

import chromedriver_binary
import pandas as pd
import requests
import tabula
import xlrd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.support.ui import Select

import MST_MAKER_URL

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.76 Safari/537.36'}
values = {'name': 'Michael Foord',
          'location': 'Northampton',
          'language': 'Python' }
data = urllib.parse.urlencode(values).encode('utf-8')

class AutomobileProductionScraping:

    def __init__(self,url,month):
        if not os.path.exists('data'):
            os.mkdir('data')
        self.update_mst_maker_url(url,month)
        self.output_excle(self.aggregation(url, month))

    # マスタ値更新
    def update_mst_maker_url(self,url ,month):
        #　各メーカーの生産台数がわかるページURLを調査する。
        makers =['szk','mzd','mtb']
        ls_url = []
        for maker in makers:
            ## スズキのURL更新
            if (maker == 'szk'):
                driver = webdriver.Chrome()
                driver.get(MST_MAKER_URL.meta_url[maker])
                dropdown = driver.find_element_by_name('ad')
                select = Select(dropdown)
                select.select_by_value('ad2020')
                elements = driver.find_elements_by_tag_name('a')
                for element in elements:
                    if re.search(MST_MAKER_URL.select[maker][month] ,element.text):
                        ls_url.append(element.get_attribute('href'))
            ## マツダのURL更新
            elif (maker == 'mzd'):
                req = urllib.request.Request(MST_MAKER_URL.meta_url[maker], data, headers)
                html = urllib.request.urlopen(req)
                soup = BeautifulSoup(html, "html.parser")
                elems = soup.select(MST_MAKER_URL.select[maker][0])
                for elem in elems:
                    if re.search(MST_MAKER_URL.select[maker][month], elem.getText()):
                        ls_url.append(elem.get('href'))

            ## TODO:三菱のURL更新
            elif (maker == 'mtb'):
                pass

            print(ls_url)
            if (len(ls_url)==0):
                if (maker == 'szk'):
                    maker_name ='スズキ'
                elif (maker == 'mzd'):
                    maker_name ='マツダ'
                else:
                    maker_name ='三菱'
                warning = tk.Tk()
                warning.withdraw()
                messagebox.showwarning("警告", maker_name +'の'+str(month)+'月の情報はまだ非公開です。'+'処理を中断します。')
                sys.exit(1)
            for (index, url) in enumerate(ls_url):
                ls_url[index] = re.search(r'/.*l', url).group(0)

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

    def aggregation(self, url, month):
        seisan_daisu = [0] * 11
        seisan_daisu[0] = self.xl_scraping(url['tyt'], '生産', 2, 6, MST_MAKER_URL.tyt_key[month], "data/toyota.xls")
        seisan_daisu[1] = self.xl_scraping(url['tyt'], '生産', 2, 13, MST_MAKER_URL.tyt_key[month], "data/toyota.xls")
        seisan_daisu[2] = self.xl_scraping(url['hnd'], '日本語', 84, 85, MST_MAKER_URL.hnd_key[month], "data/honda.xlsx")
        seisan_daisu[4] = self.web_scraping(url['szk'], 2, 1, "data/suzuki.csv")
        seisan_daisu[5] = self.web_scraping(url['mzd'], 4, 2, "data/mazda.csv")
        seisan_daisu[6] = self.web_scraping(url['mtb'], 4, 2, "data/mitsubishi.csv")
        seisan_daisu[8] = self.web_scraping(url['isz'], 0, 11, "data/isuzu.csv")
        seisan_daisu[9] = self.xl_scraping(url['tyt'], '生産', 2, 20, MST_MAKER_URL.tyt_key[month], "data/toyota.xls")
        seisan_daisu[10] =self.web_scraping(url['hso'], 0, 11, "data/huso.csv")
        return seisan_daisu
#!/usr/bin/env python3
#coding: UTF-8
import csv
import io
import os
import re
import shutil
import sys
import urllib
from tkinter import messagebox
from urllib.request import urlopen

import MST_MAKER_URL
import pandas as pd
import requests
import tabula
import xlrd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.76 Safari/537.36'}
values = {'name': 'Michael Foord',
              'location': 'Northampton',
              'language': 'Python' }
data = urllib.parse.urlencode(values).encode('utf-8')
def main():
    messagebox.showinfo('確認', '各メーカーの自動車生産台数を現在のフォルダ下に出力します。')
    if not os.path.exists('data'):
        os.mkdir('data')
    if not os.path.exists('out'):
        os.mkdir('out')
    seisan_daisu = [0] * 11
    seisan_daisu[0] = xl_scraping(MST_MAKER_URL.tyturl, '生産', 2, 6, 202101.0, "data/toyota.xls")
    seisan_daisu[1] = xl_scraping(MST_MAKER_URL.tyturl, '生産', 2, 13, 202101.0, "data/toyota.xls")
    seisan_daisu[2] = xl_scraping(MST_MAKER_URL.hndurl, '日本語', 84, 85, '1月実績', "data/honda.xlsx")
    seisan_daisu[4] = web_scraping(MST_MAKER_URL.szkurl, 2, 1, "data/suzuki.csv")
    seisan_daisu[5] = web_scraping(MST_MAKER_URL.mzdurl, 4, 2, "data/mazda.csv")
    seisan_daisu[6] = web_scraping(MST_MAKER_URL.mtburl, 4, 2, "data/mitsubishi.csv")
    seisan_daisu[8] = web_scraping(MST_MAKER_URL.iszurl, 0, 11, "data/isuzu.csv")
    seisan_daisu[9] = xl_scraping(MST_MAKER_URL.tyturl, '生産', 2, 20, 202101.0, "data/toyota.xls")
    seisan_daisu[10] = web_scraping(MST_MAKER_URL.hsourl, 0, 11, "data/huso.csv")
    output_excle(seisan_daisu)
    shutil.rmtree('data')

# マスタ値更新
def updata_mst_maker_url():
    req = urllib.request.Request(MST_MAKER_URL.prt_mzd_url, data, headers)
    html = urllib.request.urlopen(req)
    soup = BeautifulSoup(html, "html.parser")
    elems = soup.select(MST_MAKER_URL.select_wd_mzd)

    ls_url = []
    for elem in elems:
        ls_url.append(elem.get('href')) if (extraction := re.search('2021年[1-12]月の生産', elem.getText())) else None
    print(ls_url)
    for (index, url) in enumerate(ls_url):
        ls_url[index] = re.search(r'/.*l', url).group(0)

# データ出力
def output_excle(seisan_daisu):
    output_data = pd.DataFrame({'メーカー': ['トヨタ', 'ダイハツ', 'ホンダ', '日産', 'スズキ', 'マツダ', '三菱', 'スバル', 'いすゞ', '日野', '三菱ふそう'], '1月': seisan_daisu})
    # Excelワークブックの生成
    wb = Workbook()
    ws = wb.active
    ws.title = '取得結果'
    rows = dataframe_to_rows(output_data, index=False, header=True)

    for row_no, row in enumerate(rows, 3):
        for col_no, value in enumerate(row, 2):
            ws.cell(row=row_no, column=col_no, value=value)
    wb.save('out/集計結果.xlsx')

# エクセルスクレイピング
def xl_scraping(url, sheet_name, row, target_row, cell_name, FILEPATH):
    r = requests.get(url, allow_redirects=True)
    open(FILEPATH, 'wb').write(r.content)
    wb = xlrd.open_workbook(FILEPATH)
    st = wb.sheet_by_name(sheet_name)
    for col_j, cell in enumerate(st.row(row)):
        if(cell.value == cell_name):
            return st.cell_value(target_row, col_j)

# Webスクレイピング
def web_scraping(url, r, c,FILEPATH):
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

if __name__ == '__main__':
    main()
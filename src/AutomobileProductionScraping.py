#!/usr/bin/env python3
#coding: UTF-8
import os
import re
import requests
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import xlrd
import pandas as pd

# データ取得_トヨタ_ダイハツ_日野
def url_req_tyt_dht_hno():
    url = 'https://global.toyota/pages/global_toyota/company/profile/production-sales-figures/production_sales_figures_jp.xls'
    r = requests.get(url, allow_redirects=True)
    FILEPATH_TOYOTA_DAIHATSU_HINO ='data/toyota.xls'
    open(FILEPATH_TOYOTA_DAIHATSU_HINO, 'wb').write(r.content)
    wb_tyt_dht_hno = xlrd.open_workbook(FILEPATH_TOYOTA_DAIHATSU_HINO)
    st_tyt_dht_hno = wb_tyt_dht_hno.sheet_by_name('生産')
    for col_j, cell in enumerate(st_tyt_dht_hno.row(2)):
        if(cell.value ==202101.0):
            # トヨタ
            tyt_seisan_daisu = st_tyt_dht_hno.cell_value(6,col_j)
            # ダイハツ
            dht_seisan_daisu = st_tyt_dht_hno.cell_value(13,col_j)
            # 日野
            hno_seisan_daisu = st_tyt_dht_hno.cell_value(20, col_j)
    
    url = 'https://www.honda.co.jp/content/dam/site/www/investors/cq_img/financial_data/monthly/CY2020_202102_monthly_data_j.xlsx'
    r = requests.get(url, allow_redirects=True)
    FILEPATH_HONDA ='data/honda.xlsx'
    open(FILEPATH_HONDA, 'wb').write(r.content)
    wb_hnd = xlrd.open_workbook(FILEPATH_HONDA)
    st_hnd = wb_hnd.sheet_by_name('日本語')
    for col_j, cell in enumerate(st_hnd.row(84)):
        if(cell.value == '1月実績'):
            # ホンダ
            hnd_seisan_daisu = st_hnd.cell_value(85,col_j)

    output_data=pd.DataFrame({'メーカー':['トヨタ','ダイハツ','ホンダ','日産','スズキ','マツダ','三菱','スバル','いすゞ','日野','三菱ふそう'],'1月':[tyt_seisan_daisu,dht_seisan_daisu,hnd_seisan_daisu,hno_seisan_daisu,0,0,0,0,0,0,0]})
    output_excle(output_data)

# データ出力
def output_excle(output_data):
    # Excelワークブックの生成
    wb = Workbook()
    ws = wb.active
    ws.title = '取得結果'
    rows = dataframe_to_rows(output_data, index=False, header=True)
    # ワークシートへデータを書き込む
    row_start_idx = 3
    col_start_idx = 2
    for row_no, row in enumerate(rows, row_start_idx):
        for col_no, value in enumerate(row, col_start_idx):
            ws.cell(row=row_no, column=col_no, value=value)
    
    wb.save('out/集計結果.xlsx')

url_req_tyt_dht_hno()
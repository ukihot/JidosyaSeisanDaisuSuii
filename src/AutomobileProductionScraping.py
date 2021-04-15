#!/usr/bin/env python3
#coding: UTF-8
import os
import re
import requests
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import xlrd
import pandas as pd

# データ取得
def jidosya_seisan_daisu_suii():
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
            hno_seisan_daisu = st_tyt_dht_hno.cell_value(20,col_j)
    
    output_data=pd.DataFrame({'メーカー':['トヨタ','ダイハツ','ホンダ','日産','スズキ','マツダ','三菱','スバル','いすゞ','日野','三菱ふそう'],'1月':[tyt_seisan_daisu,dht_seisan_daisu,0,hno_seisan_daisu,0,0,0,0,0,0,0]})
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
    
    wb.save('集計結果.xlsx')

jidosya_seisan_daisu_suii()
#!/usr/bin/env python3
#coding: UTF-8
import requests
import openpyxl
import pandas as pd

# トヨタダイハツ日野
url = 'https://global.toyota/pages/global_toyota/company/profile/production-sales-figures/production_sales_figures_jp.xls'
r = requests.get(url, allow_redirects=True)

open('data/toyota.xls', 'wb').write(r.content)
book_toyota_daihatsu_hino = openpyxl.load_workbook('data/toyota.xls')
sheet_ = book_toyota_daihatsu_hino['生産']
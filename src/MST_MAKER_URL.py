# 資料内のマーカー（情報のありかが分かるキー）
tyt_key={1:202101.0,2:202102.0,3:202103.0,4:202104.0,5:202105.0,6:202106.0,7:202107.0,8:202108.0,9:202109.0,10:202110.0,11:202111.0,12:202112.0}
hnd_key={1:'1月実績',2:'2月実績',3:'3月実績',4:'4月実績',5:'5月実績',6:'6月実績',7:'7月実績',8:'8月実績',9:'9月実績',10:'10月実績',11:'11月実績',12:'12月実績'}

hso_key={1:11,2:13,3:15,4:22,5:24,6:26,7:33,8:35,9:37,10:44,11:46,12:48}

# 更新用の情報
meta_url ={'szk':'https://www.suzuki.co.jp/release/','mzd':'https://newsroom.mazda.com/ja/publicity/release/2021/','mtb':'https://www.mitsubishi-motors.com/jp/investors/finance_result/result.html?intcid2=investors-finance_result-result', 'hnd':'https://www.honda.co.jp/pressroom/corporate/'}

select = {'szk':['dl > a:nth-of-type(1) > dd','2021年1月 四輪車生産','2021年2月 四輪車生産','2021年3月 四輪車生産','2021年4月 四輪車生産','2021年5月 四輪車生産','2021年6月 四輪車生産','2021年7月 四輪車生産','2021年8月 四輪車生産','2021年9月 四輪車生産','2021年10月 四輪車生産','2021年11月 四輪車生産','2021年12月 四輪車生産'],'mzd':['ul > li:nth-of-type(1) > a','2021年1月の生産','2021年2月の生産','2021年3月の生産','2021年4月の生産','2021年5月の生産','2021年6月の生産','2021年7月の生産','2021年8月の生産','2021年9月の生産','2021年10月の生産','2021年11月の生産','2021年12月の生産'],'mtb':['tr > td:nth-of-type(1) > a','1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'],'hnd':['p.newsbox-txt > strong > a','2021年度 1月度 四輪車 生産・販売・輸出実績','2021年度 2月度 四輪車 生産・販売・輸出実績','2021年度 3月度 四輪車 生産・販売・輸出実績','2021年度 4月度 四輪車 生産・販売・輸出実績','2021年度 5月度 四輪車 生産・販売・輸出実績','2021年度 6月度 四輪車 生産・販売・輸出実績','2021年度 7月度 四輪車 生産・販売・輸出実績','2021年度 8月度 四輪車 生産・販売・輸出実績','2021年度 9月度 四輪車 生産・販売・輸出実績','2021年度 10月度 四輪車 生産・販売・輸出実績','2021年度 11月度 四輪車 生産・販売・輸出実績','2021年度 12月度 四輪車 生産・販売・輸出実績']}

# とってきた文字列と結合してURLを生成する
home_mzd = 'https://newsroom.mazda.com/'
hoge_mtb = 'https://www.mitsubishi-motors.com/'
home_szk = 'https://www.suzuki.co.jp/'

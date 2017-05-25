import talib
from talib import MA_Type
import sqlite3
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta
import sys
import time

#檢查list數值是否呈現遞減(mode='D')或遞增(mode='U')，回傳Y/N
def trend_chk(df, mode):
	i = 0
	flag = "Y"
	for elem in df:
		if i == 0:
			prev_elem = elem
			i += 1
			continue
		else:
			if (mode == "D") and (elem >= prev_elem):
				flag = "N"
				break
			elif (mode == "U") and (elem <= prev_elem):
				flag = "N"
				break
			else:
				prev_elem = elem
		i += 1
	return flag

def cal_reward(in_price, out_price):
	if in_price > 0:
		rt = (out_price - in_price) / in_price * 100	#報酬率
	else:
		rt = 0

	return rt

#K線型態判斷
def Patt_Recon(arg_stock, str_prev_date, str_today):
	sear_comp_id = arg_stock[0]
	#日線資料讀取
	strsql  = "select quo_date, open, high, low, close, vol from STOCK_QUO "
	strsql += "where "
	strsql += "SEAR_COMP_ID='" + sear_comp_id + "' and "
	strsql += "QUO_DATE between '" + str_prev_date + "' and '" + str_today + "' "
	strsql += "order by QUO_DATE "

	cursor = conn.execute(strsql)
	result = cursor.fetchall()

	none_flag = False
	if len(result) > 0:
		df = pd.DataFrame(result, columns = ['date','open','high','low','close','vol'])
		#print(df)
	else:
		none_flag = True

	#關閉cursor
	cursor.close()

	ls_result = []
	if none_flag == False:
		#LIST轉換為Numpy Array
		npy_close = np.array(df['close'])

		#計算50MA
		ma50 = talib.MA(npy_close, timeperiod=50, matype=0)
		df['ma50'] = ma50

		"""
		# for test 股價原始資料寫入EXCEL檔
		file_name = 'TOU_TEST.xlsx'
		writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
		df.to_excel(writer, sheet_name='stock', index=False)
		writer.save()
		"""
		for i in range(2,len(df)-63):
			dt = df.loc[i]['date']					#報價日期
			od1 = df.loc[i]['open']					#第一天開盤價
			od2 = df.loc[i+1]['open']				#第二天開盤價
			od3 = df.loc[i+2]['open']				#第三天開盤價

			cd1 = df.loc[i]['close']				#第一天收盤價
			cd2 = df.loc[i+1]['close']				#第二天收盤價
			cd3 = df.loc[i+2]['close']				#第三天收盤價

			vo2 = df.loc[i+1]['vol']				#第二天成交量
			vo3 = df.loc[i+2]['vol']				#第三天成交量

			#型態完成後，隔天進場，並計算一周、兩周、
			#一個月、兩個月後績效
			dt4 = df.loc[i+3]['date']				#進場第01天日期(底:進場日 / 頭:出場日)
			od4 = df.loc[i+3]['open']				#進場第01天開盤價(底:進場價 / 頭:出場價)
			dt7 = df.loc[i+10]['date']				#進場第07天日期(底:出場日 / 頭:進場日)
			cd7 = df.loc[i+10]['close']				#進場第07天收盤價(底:出場價 / 頭:進場價)
			dt14 = df.loc[i+17]['date']				#進場第14天日期(底:出場日 / 頭:進場日)
			cd14 = df.loc[i+17]['close']			#進場第14天收盤價(底:出場價 / 頭:進場價)
			dt30 = df.loc[i+33]['date']				#進場第30天日期(底:出場日 / 頭:進場日)
			cd30 = df.loc[i+33]['close']			#進場第30天收盤價(底:出場價 / 頭:進場價)
			dt60 = df.loc[i+63]['date']				#進場第60天日期(底:出場日 / 頭:進場日)
			cd60 = df.loc[i+63]['close']			#進場第60天收盤價(底:出場價 / 頭:進場價)

			#型態出現前，必須是已跌一段時間或漲一段
			#時間(以50MA取六天，判斷是否遞增/遞減)
			ma50_slice = df.loc[i-5:i]['ma50']		#往前取6天50MA值
			chk_d_yn = trend_chk(ma50_slice, 'D')	#判斷是否downtrend
			chk_u_yn = trend_chk(ma50_slice, 'U')	#判斷是否uptrend

			vol_3days_avg = df.loc[i:i+2]['vol'].mean()	#取3K棒3MV平均值

			#print(dt + " " + str(od1) + " " + str(od2) + " " + str(od3) + " " + str(od4))
			#print(ma50_slice2)
			#print(bull_chk_u_yn)
			#if i == 10:
			#	sys.exit("test end.")

			rec_data_yn = False
			patt_type = ""
			#Bullish patterns after downtrends
			#TWS判斷
			if (cd1 > od1 and cd2 > od2 and cd3 > od3) and \
			   (cd3 > cd2 > cd1) and (cd1 > od2 > od1) and \
			   (cd2 > od3 > od2) and (chk_d_yn == "Y"):
				rec_data_yn = True
				patt_type = "TWS"

			#TIU判斷
			if (od1 > cd1) and (od1 >= od2 > cd1) and (od1 > cd2 >= cd1) and \
			   (cd3 > od3) and (cd3 > od1) and (chk_d_yn == "Y"):
				rec_data_yn = True
				patt_type = "TIU"

			#TOU判斷
			if (od1 > cd1 and cd2 > od1 > cd1 > od2 and cd3 > od3 and cd3 > cd2) and \
			   (chk_d_yn == "Y"):
				rec_data_yn = True
				patt_type = "TOU"

			#MS判斷
			if (od1 > cd1) and (abs(od2 - cd2) > 0) and (cd1 > cd2) and \
			   (cd1 > od2) and (cd3 > od3) and (cd3 > (cd1 + (od1 - cd1)/2)) and \
			   (chk_d_yn == "Y"):
				rec_data_yn = True
				patt_type = "MS"

			#型態完成當天(第三天)成交量，必須比前一天(第二天)多20%以上成交量
			if rec_data_yn == True:
				if vo2 > 0:
					vol_up_rt = (vo3 - vo2) / vo2 * 100
				else:
					vol_up_rt = 0

				if vol_up_rt >= 20:
					rec_data_yn = True
				else:
					rec_data_yn = False

			#排除均量小於500張的股票
			if rec_data_yn == True:
				if vol_3days_avg < 500000:
					rec_data_yn = False

			#計算報酬率
			if rec_data_yn == True:
				#進場日期與進場價
				in_dt = dt4
				in_price = od4

				#計算進場後一周績效
				out_1w_dt = dt7
				out_1w_price = cd7
				rt_1w = cal_reward(in_price, out_1w_price)

				rt_1w_yn = "N"
				if rt_1w > 0.0:
					rt_1w_yn = "Y"

				#計算進場後兩周績效
				out_2w_dt = dt14
				out_2w_price = cd14
				rt_2w = cal_reward(in_price, out_2w_price)

				rt_2w_yn = "N"
				if rt_2w > 0.0:
					rt_2w_yn = "Y"

				#計算進場後一個月績效
				out_1m_dt = dt30
				out_1m_price = cd30
				rt_1m = cal_reward(in_price, out_1m_price)

				rt_1m_yn = "N"
				if rt_1m > 0.0:
					rt_1m_yn = "Y"

				#計算進場後二個月績效
				out_2m_dt = dt60
				out_2m_price = cd60
				rt_2m = cal_reward(in_price, out_2m_price)

				rt_2m_yn = "N"
				if rt_2m > 0.0:
					rt_2m_yn = "Y"

				ls_result.append([arg_stock[0],arg_stock[1], dt, dt[0:4], patt_type, in_dt, in_price, out_1w_dt, out_1w_price, rt_1w, rt_1w_yn, out_2w_dt, out_2w_price, rt_2w, rt_2w_yn, out_1m_dt, out_1m_price, rt_1m, rt_1m_yn, out_2m_dt, out_2m_price, rt_2m, rt_2m_yn])

	#print(ls_result)
	df_result = pd.DataFrame(ls_result, columns=['stock_id', 'stock_name', 'event_date', 'yyyy', 'pattern_type','in_date','in_price','out_1w_dt','out_1w_price','return_1week','1week_yn','out_2w_dt','out_2w_price','return_2week','2week_yn','out_1m_dt','out_1m_price','return_1mon','1mon_yn','out_2m_dt','out_2m_price','return_2mon','2mon_yn'])
	return df_result

############################################################################
# Main                                                                     #
############################################################################
#回測日期區間
str_prev_date = "20070101"
str_today = "20170524"

# 寫入LOG File
dt=datetime.datetime.now()
str_date = parser.parse(str(dt)).strftime("%Y%m%d")

name = "BotRev_backtesting_" + str_date + ".txt"
file = open(name, 'a', encoding = 'UTF-8')
tStart = time.time()#計時開始
file.write("\n\n\n*** LOG datetime  " + str(datetime.datetime.now()) + " ***\n")
file.write("回測日期區間:" + str_prev_date + "~" + str_today + "\n")

#建立資料庫連線
conn = sqlite3.connect("market_price.sqlite")

strsql  = "select SEAR_COMP_ID,COMP_NAME, STOCK_TYPE from STOCK_COMP_LIST "
#strsql += "where SEAR_COMP_ID='6217.TW' "
strsql += "order by STOCK_TYPE, SEAR_COMP_ID "
#strsql += "limit 100 "

cursor = conn.execute(strsql)
result = cursor.fetchall()

df_result = pd.DataFrame()
if len(result) > 0:
	for stock in result:
		print(stock)
		df = Patt_Recon(stock, str_prev_date, str_today)

		if len(df)>0:
			df_result = pd.concat([df_result, df], ignore_index=True)

#關閉cursor
cursor.close()

#關閉資料庫連線
conn.close()

#結果寫入EXCEL檔
file_name = 'BotRev_backtesting_day_50ma_3kar_' + str_date + '.xlsx'
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
df_result.to_excel(writer, sheet_name='stock', index=False)
writer.save()

tEnd = time.time()#計時結束
file.write ("\n\n\n結轉耗時 %f sec\n" % (tEnd - tStart)) #會自動做進位
file.write("*** End LOG ***\n")

# Close File
file.close()

print("End of prog.")
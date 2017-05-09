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

#K線型態判斷
def Patt_Recon(arg_stock, str_prev_date, str_today):
	sear_comp_id = arg_stock[0]
	#日線資料讀取
	strsql  = "select quo_date, open, high, low, close from STOCK_QUO "
	strsql += "where "
	strsql += "SEAR_COMP_ID='" + sear_comp_id + "' and "
	strsql += "QUO_DATE between '" + str_prev_date + "' and '" + str_today + "' "
	strsql += "order by QUO_DATE "

	cursor = conn.execute(strsql)
	result = cursor.fetchall()

	none_flag = False
	if len(result) > 0:
		df = pd.DataFrame(result, columns = ['date','open','high','low','close'])
		#print(df)
	else:
		none_flag = True

	#關閉cursor
	cursor.close()

	ls_result = []
	if none_flag == False:
		#LIST轉換為Numpy Array
		npy_close = np.array(df['close'])

		#計算3MA
		ma3 = talib.MA(npy_close, timeperiod=3, matype=0)
		df['ma3'] = ma3

		"""
		# for test 股價原始資料寫入EXCEL檔
		file_name = 'TOU_TEST.xlsx'
		writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
		df.to_excel(writer, sheet_name='stock', index=False)
		writer.save()
		"""
		for i in range(2,len(df)-4):
			dt = df.loc[i]['date']					#報價日期
			od1 = df.loc[i]['open']					#第一天開盤價
			od2 = df.loc[i+1]['open']				#第二天開盤價
			od3 = df.loc[i+2]['open']				#第三天開盤價
			cd1 = df.loc[i]['close']				#第一天收盤價
			cd2 = df.loc[i+1]['close']				#第二天收盤價
			cd3 = df.loc[i+2]['close']				#第三天收盤價
			ma3_slice = df.loc[i-5:i]['ma3']		#往前取6天3MA值
			chk_d_yn = trend_chk(ma3_slice, 'D')	#判斷是否downtrend
			chk_u_yn = trend_chk(ma3_slice, 'U')	#判斷是否uptrend
			ma3_slice2 = df.loc[i+2: i+4]['ma3']	#往前取2天3MA值
			bull_chk_u_yn = trend_chk(ma3_slice2, 'U')	#判斷是否底部空頭反轉
			bear_chk_u_yn = trend_chk(ma3_slice2, 'D')	#判斷是否頭部多頭反轉

			#print(cd1)
			#print(ma3_slice2)
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
				#print(dt)
				rec_data_yn = True
				patt_type = "TOU"
			
			#MS判斷
			if (od1 > cd1) and (abs(od2 - cd2) > 0) and (cd1 > cd2) and \
			   (cd1 > od2) and (cd3 > od3) and (cd3 > (cd1 + (od1 - cd1)/2)) and \
			   (chk_d_yn == "Y"):
				rec_data_yn = True
				patt_type = "MS"

			#Bearish patterns after uptrends
			#TBC判斷
			if (od1 > cd1 and od2 > cd2 and od3 > cd3) and \
			   (cd1 > cd2 > cd3) and (od1 > od2 > cd1) and \
			   (od2 > od3 > cd2) and (chk_u_yn == "Y"):
				rec_data_yn = True
				patt_type = "TBC"

			#TID判斷
			if (cd1 > od1) and (cd1 > od2 >= od1) and (cd1 >= cd2 > od1) and \
			   (od3 > cd3) and (od1 > cd3) and (chk_u_yn == "Y"):
				rec_data_yn = True
				patt_type = "TID"

			#TOD判斷
			if (cd1 > od1 and od2 > cd1 > od1 > cd2 and od3 > cd3 and cd2 > cd3) and \
			   (chk_u_yn == "Y"):
				rec_data_yn = True
				patt_type = "TOD"

			#ES判斷
			if (cd1 > od1) and (abs(od2 - cd2) > 0) and (cd2 > cd1) and \
			   (od2 > cd1) and (od3 > cd3) and (cd3 < (od1 + (cd1 - od1)/2)) and \
			   (chk_u_yn == "Y"):
				rec_data_yn = True
				patt_type = "ES"

			if rec_data_yn == True:
				ls_result.append([arg_stock[0],arg_stock[1], dt, patt_type, bull_chk_u_yn, bear_chk_u_yn])

	#print(ls_result)
	df_result = pd.DataFrame(ls_result, columns=['stock_id', 'stock_name', 'event_date', 'pattern_type','bottom_reverse','top_reverse'])
	return df_result

############################################################################
# Main                                                                     #
############################################################################
#回測日期區間
str_prev_date = "20070101"
str_today = "20170505"

# 寫入LOG File
dt=datetime.datetime.now()
str_date = parser.parse(str(dt)).strftime("%Y%m%d")

name = "PATT_RECON_" + str_date + ".txt"
file = open(name, 'a', encoding = 'UTF-8')
tStart = time.time()#計時開始
file.write("\n\n\n*** LOG datetime  " + str(datetime.datetime.now()) + " ***\n")
file.write("回測日期區間:" + str_prev_date + "~" + str_today + "\n")

#建立資料庫連線
conn = sqlite3.connect("market_price.sqlite")

strsql  = "select SEAR_COMP_ID,COMP_NAME, STOCK_TYPE from STOCK_COMP_LIST "
#strsql += "where STOCK_TYPE = '上櫃' and SEAR_COMP_ID='3662.TW' "
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

#結果寫入CSV FILE
#print(df_result)
df_result.to_csv('PATT_RECON_RESULT.csv', encoding='utf-8')

tEnd = time.time()#計時結束
file.write ("\n\n\n結轉耗時 %f sec\n" % (tEnd - tStart)) #會自動做進位
file.write("*** End LOG ***\n")

# Close File
file.close()

print("End of prog.")
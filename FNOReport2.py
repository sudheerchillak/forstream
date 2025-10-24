# -*- coding: utf-8 -*-
"""
Created on Fri Oct 24 12:29:04 2025

@author: Laptop
"""

import streamlit as st
import sys
# import matplotlib.pyplot as plt
from io import StringIO
import yfinance as yf
import pandas as pd
import xlsxwriter
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Font
import numpy as np
import csv 
import datetime
from datetime import date,datetime,timedelta
import time
from time import strftime, localtime
import warnings
warnings.filterwarnings('ignore')
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

stk_dict = {'COALINDIA.NS':[342.3,425.9], 'BANKBARODA.NS':[187.85,276.3], 'PNB.NS':[78.05,128.8], 'PFC.NS':[371.9,474.8], 'ONGC.NS':[203.65,299.7], 'NTPC.NS':[261.75,375.15], 'VEDL.NS':[273.35,526.95],
             'BHEL.NS':[193.5,285.5], 'JUBLFOOD.NS':[548.75,796.75], 'LAURUSLABS.NS':[721.05,1056.15], 'MFSL.NS':[1248.6,1686.75], 'NAUKRI.NS':[972.45,1437.8], 'TATACHEM.NS':[880.85,1048],
             'TRENT.NS':[3750.25,6258], 'UBL.NS':[1736,2134.7], 'CANBK.NS':[90.95,130.14], 'DELTACORP.NS':[65.75,103.89], 'HINDALCO.NS':[501.2,944.6], 'LICHSGFIN.NS':[522.65,668.75],
             'NATIONALUM.NS':[176.34,263], 'NMDC.NS':[66.8,85.13], 'RECLTD.NS':[259.45,460.2], 'SBIN.NS':[730,912], 'ATUL.NS':[4176.05,8335.85], 'METROPOLIS.NS':[1691.1,2365],
             'UPL.NS':[618,829.4], 'CANFINHOME.NS':[691.55,891.9], 'INDUSTOWER.NS':[292,430], 'AUBANK.NS':[655.5,971], 'BERGEPAINT.NS':[486.55,624], 'DALBHARAT.NS':[2011.6,2548.4],
             'PEL.NS':[905.3,1402], 'RBLBANK.NS':[205.5,346.95], 'TATACOMM.NS':[1030.8,2175], 'KOTAKBANK.NS':[1871.6,2301.9], 'ACC.NS':[1704.45,2314.9], 'AARTIIND.NS':[332.85,529.5],
             'ASIANPAINT.NS':[2117.15,2965.75], 'DABUR.NS':[326.2,629.75], 'IGL.NS':[193.12,284.8], 'INDIAMART.NS':[2155.1,3453.5], 'INDUSINDBK.NS':[590,1098.6],'RAIN.NS':[110.55,241.7],
             'WHIRLPOOL.NS':[936,1981.1], 'SUNTV.NS':[436.45,701.95], 'TATAMOTORS.NS':[365.8,820.35], 'ABFRL.NS':[68.74,98.76], 'BSOFT.NS':[327.1,541.75], 'CROMPTON.NS':[268.95,373],
             'DEEPAKNTR.NS':[1533.25,2399.95], 'INFY.NS':[1105.05,1732.95], 'IRCTC.NS':[604.1,863.3], 'LTTS.NS':[3670,4879.95], 'SYNGENE.NS':[576.55,761.4], 'TCS.NS':[2600.05,3710],
             'AXISBANK.NS':[1032.25,1339.65], 'BANDHANBNK.NS':[153.43,215.44], 'BPCL.NS':[262,376], 'CUB.NS':[187.05,249.35], 'FEDERALBNK.NS':[183.15,262.3], 'GAIL.NS':[159.61,216.47],
                 'HINDPETRO.NS':[341.5,546.45], 'IOC.NS':[122.35,181.34], 'MGL.NS':[1043.1,1958], 'PETRONET.NS':[222.5,368.65], 'ABB.NS':[3390.4,6947.9], 'ABCAPITAL.NS':[243,363.85],
             'ADANIENT.NS':[2145,3070], 'ADANIPORTS.NS':[1242.7,1499.5], 'ALKEM.NS':[4611.85,5751], 'AMBUJACEM.NS':[480.35,643.3], 'APOLLOTYRE.NS':[375,530.7], 'ASHOKLEY.NS':[108.1,149.17],
             'ASTRAL.NS':[959.05,1657.95], 'AUROPHARMA.NS':[805.65,1417.3], 'BAJAJFINSV.NS':[1617,2316], 'BALKRISIND.NS':[2067.15,3155.8], 'BALRAMCHIN.NS':[422,633.55],
             'BATAINDIA.NS':[813.1,1654.45], 'BHARATFORG.NS':[887.2,1394.55], 'BIOCON.NS':[327.6,424.75], 'BOSCHLTD.NS':[32120,46898.5], 'BRITANNIA.NS':[5276.5,6469.9],
             'CIPLA.NS':[1451.2,1702.05],'COFORGE.NS':[1464.3,2005.35], 'COLPAL.NS':[1969.15,3115], 'CONCOR.NS':[512.9,629.2], 'CUMMINSIND.NS':[3312.4,4171.9], 'DIXON.NS':[10620,19148.9],
             'DLF.NS':[565,967.6], 'DRREDDY.NS':[1015.25,1421.5], 'ESCORTS.NS':[2470.25,4420], 'EXIDEIND.NS':[370.1,431.6], 'GLENMARK.NS':[1740.1,2284.8], 'GMRAIRPORT.NS':[75.83,98.23],
             'GODREJCP.NS':[902.05,1314], 'GODREJPROP.NS':[1854.55,2522], 'GRANULES.NS':[365.45,628], 'GRASIM.NS':[2465.5,3000], 'GUJGASLTD.NS':[371.95,553], 'HAL.NS':[3820.10,5674.75],
             'HAVELLS.NS':[1292.90,1721.2], 'HEROMOTOCO.NS':[4158.1,6246.25], 'HINDCOPPER.NS':[173.2,377], 'HINDUNILVR.NS':[2286.6, 3035], 'HONAUT.NS':[33910,46599], 'ICICIGI.NS':[1702,2301.9],
             'IDEA.NS':[5.8,10.53], 'IDFCFIRSTB.NS':[63.95,89.65], 'IEX.NS':[124.6,244.4], 'INDHOTEL.NS':[645.9,894.9], 'INDIACEM.NS':[329.75,420], 'ITC.NS':[357.3,498.85],
             'JINDALSTEL.NS':[882.65,1097], 'LALPATHLAB.NS':[2708,3653.95], 'LT.NS':[2856.15,3963.5], 'M&MFIN.NS':[226.7,330.7], 'MARUTI.NS':[10800.2,18519.45], 'MOTHERSON.NS':[87.08,116.38],
             'MPHASIS.NS':[1806.6,3237.95], 'MRF.NS':[131500,176827.45], 'NAVINFLUOR.NS':[4189.4,5444], 'NESTLEIND.NS':[968.25,1389], 'OBEROIRLTY.NS':[1359.25,2343.65], 'OFSS.NS':[6381,13220],
             'PAGEIND.NS':[38300,50799.9], 'PERSISTENT.NS':[3135.4,6303.95], 'PIDILITIND.NS':[1393.95,1707.5], 'PIIND.NS':[3051,4804.05], 'POLYCAB.NS':[6560,9369.1], 'POWERGRID.NS':[226.05,345.4],
             'RAMCOCEM.NS':[964,1209], 'RELIANCE.NS':[1095.35,1608.8], 'SAIL.NS':[92.2,144.2], 'SBILIFE.NS':[1693.1,1936], 'SIEMENS.NS':[2788.3,3500.3], 'SUNPHARMA.NS':[1498.3,1960.35],
             'TATACONSUM.NS':[989.25,1253.7], 'TATAPOWER.NS':[319.6,454.75], 'TATASTEEL.NS':[139.45,184.6], 'TECHM.NS':[1060.1,1807.7], 'TITAN.NS':[3245.5,3886.95], 'TORNTPHARM.NS':[3287.3,3787.9],
             'TORNTPOWER.NS':[1087.55,1492], 'TVSMOTOR.NS':[2640,4316.95],'VOLTAS.NS':[1101.2,1944.9], 'WIPRO.NS':[201.05,292], 'ZEEL.NS':[95.15,209.7], 'ZYDUSLIFE.NS':[917.05,1087.25],
             'BAJAJ-AUTO.NS':[7612,10079.8], 'LTF.NS':[170.46,275.9], 'UNITDSPR.NS':[1249.3,1433.1], 'SHRIRAMFIN.NS':[403.2,730.45], 'ARE&M.NS':[768.3,1220.05], 'PVRINOX.NS':[915,1329],
             'DIVISLAB.NS':[4615.55,7071.5], 'INTELLECT.NS':[774.45,1255], 'JKCEMENT.NS':[6060,7565.5], 'BEL.NS':[304.8,436], 'INDIGO.NS':[4634.6,6232.5], 'MANAPPURAM.NS':[231.1,340.7],
             'MCX.NS':[6561,10453.85], 'MUTHOOTFIN.NS':[2189.4,3414.7], 'SRF.NS':[2200.45,3325], 'BAJFINANCE.NS':[821.15,1178.95], 'HDFCAMC.NS':[4984,6671], 'ABBOTINDIA.NS':[22852,37000],
             'HDFCBANK.NS':[869.1,1018.85], 'BHARTIARTL.NS':[1799.5,2494.9], 'CHAMBLFERT.NS':[316.5,742.2], 'CHOLAFIN.NS':[1225.35,2075.5], 'COROMANDEL.NS':[1855.05,2718.9],
             'EICHERMOT.NS':[4999.95,7456.05], 'HDFCLIFE.NS':[708.4,820.75], 'ICICIBANK.NS':[1265,1500], 'MARICO.NS':[641,759], 'SHREECEM.NS':[26419.75,37490], 'ULTRACEMCO.NS':[10240,14696.25],
             'APOLLOHOSP.NS':[6677.5,8379.55], 'FSL.NS':[208.45,422.3], 'HCLTECH.NS':[1208.55,2012.2], 'IPCALAB.NS':[1041,1755.9], 'LUPIN.NS':[1491.05,2225], 'JSWSTEEL.NS':[1016,1295.65],
             'M&M.NS':[3053.6,4027.35]}
# ,'^NSEI':[23935.75,26277.35],'^NSEBANK':[53483.05,57628.4]
data2 =[]
data3 =[]
cols_dict = {'Date':'last','Open':'first', 'High':'max', 'Low':'min', 'Close':'last','Volume':'sum','Year':'last',
             'Week_Number':'last', 'Mth_Number' :'last','Stock_Name':'last' }
for k, v in stk_dict.items():
    # Argument to be passed - Stock name
    ticker = k
    # print(ticker)
    # Download historical data from yahoo
    try:
        start_date = (datetime.today()- timedelta(days=150)).strftime('%Y-%m-%d')
        end_date = (datetime.today()+ timedelta(days=1)).strftime('%Y-%m-%d')
        # print(start_date, end_date)
        data1 = yf.Ticker(ticker)
        data = data1.history(start=start_date, end=end_date)
        # Defining the Date Column and eliminating the time component from date field
        data["Date"] = data.index
        data['Date'] = pd.to_datetime(data['Date'])
        data['Date'] = data['Date'].dt.date
        # Assigning the column names and filtering the dataset for the needed columns
        data = data[["Date", "Open", "High", "Low", "Close",  "Volume"]]
        data['Year'] = pd.to_datetime(data['Date']).dt.year
        data['Week_Number'] = pd.to_datetime(data['Date']).dt.isocalendar().week
        data['Mth_Number'] = pd.to_datetime(data['Date']).dt.month
        
        if 'trailingEps' in data1.info:
                data["trailingEps"] =  data1.info['trailingEps']
        else:
                data["trailingEps"] = False
        if ["trailingEps"] == False:
                data["Trailing_PE"] = "NA"
        else:
                data["Trailing_PE"] = round(data['Close']/data["trailingEps"], 2)
        if 'industry' in data1.info:
            data["Industry"] =  data1.info['industry']
        else:
            data["Industry"] = "NA"
        if 'sector' in data1.info:
            data["sector"] =  data1.info['sector']
        else:
            data["sector"] = "NA"
        if 'marketCap' in data1.info:
            data["marketCap"] =  data1.info['marketCap']
        else:
            data["marketCap"] = 0
        if 'bookValue' in data1.info:
            data["bookValue"] =  data1.info['bookValue']
        else:
            data["bookValue"] = False 
        if 'returnOnAssets' in data1.info:
            data["returnOnAssets"] =  data1.info['returnOnAssets']
            data["returnOnAssets"] =  pd.to_numeric(data["returnOnAssets"], errors='coerce')
            data["returnOnAssets"] = data["returnOnAssets"].map('{:.2%}'.format)
        else:
            data["returnOnAssets"] = 0
        if 'returnOnEquity' in data1.info:
            data["returnOnEquity"] =  data1.info['returnOnEquity']
            data["returnOnEquity"] =  pd.to_numeric(data["returnOnEquity"], errors='coerce')
            data["returnOnEquity"] = data["returnOnEquity"].map('{:.2%}'.format)
        else:
            data["returnOnEquity"] = 0
        if ["bookValue"] == False:
            data["priceToBook"] =  "NA"
        else:
            data["priceToBook"] = round(data['Close']/data["bookValue"], 2)
        if 'earningsQuarterlyGrowth' in data1.info:
            data["earningsQuarterlyGrowth"] =  data1.info['earningsQuarterlyGrowth']
            data["earningsQuarterlyGrowth"] =  pd.to_numeric(data["earningsQuarterlyGrowth"], errors='coerce')
            data["earningsQuarterlyGrowth"] = data["earningsQuarterlyGrowth"].map('{:.2%}'.format)
            
        else:
            data["earningsQuarterlyGrowth"] = 0
        if 'dividendRate' in data1.info:
                data["dividendRate"] =  data1.info['dividendRate']/100
                data["dividendRate"] =  pd.to_numeric(data["dividendRate"], errors='coerce')
                data["dividendRate"] = data["dividendRate"].map('{:.2%}'.format)
        else:
                data["dividendRate"] = False
        if 'dividendYield' in data1.info:
                data["dividendYield"] =  data1.info['dividendYield']/100
                data["dividendYield"] =  pd.to_numeric(data["dividendYield"], errors='coerce')
                data["dividendYield"] = data["dividendYield"].map('{:.2%}'.format)
        else:
                data["dividendYield"] = False
        if 'lastDividendValue' in data1.info:
            data["lastDividendValue"] =  data1.info['lastDividendValue']
        else:
            data["lastDividendValue"] = False
        if 'lastDividendDate' in data1.info:
            data["lastDividendDate"] =  data1.info['lastDividendDate']
            data["lastDividendDate"] =  pd.to_datetime(data['lastDividendDate'], unit='s')
            data['lastDividendDate'] = data['lastDividendDate'].dt.date
        else:
            data["lastDividendDate"] = None
        if 'lastDividendValue' in data1.info:
            data["lastDividendValue"] =  data1.info['lastDividendValue']
        else:
                data["lastDividendValue"] = False
        if 'totalCash' in data1.info:
            data["totalCash"] =  data1.info['totalCash']
        else:
            data["totalCash"] = False
        if 'totalCashPerShare' in data1.info:
            data["totalCashPerShare"] =  data1.info['totalCashPerShare']
        else:
            data["totalCashPerShare"] = False
        if 'debtToEquity' in data1.info:
            data["debtToEquity"] =  data1.info['debtToEquity']
        else:
            data["debtToEquity"] = "NA"			
        if 'totalDebt' in data1.info:
            data["totalDebt"] =  data1.info['totalDebt']
        else:
              data["totalDebt"] = False			
        if 'sharesOutstanding' in data1.info:
            data["sharesOutstanding"] =  data1.info['sharesOutstanding']
        else:
            data["sharesOutstanding"] = False
        if 'totalCash' in data1.info:
            data["totalCash"] =  data1.info['totalCash']
        else:
              data["totalCash"] = False			
        if 'totalRevenue' in data1.info:
                    data["totalRevenue"] =  data1.info['totalRevenue']
        else:
                      data["totalRevenue"] = False
        if 'freeCashflow' in data1.info:
                    data["freeCashflow"] =  data1.info['freeCashflow']
        else:
                      data["freeCashflow"] = False
        if 'operatingCashflow' in data1.info:
                    data["operatingCashflow"] =  data1.info['operatingCashflow']
        else:
                      data["operatingCashflow"] = False			  
        if 'ebitda' in data1.info:
                    data["ebitda"] =  data1.info['ebitda']
        else:
                      data["ebitda"] = False        
        try:
            data["RevperShr"] = round(data["totalRevenue"]/data["sharesOutstanding"],2)
        except:
            data["RevperShr"] = 0
        try:
            data["DebtperShr"] = round(data["totalDebt"]/data["sharesOutstanding"],2)
        except:
            data["DebtperShr"] = 0    
        try:
            data["freeCashflowperShr"] = round(data["freeCashflow"]/data["sharesOutstanding"],2)
        except:
            data["freeCashflowperShr"] = 0 
        try:
            data["operatingCashflowperShr"] = round(data["operatingCashflow"]/data["sharesOutstanding"],2)
        except:
            data["operatingCashflowperShr"] = 0 
        try:
            data["ebitdaperShr"] = round(data["ebitda"]/data["sharesOutstanding"],2)
        except:
            data["ebitdaperShr"] = 0
        data['Stock_Name'] = k
        data.reset_index(drop=True, inplace=True)
        
        data1 = data[["Date"]].tail(1)
        datefor = list(data[["Date"]].iloc[-1].astype(str))
        # print(data1)
        # data1['Curr_Day'] = data['Date'].iloc[-1]
        data1['Stock_Name'] = data['Stock_Name'].iloc[-1]
        data1['Trail_PE'] = data["Trailing_PE"].iloc[-1]
        data1['ROA'] = data["returnOnAssets"].iloc[-1]
        data1['ROE'] = data["returnOnEquity"].iloc[-1]
        data1['Curr_Day_Close'] = round(data['Close'].iloc[-1],2)
        data1['Curr_Day_Close%'] =  ('{:,.2%}'.format(round((data['Close'].iloc[-1] - data['Close'].iloc[-2])/data['Close'].iloc[-2],4)))
        data1['Curr_Day_Volume%'] =  ('{:,.2%}'.format(round((data['Volume'].iloc[-1] - data['Volume'].iloc[-2])/data['Volume'].iloc[-2],4)))
        data1['Curr_Day_High'] = round(data['High'].iloc[-1],2)
        data1['Curr_Day_High%'] =  ('{:,.2%}'.format(round((data['High'].iloc[-1] - data['High'].iloc[-2])/data['High'].iloc[-2],4)))
        data1['Curr_Day_Low'] = round(data['Low'].iloc[-1],2)
        data1['Curr_Day_Low%'] =  ('{:,.2%}'.format(round((data['Low'].iloc[-1] - data['Low'].iloc[-2])/data['Low'].iloc[-2],4)))        
        data1['Buy_Price1'] = v[0]
        data1['Buy_Distance%'] =  (round((data1['Curr_Day_Close'] - data1['Buy_Price1']) / data1['Buy_Price1'], 2)).apply(lambda x: "{:.2f}%".format(x * 100))
        data1['Buy_Price1_LoTch'] = data1['Curr_Day_Low'] <= data1['Buy_Price1']
        data1['Sell_Price1'] = v[1]
        data1['Sell_Distance%'] =  (round((data1['Curr_Day_Close'] - data1['Sell_Price1']) / data1['Sell_Price1'], 2)).apply(lambda x: "{:.2f}%".format(x * 100))    
        data1['Sell_Price1_HiTch'] = data1['Curr_Day_High'] >= data1['Sell_Price1']
        data1['1Day_Gap2%'] =  (data['Low'].iloc[-1] - data['High'].iloc[-2]) >= (data['Close'].iloc[-1]*0.02)
        data1['5Day_Perf'] = ('{:,.2%}'.format((data['Close'].iloc[-1] - data['Close'].iloc[-5])/data['Close'].iloc[-5]))
        data1['30Day_Perf'] = ('{:,.2%}'.format((data['Close'].iloc[-1] - data['Close'].iloc[-30])/data['Close'].iloc[-30]))
        data1['90Day_Perf'] = ('{:,.2%}'.format((data['Close'].iloc[-1] - data['Close'].iloc[-90])/data['Close'].iloc[-90]))
        data1['Year'] = data["Year"].iloc[-1]
        data1['Week_Number'] = data["Week_Number"].iloc[-1]
        data1['Mth_Number'] = data["Mth_Number"].iloc[-1]
        # data_monthly = data.groupby(['Year','Mth_Number']).agg(cols_dict)
                
        data1['TrailEps'] = data["trailingEps"].iloc[-1]
        data1['Industry'] = data["Industry"].iloc[-1]
        data1['Sector'] = data["sector"].iloc[-1]
        data1['MCAP'] = data["marketCap"].iloc[-1]
        data1['BV'] = data["bookValue"].iloc[-1]        
        data1['PB'] = data["priceToBook"].iloc[-1]
        data1['EarQtrGrw'] = data["earningsQuarterlyGrowth"].iloc[-1]
        data1['dividendRate'] = data["dividendRate"].iloc[-1]
        data1['dividendYield'] = data["dividendYield"].iloc[-1]
        data1['lastDividendValue'] = data["lastDividendValue"].iloc[-1]
        data1['lastDividendDate'] = data["lastDividendDate"].iloc[-1]
        data1['lastDividendValue'] = data["lastDividendValue"].iloc[-1]
        data1['totalCash'] = data["totalCash"].iloc[-1]
        data1['totalCashPerShare'] = data["totalCashPerShare"].iloc[-1]
        data1['debtToEquity'] = data["debtToEquity"].iloc[-1]
        data1['totalDebt'] = data["totalDebt"].iloc[-1]
        data1['sharesOutstanding'] = data["sharesOutstanding"].iloc[-1]
        data1['totalCash'] = data["totalCash"].iloc[-1]
        data1['totalRevenue'] = data["totalRevenue"].iloc[-1]
        data1['netCash'] = data1["totalCash"].iloc[-1] - data1['totalDebt']
        data1['freeCashflow'] = data["freeCashflow"].iloc[-1]
        data1['operatingCashflow'] = data["operatingCashflow"].iloc[-1]
        data1['ebitda'] = data["ebitda"].iloc[-1]        
        data1['netCashShr'] = round(data1["netCash"].iloc[-1]/data1["sharesOutstanding"].iloc[-1],2)
        data1['netCashShrperCls'] = round(data1["netCashShr"].iloc[-1]/data['Close'].iloc[-1],2)        
        data1['RevperShr'] = data["RevperShr"].iloc[-1]
        data1['RevperShrperCls'] = round(data["RevperShr"].iloc[-1]/data['Close'].iloc[-1],2)        
        data1['DebtperShr'] = data["DebtperShr"].iloc[-1]
        data1['DebtperShrperCls'] = round(data["DebtperShr"].iloc[-1]/data['Close'].iloc[-1],2)
        data1['freeCashflowperShr'] = data["freeCashflowperShr"].iloc[-1]
        data1['freeCashflowperShrperCls'] = round(data["freeCashflowperShr"].iloc[-1]/data['Close'].iloc[-1],2)
        data1['operatingCashflowperShr'] = data["operatingCashflowperShr"].iloc[-1]
        data1['operatingCashflowperShrperCls'] = round(data["operatingCashflowperShr"].iloc[-1]/data['Close'].iloc[-1],2)
        data1['ebitdaperShr'] = data["ebitdaperShr"].iloc[-1]
        data1['ebitdaperShrperCls'] = round(data["ebitdaperShr"].iloc[-1]/data['Close'].iloc[-1],2)
        data1['Buy_Distance'] =  (round((data1['Curr_Day_Close'] - data1['Buy_Price1']) / data1['Buy_Price1'], 2))
        data1['Sell_Distance'] =  (round((data1['Curr_Day_Close'] - data1['Sell_Price1']) / data1['Sell_Price1'], 2))
        data1['Prev_Day_Close'] = round(data['Close'].iloc[-2],2)
        data1['Buy_Distance2'] =  (round((data1['Prev_Day_Close'] - data1['Buy_Price1']) / data1['Buy_Price1'], 2))
        data1['Sell_Distance2'] =  (round((data1['Prev_Day_Close'] - data1['Sell_Price1']) / data1['Sell_Price1'], 2))
        data2.append(data1)
        data3.append(data)
    except:
        continue
final_df = pd.concat(data2, ignore_index=True)
final_df_new = pd.concat(data3, ignore_index=True)

final_df.style
try:
    final_df["MCAP"] =  round(final_df["MCAP"]/1000000000, 2)
    final_df["MCAP"] = '$' + ((final_df["MCAP"].astype(float))).astype(str) + ' B'
except:
    final_df["MCAP"] = 0
final_df['netCashShrperCls'] = final_df['netCashShrperCls'].map('{:,.2%}'.format)
final_df['RevperShrperCls'] = final_df['RevperShrperCls'].map('{:,.2%}'.format)
final_df['DebtperShrperCls'] = final_df['DebtperShrperCls'].map('{:,.2%}'.format)
final_df['freeCashflowperShrperCls'] = final_df['freeCashflowperShrperCls'].map('{:,.2%}'.format)
final_df['operatingCashflowperShrperCls'] = final_df['operatingCashflowperShrperCls'].map('{:,.2%}'.format)
final_df['ebitdaperShrperCls'] = final_df['ebitdaperShrperCls'].map('{:,.2%}'.format)
final_df1 = final_df.sort_values(by='Buy_Distance')
final_df2 = final_df.sort_values(by='Sell_Distance', ascending=False)
final_df1['Rank_Curr_Buy_Price1'] =  final_df['Buy_Distance'].rank(method='dense')
final_df1['Rank_Prev_Buy_Price1'] =  final_df['Buy_Distance2'].rank(method='dense')
final_df2['Rank_Curr_SellPrice1'] =  final_df['Sell_Distance'].rank(method='dense')
final_df2['Rank_Prev_Sell_Price1'] =  final_df['Sell_Distance2'].rank(method='dense')
final_df1 = final_df1.drop('Buy_Distance', axis=1)
final_df2 = final_df2.drop('Sell_Distance', axis=1)

final_df1.reset_index(drop=True, inplace=True)
final_df2.reset_index(drop=True, inplace=True)

cols = final_df1.columns.tolist()
cols = cols[-1:] + cols[:-1]
cols = cols[-1:] + cols[:-1]
final_df1 = final_df1[cols]
cols = final_df2.columns.tolist()
cols = cols[-1:] + cols[:-1]
cols = cols[-1:] + cols[:-1]
final_df2 = final_df2[cols]
# final_df.style.set_table_styles([{'selector': 'th,td', 'props': [('border-style', 'solid'), ('border-width', '1px'), ('border-color', 'grey')]}]).to_html()
# print(final_df)
final_df1.to_csv('FNO_Buy_Report '+datefor[0]+'.csv')
df1 = final_df1[final_df1['Buy_Price1_LoTch']==True]
print(df1[['Stock_Name','Curr_Day_Close']])

final_df2.to_csv('FNO_Sell_Report '+datefor[0]+'.csv')
df2 = final_df2[final_df2['Sell_Price1_HiTch']==True]
print(df2[['Stock_Name','Curr_Day_Close']])



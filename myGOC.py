import pandas as pd
import pprint as pp
from yahooquery import Ticker
from datetime import datetime, timedelta
from pandas.tseries.offsets import Day
import time

myDay = [45, 5, 20, 30, 73]                                 ### 0: Index series period
                                                            ### 1: 1st no. of days performance
                                                            ### 2: 2nd no. of days performance
                                                            ### 3: extract top5 from 30 days
															### 4: GOC series period
															
myTic = [5]                                                 ### 0: Top 5 daily return tickers
ConcDay = [26]												### Shorter version of GOC: COnv = 26 days & Base = 9 days
BaseDay = [9]

### NASDAQ 100 ###
NDX=['AAPL',	'ADBE',	'ADI',	'ADP',	'ADSK',	'ALGN',	'ALXN',	'AMAT',	'AMD',	'AMGN',	'AMZN',	'ANSS',	'ASML',	'ATVI',	'AVGO',	'BIDU',	'BIIB',	'BKNG',	'BMRN',	'CDNS',	'CDW',	'CERN',	'CHKP',	'CHTR',	'CMCSA',	'COST',	'CPRT',	'CSCO',	'CSX',	'CTAS',	'CTSH',	'CTXS',	'DLTR',	'DOCU',	'DXCM',	'EA',	'EBAY',	'EXC',	'EXPE',	'FAST',	'FB',	'FISV',	'FOX',	'FOXA',	'GILD',	'GOOG',	'GOOGL',	'IDXX',	'ILMN',	'INCY',	'INTC',	'INTU',	'ISRG',	'JD',	'KHC',	'KLAC',	'LBTYA',	'LBTYK',	'LRCX',	'LULU',	'MAR',	'MCHP',	'MDLZ',	'MELI',	'MNST',	'MRNA',	'MSFT',	'MU',	'MXIM',	'NFLX',	'NTAP',	'NTES',	'NVDA',	'NXPI',	'ORLY',	'PAYX',	'PCAR',	'PEP',	'PYPL',	'QCOM',	'REGN',	'ROST',	'SBUX',	'SGEN',	'SIRI',	'SNPS',	'SPLK',	'SWKS',	'TCOM',	'TMUS',	'TSLA',	'TTWO',	'TXN',	'ULTA',	'VRSK',	'VRSN',	'VRTX',	'WBA',	'WDAY',	'WDC',	'XEL',	'XLNX',	'ZM']

starttime = time.perf_counter()
myDT = datetime.now().strftime("%Y%m%d-%H%M")
path = r"C:\myGOC-" + myDT + ".xlsx"		# Home path
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')

def split_list(a_list):
    half = len(a_list)//2
    return a_list[:half], a_list[half:]

def split_list_square(aa_list):
	bb, cc = split_list(aa_list)
	mm, nn = split_list(bb)
	xx, yy = split_list(cc)

	return mm, nn, xx, yy

def SimpleLastTD(adate):
	adate -= timedelta(days=1)
	while adate.weekday() > 4: 							          # Mon-Fri are 0-4
		adate -= timedelta(days=1)
	return adate

def myPriceChg(mytickerlist):
	AllPricesChg = []
	AllPricesChg = pd.DataFrame(AllPricesChg)
	try:                                                          # Skip invalid ticker
		Oneticker = Ticker(mytickerlist)
		OnePrice = Oneticker.price						          # Price Dict from whole series of Index
		OnePrice = pd.DataFrame(OnePrice)		                  # Dict convert to dataframe
		OnePriceChg = OnePrice.loc['regularMarketChangePercent']  # Filter 1-day % chg
		OnePriceChg = pd.DataFrame(OnePriceChg)
		AllPricesChg = pd.concat([AllPricesChg, OnePriceChg])
		#AllPricesChg = pd.concat([AllPricesChg, OnePriceChg], sort=True)
	except Exception:
		pass
	return AllPricesChg


def myPriceHist(mytickerlist, myDiff):
	AllDays = []
	AllDays = pd.DataFrame(AllDays)
	DateNow = datetime.now()
	PreviousDate = SimpleLastTD(datetime.now() - Day(myDiff))     # 45-day raw data

	try:
		Oneticker = Ticker(mytickerlist)
		Oneticker = Oneticker.history(start=PreviousDate, end=DateNow, interval='1d')
		Oneticker.fillna(method='ffill', inplace=True)       	  # If NaN, use previous value
		Oneticker = pd.DataFrame(Oneticker)
		AllDays = pd.concat([AllDays, Oneticker])
		#AllDays = pd.concat([AllDays, Oneticker], sort=True)
	except Exception:
		pass
	return AllDays

def myFundamental1(mytickerlist1, funda1):
	AllFundamental1 = []
	AllFundamental1 = pd.DataFrame(AllFundamental1)
	try:
		ThisFunda1 = Ticker(mytickerlist1)
		ThisFunda1 = ThisFunda1.summary_detail              # Fundamental1: 'volume','beta','trailingPE','dividendRate'
		ThisFunda1 = pd.DataFrame(ThisFunda1)
		AllFundamental1 = ThisFunda1.loc[funda1]
	except Exception:
		pass
	return AllFundamental1

def myFundamental2(mytickerlist2, funda2):
	AllFundamental2 = []
	AllFundamental2 = pd.DataFrame(AllFundamental2)
	try:
		ThisFunda2 = Ticker(mytickerlist2)
		ThisFunda2 = ThisFunda2.key_stats              	  # Fundamental2: 'bookValue','priceToBook','trailingEps','forwardEps'
		ThisFunda2 = pd.DataFrame(ThisFunda2)
		AllFundamental2 = ThisFunda2.loc[funda2]
	except Exception:
		pass
	return AllFundamental2

def myDropStr(df, colname):
	myD = df.drop(df[df[colname].apply(lambda x: isinstance(x, str))].index)
	return myD

def myHighLow(df, hl):
	try:
		df[hl]
	except Exception:
		pass
	return df[hl]

def DailyReturn(theList):
	A, B, C, D = split_list_square(theList)

	ThisPriceChgA = myPriceChg(A)
	ThisPriceChgB = myPriceChg(B)
	ThisPriceChgC = myPriceChg(C)
	ThisPriceChgD = myPriceChg(D)

	ThisPriceChg = []
	ThisPriceChg = pd.DataFrame(ThisPriceChg)
	ThisPriceChg = pd.concat([ThisPriceChg, ThisPriceChgA, ThisPriceChgB, ThisPriceChgC, ThisPriceChgD])
	ThisPriceChg.columns=['OneDayReturn']

    ####### Drop the row of ticker without 1-day return ########

	AllPricesChg = myDropStr(ThisPriceChg, 'OneDayReturn')	    ### Drop if 1D return contains string
	AllPricesChg = AllPricesChg.sort_values(by='OneDayReturn', ascending=False)	   # Sort by max % chg

	return AllPricesChg

def GOC_Strategy(theList):
	listLen = len(theList)
	myAllDays1 = myPriceHist(theList, myDay[0])

	myAllDays2 = myDropStr(myAllDays1, 'high')
	myAllDays = myDropStr(myAllDays2, 'low')

	if listLen == 1:
		AllDays27 = myAllDays.tail(ConcDay[0]+1)
	else:
		AllDays27 = myAllDays.groupby(level=0).tail(ConcDay[0]+1)

	AllDaysHigh27 = myHighLow(AllDays27, 'high')
	AllDaysLow27 = myHighLow(AllDays27, 'low')

### 26 ###

#	AllDaysClose26_T_1 = AllDaysClose27.groupby(level=0).head(26)       # T-1 = Top26 from 27 days
	if listLen == 1:
		AllDaysHigh26_T_1 = AllDaysHigh27.head(ConcDay[0])
		AllDaysLow26_T_1 = AllDaysLow27.head(ConcDay[0])
		AllDaysHighestHigh26_T_1 = AllDaysHigh26_T_1.max()
		AllDaysLowestLow26_T_1 = AllDaysLow26_T_1.min()
	else:
		AllDaysHigh26_T_1 = AllDaysHigh27.groupby(level=0).head(ConcDay[0])
		AllDaysLow26_T_1 = AllDaysLow27.groupby(level=0).head(ConcDay[0])
		AllDaysHighestHigh26_T_1 = AllDaysHigh26_T_1.max(level=0)
		AllDaysLowestLow26_T_1 = AllDaysLow26_T_1.min(level=0)

	Base_T_1 = (AllDaysHighestHigh26_T_1 + AllDaysLowestLow26_T_1) / 2	################ Base T - 1

#	AllDaysClose26_T = AllDaysClose27.groupby(level=0).tail(26)         # T = Bottom26 from 27 days
	if listLen == 1:
		AllDaysHigh26_T = AllDaysHigh27.tail(ConcDay[0])
		AllDaysLow26_T = AllDaysLow27.tail(ConcDay[0])
		AllDaysHighestHigh26_T = AllDaysHigh26_T.max()
		AllDaysLowestLow26_T = AllDaysLow26_T.min()
	else:
		AllDaysHigh26_T = AllDaysHigh27.groupby(level=0).tail(ConcDay[0])
		AllDaysLow26_T = AllDaysLow27.groupby(level=0).tail(ConcDay[0])
		AllDaysHighestHigh26_T = AllDaysHigh26_T.max(level=0)
		AllDaysLowestLow26_T = AllDaysLow26_T.min(level=0)

	Base_T = (AllDaysHighestHigh26_T + AllDaysLowestLow26_T) / 2		################ Base T

### 9 ###

#	AllDaysClose10 = AllDaysClose27.groupby(level=0).tail(10)
	if listLen == 1:
		AllDaysHigh10 = AllDaysHigh27.tail(BaseDay[0]+1)
		AllDaysLow10 = AllDaysLow27.tail(BaseDay[0]+1)
	else:
		AllDaysHigh10 = AllDaysHigh27.groupby(level=0).tail(BaseDay[0]+1)
		AllDaysLow10 = AllDaysLow27.groupby(level=0).tail(BaseDay[0]+1)

#	AllDaysClose9_T_1 = AllDaysClose10.groupby(level=0).head(9)         # T-1 = Top9 from 10 days
	if listLen == 1:
		AllDaysHigh9_T_1 = AllDaysHigh10.head(BaseDay[0])
		AllDaysLow9_T_1 = AllDaysLow10.head(BaseDay[0])
		AllDaysHighestHigh9_T_1 = AllDaysHigh9_T_1.max()
		AllDaysLowestLow9_T_1 = AllDaysLow9_T_1.min()
	else:
		AllDaysHigh9_T_1 = AllDaysHigh10.groupby(level=0).head(BaseDay[0])
		AllDaysLow9_T_1 = AllDaysLow10.groupby(level=0).head(BaseDay[0])
		AllDaysHighestHigh9_T_1 = AllDaysHigh9_T_1.max(level=0)
		AllDaysLowestLow9_T_1 = AllDaysLow9_T_1.min(level=0)

	Conv_T_1 = (AllDaysHighestHigh9_T_1 + AllDaysLowestLow9_T_1) / 2	################ Conv T - 1

#	AllDaysClose9_T = AllDaysClose10.groupby(level=0).tail(9)           # T = Bottom9 from 10 days
	if listLen == 1:
		AllDaysHigh9_T = AllDaysHigh10.tail(BaseDay[0])
		AllDaysLow9_T = AllDaysLow10.tail(BaseDay[0])
		AllDaysHighestHigh9_T = AllDaysHigh9_T.max()
		AllDaysLowestLow9_T = AllDaysLow9_T.min()
	else:
		AllDaysHigh9_T = AllDaysHigh10.groupby(level=0).tail(BaseDay[0])
		AllDaysLow9_T = AllDaysLow10.groupby(level=0).tail(BaseDay[0])
		AllDaysHighestHigh9_T = AllDaysHigh9_T.max(level=0)
		AllDaysLowestLow9_T = AllDaysLow9_T.min(level=0)

	Conv_T = (AllDaysHighestHigh9_T + AllDaysLowestLow9_T) / 2			################ Conv T

	#print("Buy if Conv > Base on T-day; and Conv < Base on T-1 day")

	Diff_T = Conv_T - Base_T
	Diff_T_1 = Conv_T_1 - Base_T_1


	if listLen == 1:
		if Diff_T > 0 and Diff_T_1 < 0:
			GOC_result = pd.DataFrame({'T':Diff_T, 'T_1':Diff_T_1, 'GOC_signal':'True'}, index=theList)
		else:
			GOC_result = []

	else:
		GOC = pd.concat([Diff_T, Diff_T_1], axis=1)
		GOC = pd.DataFrame(GOC)
		GOC.columns = ['T','T_1']

		GOC['GOC_signal'] = GOC.apply(lambda row: (row.T > 0).any() and (row.T_1 < 0).any(), axis=1)
		GOC_result = GOC.loc[GOC['GOC_signal'] == True]

	pp.pprint(GOC_result)
	return GOC_result


def myGOC(tickers, ticName):



	AllPricesChg = DailyReturn(tickers)							### DailyReturn Strategy #############

	print(f"=== 01: Top {myDay[1]} Daily returns ===================================")
	pp.pprint(AllPricesChg)

	print("--- NewTickerList (drop 1-day return string) ---")
	newTickerList = list(AllPricesChg.index)                    # New Ticker List ####################
	pp.pprint(newTickerList)
	print(f"No. of tickers: {len(newTickerList)}")

	#AllPricesChg.to_excel(writer, sheet_name='PriceChg', merge_cells=False)

    ### Top 5 ###

	Top5R = AllPricesChg.head(myTic[0])                         # Top 5 1-day return from Index

	Top5R_Tickers = list(Top5R.index)
	Top5_Sector = Ticker(Top5R_Tickers)

	print("--- Top5 Sector ---")
	pp.pprint(Top5R_Tickers)

	#Top5_Sector = Top5_Sector.summary_profile
	#Top5_Sector = pd.DataFrame(Top5_Sector)
	#Top5_Sector = Top5_Sector.loc['sector']                     # Sector from top 5 1-day return tickers
	#Top5R_n_Sector = pd.concat([Top5R, Top5_Sector], axis=1)    # Concat daily return and sector
	#pp.pprint(Top5R_n_Sector)
	#Top5R_n_Sector.to_excel(writer, sheet_name = 'DailyR', merge_cells=False)

	print(f"=== Tickers price series ======================================")

###################################################################################
### Strategy: Ichimoku Kinko Hyo equilibrium                                    ###
### Buy if Conv > Base on T-day; and Conv < Base on T-1 day                     ###
### Conv=(Highest High over Period1 + Lowest Low over Period1) divided by 2     ###
### Base=(Highest High over Period2 + Lowest Low over Period2) divided by 2     ###
### Period1 = 9 & Period2 = 26                                                  ###
###################################################################################

	GOC_result = GOC_Strategy(newTickerList)					### GOC Strategy ######################
	'''
	A, B, C, D = split_list_square(newTickerList)

	GOC_A = GOC_Strategy(A)
	GOC_B = GOC_Strategy(B)
	GOC_C = GOC_Strategy(C)
	GOC_D = GOC_Strategy(D)

	GOC_result = []
	GOC_result = pd.DataFrame(GOC_result)
	GOC_result = pd.concat([GOC_result, GOC_A , GOC_B, GOC_C, GOC_D])
	'''
	GOC_ResultTickers = list(GOC_result.index)
	print("--- Actual no. of GOC ---")
	GOC_len = len(GOC_result.index)
	pp.pprint(GOC_len)

	if not GOC_ResultTickers:
		print("No GOC tickers")
	else:

		##### GOC Fundamental #####

		funda1 = (['volume','beta','trailingPE'])
		funda2 = (['bookValue','priceToBook','trailingEps','forwardEps'])
		Fundamentall = myFundamental1(GOC_ResultTickers, funda1)
		Fundamental2 = myFundamental2(GOC_ResultTickers, funda2)

		Funda = Fundamentall.append(Fundamental2)
		Funda = Funda.transpose()

		GOC_DailyR = AllPricesChg.loc[GOC_ResultTickers]				# All GOC daily return
		GOC_DailyRTop5 = GOC_DailyR.sort_values(by=['OneDayReturn'], ascending=False).head(5)
		GOC_Funda = pd.concat([GOC_DailyR, Funda], axis=1)				# 5 GOC with highest daily return
		GOC_Funda = GOC_Funda.sort_values(by=['OneDayReturn'], ascending=False)
		pp.pprint(GOC_Funda)

		GOC_Funda.to_excel(writer, sheet_name = ticName, merge_cells=False)

		##### GOC Series #####

		GOC_tickers = myPriceHist(GOC_ResultTickers, myDay[4])

#		GOC_tickers = AllDays.loc[GOC_Tickers]					# Select specific rows from dataframe
#		GOC_high = GOC_tickers['high']							# Select specific col from dataframe

		GOC_high9 = GOC_tickers.high.rolling(BaseDay[0]).max()			# Select high column and max in rolling period
		GOC_low9 = GOC_tickers.low.rolling(BaseDay[0]).min()				# Select low  column and min in rolling period
		GOC_9Series = (GOC_high9 + GOC_low9) / 2

		GOC_high26 = GOC_tickers.high.rolling(ConcDay[0]).max()
		GOC_low26 = GOC_tickers.low.rolling(ConcDay[0]).min()
		GOC_26Series = (GOC_high26 + GOC_low26) / 2

		GOC_close = GOC_tickers['close']
		GOC_series = pd.concat([GOC_close, GOC_9Series, GOC_26Series], axis=1)
		GOC_series.columns = ['Price','9-day','26-day']

		Day26 = 26												# 26 data points will be exported

		if GOC_len == 1:
			GOC_oneticker = GOC_result.index.copy()
			GOC_series = GOC_series.tail(Day26)
			GOC_series = pd.concat([GOC_series], keys=GOC_oneticker)
		else:
			GOC_series = GOC_series.groupby(level=0,as_index=False).tail(Day26)
		pp.pprint(GOC_series)

		GOC_series.to_excel(writer, sheet_name= ticName + "GOC", merge_cells=False)
		


############################## Main #################################

myGOC(NDX,"NDX")

writer.save()
writer.close()
		
endtime = time.perf_counter()
pp.pprint(f"Execution time of this code {round(endtime,2)}")

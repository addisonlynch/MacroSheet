#! /usr/bin/env python


import os
import sys
import pandas as pd
import datetime

import json

import pandas_datareader.data as web
from pandas_datareader.data import Options
from yahoo_finance import Share
from fredapi import Fred
from openpyxl import load_workbook
from optparse import OptionParser
from openpyxl.worksheet import *
from insertrows import insert_rows_util
import openpyxl

Worksheet.insert_rows = insert_rows_util


parser = OptionParser()
parser.add_option("-v", "--file", action="store_true", dest='pretty', default=False)
(opts, args) = parser.parse_args()



#######################
# EQUITIES INPUTS

wb_filename = 'filled.xlsx'


#datapoint lists
sIndex_large =[{'ticker': 'SP500', 'source':'fred'}, {'ticker':'DJIA', 'source': 'fred'}, {'ticker':'NASDAQCOM', 'source' :'fred'},{ 'ticker':'NASDAQ100','source':'fred'}]
sIndex_mid=  [{'ticker': 'RU2000PR', 'source':'fred'}, {'ticker': 'RUMIDCAPTR', 'source':'fred'}]
sIndex_small = [{'ticker': 'WILLSMLCAP', 'source':'fred'}]

#sheet lists
stock_indicies = [{'name' : 'large_cap', 'arr' : sIndex_large, },{'name' : 'mid_cap', 'arr' :sIndex_mid}, {'name' :'small_cap', 'arr' : sIndex_small}]


sIndex_large_offset = 4
sIndex_mid_offset = 11
sIndex_small_offset = 19

volatility = [{'ticker': 'WIV', 'source':'fred'}, {'STLFSI': 'fred'}]
manufacturing = [{'ticker': 'MNFCTRIRSA', 'source':'fred'}, {'MNFCTRSMSA': 'fred'}, {'MNFCTRIMSA':"fred"}]
housing =[{'ticker': 'ASPUS', 'source':'fred'}, {'MSPNHSUS': 'fred'}, {'MSACSR':"fred"},{ 'MNMFS':'fred'}, {'HSTFCM':"fred"},{ 'HSTFC':'fred'}, {'HSTFFHAI':"fred"},{ 'HSTFVAG':'fred'}, {'CSUSHPINSA':"fred"}]
fredKey = 'eb701cc93a5ec9bc22987c681c138b12'


sheet_list = [{'name' : 'stock_indicies', 'list' : stock_indicies }]

########################


def print_x(msg=''):
	if(msg=='') | opts.pretty == False:
		return
	else:
		print(msg)


def dict_to_datapoints(items=[]):
	dpoints=[]
	for d in items:
		vl = d['ticker']
		sc = d['source']
		dp = datapoint(vl,sc)
		dpoints.append(dp)

	return(dpoints)


#########################



class wbook(object):

	def __init__(self, filename, masterList=[]):
		self.m_sheets = []
		for sheetx in masterList:
			self.m_sheets.append(sheet(sheetx['name'], sheetx['list']))
		self.m_nsheets = len(self.m_sheets)
		print_x('sheets: ' +str(self.m_nsheets))
		self.filename = filename
		self.masterList = masterList
		self.book = load_workbook(filename)
		self.writer = pd.ExcelWriter(filename, engine='openpyxl')
		self.writer.book = self.book
		self.writer.sheets = dict((ws.title,ws) for ws in self.book.worksheets)


class sheet(object):

	def __init__(self, name, components=[]):
		self.m_components = []
		self.name = name
		temp = components
		for componentx in temp:
			new = component(componentx['name'], componentx['arr'])
			component.m_sheetName = self.name
			self.m_components.append(new)

		self.m_ncomponents = len(self.m_components)
		self.m_name = name 
		self.dframes=[]
		print_x('sheet ' + str(self.m_name) + ' components: ' +str(self.m_ncomponents))



class component(object):

	def __init__(self, name,lists=[]):
		self.m_datapoints = []
		self.m_datapoints = dict_to_datapoints(lists)
		self.m_ndatapoints = len(self.m_datapoints)
		self.m_name = name
		self.m_offset = 0
		self.m_dframe = pd.DataFrame()
		print_x('component ' + str(self.m_name) + ' datapoints: ' + str(self.m_ndatapoints))

class datapoint(object):

	def __init__(self, symbol, source):
		self.sym = symbol
		self.source = source
		if(self.source == 'fred'):
			self.fred = Fred(fredKey)
			self.fred_info(symbol)

	def fred_info(self, symbol):
		info = self.fred.get_series_info(symbol)
		self.title = info.title
		self.frequency = info.frequency_short.lower()
		self.units = info.units_short.lower()






#########################


class bot(object):
	

	def __init__(self, workbook):
		self.fred = Fred(api_key=fredKey)
		print("DailyUpdate\nv1.0\n...")
		self.wb = workbook 


	def dailyUpdate(self):
		self.updateSheets()
		self.updateExcel()
		self.jsonDump()

	def jsonDump(self):
		with open('result2.json', 'w') as fp:
			json.dump(self.wb.masterList, fp)



	def fredDownload(self, keyList=[]):

		#init dictionary array to hold series
		df = {}
		#create numPy series for each key in the given key list, add to array df
		for key in keyList:
			df[key] = self.fred.get_series(key, observation_start='2017-01-01', observation_end='2017-06-27')
		#create a data frame from array df
		df = pd.DataFrame(df)
		#sort descending by date
		sorted = df.sort_index(axis=0, ascending=False)
		print(sorted)
		return(sorted)

	def dataReader(self, service, ticker):
		start = datetime.datetime(2017,1,1)
		end = datetime.datetime(2017,1,27)
		df = web.DataReader(ticker, service, start, end)
		return(df)

	def update_handler(self, datapoint):
		dp = datapoint
		df = self.dataReader(dp.source, dp.sym)
		return(df)

	def updateExcel(self):

		for sheet in self.wb.m_sheets:
			for component in sheet.m_components:
				self.excel_write_exist(component)
		return

	def updateSheets(self):

		for sheet in self.wb.m_sheets:
			self.updateComponents(sheet)
			#arr.append({'name' : sheet.m_name, 'vals' : self.runComponents(sheet), 'offset' : total })
			#i = i+
		return


	def updateComponents(self, sheet):
		arr = []
		total =0 

		for component in sheet.m_components:
			self.updateDatapoints(component)
			component.m_offset = total + 4
			print('component: ' + component.m_name + ' offset: ' + str(total+4))
			total = total + component.m_ndatapoints + 3


		return


	def updateDatapoints(self, component):
		df = pd.DataFrame()

		for datapoint in component.m_datapoints:
			dfx = self.update_handler(datapoint)
			df = pd.concat([df, dfx], axis=1)


		component.m_dframe = df
		return


	def excel_write_exist(self, component):

		component.m_dframe.to_excel(self.wb.writer, index=False, header=False, sheet_name=component.m_sheetName, startrow=22, startcol=component.m_offset)
		self.wb.writer.save()
		return

	def excel_write_new(self, dFrame, sheetName,row, col):
		df = dFrame
		writer = pd.ExcelWriter('pandas_simple.xlsx', engine= 'xlsxwriter')
		df.to_excel(writer, sheet_name=sheetName, startrow=row, startcol=col)
		writer.save()



def main():

	wb = wbook(wb_filename, sheet_list)
	bot1 = bot(wb)
	bot1.dailyUpdate()


if __name__ == "__main__":
	main()






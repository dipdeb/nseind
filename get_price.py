"""
	This script displays the NSE stocks.
	It uses the nsetools library
"""

import json, sys
from nsetools import Nse
from pprint import pprint
from prettytable import PrettyTable
from xlsxwriter import Workbook
from string import ascii_lowercase
from multiprocessing.dummy import Pool as ThreadPool

nse = Nse()

with open('all_stock_codes.json', 'r') as f:
    codes = json.load(f)

def show_top(top_type, typ):
	t = PrettyTable(['Name', 'Open', 'High', 'Low', 'LTP', 'PreviousPrice', '52wk H-L'])
	t.title = 'Top ' + typ
	t.align	= "r"

	for tg in top_type:
		quote = nse.get_quote(tg['symbol'])
		t.add_row([codes[tg['symbol']] + ' (' + tg['symbol'] + ')', tg['openPrice'], 
						tg['highPrice'], tg['lowPrice'], tg['ltp'], tg['previousPrice'], 
						str(quote['high52'])+'-'+str(quote['low52'])])
	print(t)

def show_quote(wlist):
	with open('watchlist.json', 'r') as f:
		codes = json.load(f)

	t = PrettyTable(['Name', 'Open', 'High', 'Low', 'Close', 'LTP', 'VWAP', 
					'PreviousPrice', '52wk H-L'])
	t.title = "Watchlist"
	t.align	= "r"

	items = codes[wlist]
	pool = ThreadPool(4)
	quotes = pool.map(nse.get_quote, items)

	for q in quotes:
		t.add_row([q['symbol'], q['open'], q['dayHigh'], q['dayLow'], q['closePrice'], 
					q['lastPrice'], q['averagePrice'], q['previousClose'], 
					str(q['high52'])+' (' + q['cm_adj_high_dt']+') - '+str(q['low52'])+
					' (' + q['cm_adj_low_dt']+')'])

	print(t)

def show_portfolio(xls=False):
	with open('portfolio.json', 'r') as f:
		securities = json.load(f)

	header = ['Name', 'Open', 'High', 'Low', 'Close', 'LTP',
                    'PreviousPrice', '52wk H-L', 'PriceBoughtAt', 'CMP']
	t = PrettyTable(header)
	t.title = "Portfolio"

	quotes = []
	for sec in securities:
		code = sec['code'].upper()
		q = nse.get_quote(code)
		quotes.append([q['symbol'], q['open'],  q['dayHigh'], q['dayLow'], q['closePrice'], q['lastPrice'], q['previousClose'], str(q['high52'])+' (' + q['cm_adj_high_dt']+') - '+str(q['low52'])+
            ' (' + q['cm_adj_low_dt']+')', round(float(sec['bought'])*int(sec['qty']), 2), round(float(q['lastPrice'])*int(sec['qty']), 2)])

		t.add_row([q['symbol'], q['open'], q['dayHigh'], q['dayLow'], q['closePrice'], 
			q['lastPrice'], q['previousClose'], 
			str(q['high52'])+' (' + q['cm_adj_high_dt']+') - '+str(q['low52'])+
			' (' + q['cm_adj_low_dt']+')', round(float(sec['bought'])*int(sec['qty']), 2), round(float(q['lastPrice'])*int(sec['qty']), 2)])

	if xls is True:
		workbook = Workbook('example.xlsx') 
		align = workbook.add_format({'align': 'center', 'bold': True})
		worksheet = workbook.add_worksheet("Portfolio1") 
		# Widen the first column to make the text clearer.
		worksheet.set_column('A:A', 15)
		header_len = len(header)

		for i in range(0, header_len):
			worksheet.write(0, i, header[i], align) 

		# Iterate over the data and write it out row by row. 
		tot_rows = len(quotes)
		tot_cols = len(quotes[0])
		row = 1
		
		for quote in quotes:
			for i in range(0, tot_cols): 
				worksheet.write(row, i, quote[i]) 
			row += 1

		sc = chr(96+tot_cols)
		sc = '=SUM('+sc+'2:'+sc+str(tot_rows+1)+')'
		worksheet.write_formula(tot_rows+1, tot_cols-1, sc) 

		workbook.close() 
	print(t)

if __name__ == '__main__':
	import argparse

	parser = argparse.ArgumentParser(description='NSE utility ')
	parser.add_argument("-g", "--gain", action='store_true', help="get_price.py -g")
	parser.add_argument("-l", "--loss", action='store_true', help="commands")
	parser.add_argument("-p", "--portfolio", nargs='?', const="f", help="get_price.py -p [xl] ; t - to create xls")
	parser.add_argument("-q", "--quote", action='store_true', help="commands")
	if len(sys.argv)==1:
		parser.print_help(sys.stderr)
		sys.exit(1)
	args = parser.parse_args()

	if args.gain: 
		show_top(nse.get_top_gainers(), 'Gainers')

	if args.loss:
		show_top(nse.get_top_losers(), 'Losers')

	if args.portfolio:
		if args.portfolio == 'xl':
			show_portfolio(True)
		else:
			show_portfolio()

	if args.quote:
		show_quote('watchlist1')

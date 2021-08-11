import pandas as pd
import numpy as np
import math
import xlsxwriter
import requests
from scipy import stats

#Import List of stocks and API token
stocks= pd.read_csv('sp_500_stocks.csv')
#print(stocks)

from secrets import IEX_CLOUD_API_TOKEN
#first API Call
symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
#print(data)

#Parse Api call

price = data['latestPrice']
pe_ratio = data['peRatio']
#print(pe_ratio)

#Batch Api Call
def  chunks(lst, n):  #function to chunk list in seperate list of 100
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

#splits into lists of 100
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = [] # empty list symbol string is a ist of string where each is a comma seperated string of all stocks in object(symbol_groups)
#print(symbol_groups)

#loop through every panda series within the list and executes batch api call and  for every stock in list append info for every stock in list to pandas dataframe
for i in range(0, len(symbol_groups)):
    #print(symbol_groups[i])
    symbol_strings.append(','.join(symbol_groups[i]))
    #print(symbol_strings[i])

my_columns = ['Ticker', 'Stock Price', 'Price-to-Earnings Ratio', 'Number of Shares to Buy']
#print(my_columns)

final_dataframe= pd.DataFrame(columns=my_columns)
#print(final_dataframe)

#loop over symbol string in 'symbol_strings'
for symbol_string in symbol_strings:
    batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()
    #print(data)
    for symbol in symbol_string.split(','):

        final_dataframe= final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['peRatio'],
                    'N/A'
                ],
                index= my_columns
            ),
            ignore_index= True
        )
#print(final_dataframe)

#Removing Glamour Stocks

final_dataframe.sort_values('Price-to-Earnings Ratio', ascending=True, inplace=True)
final_dataframe = final_dataframe[final_dataframe['Price-to-Earnings Ratio'] > 0] #remove stocks with negative peRatio from dataframe
final_dataframe.reset_index(inplace=True, drop=True)
final_dataframe = final_dataframe[:100]
#print(final_dataframe)

#Calculating number of shares to buy
def portfolio_input():
    global portfolio_size
    portfolio_size= input('Enter the size of your portfolio:')

    try:
        float(portfolio_size)
    except ValueError:

        print("This is not a number! \n Please try again:")
        portfolio_size = input('Enter the size of your portfolio:')

portfolio_input()
position_size= float(portfolio_size) / len(final_dataframe.index)

for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / final_dataframe.loc[i, 'Stock Price'])
print(final_dataframe)

#Better Quantitative Value Strategy
"""A composite basket of valuation will be used to build a strong quantitative value strategy"""
"""The lowest percentiles from the following metrics will be used.
Price-to-earnings Ratio, Price-to-book-Ratio, Price-to-Sales Ratio,Enterprise Value divided by Earnings Before Interest, Taxes, Depreciation and Amortization(EV/EBITDA)
Enterprise Value Divided by Gross Profit(EV/GP)
Price-Earnings-Growth Ratio(PEG)
"""
symbol = 'AAPL'
batch_api_url=  f'https://sandbox.iexapis.com/stable/stock/market/batch?types=quote,advanced-stats&symbols={symbol}&token={IEX_CLOUD_API_TOKEN}'
data= requests.get(batch_api_url).json()

#Price-to-Earnings Ratio
pe_ratio = data[symbol]['quote']['peRatio']

#Price-to-Book Ratio
pb_ratio =  data[symbol]['advanced-stats']['priceToBook']

#Price-to-Sales Ratio
ps_ratio = data[symbol]['advanced-stats']['priceToSales']

#Enterprise Value
enterprise_value= data[symbol]['advanced-stats']['enterpriseValue']

#Gross-Profit
gross_profit = data[symbol]['advanced-stats']['grossProfit']
#Earnings Before Interest, Taxes, Depreciation and Amortization
ebitda = data[symbol]['advanced-stats']['EBITDA']

#Enterprise Value Divided by Gross Profit(EV/GP)
ev_to_gp = enterprise_value/gross_profit

#Enterprise Value divided by Earnings Before Interest, Taxes, Depreciation and Amortization(EV/EBITDA)
ev_to_ebitda= enterprise_value/ebitda

#Price-Earnings-Growth Ratio
peg_ratio= data[symbol]['advanced-stats']['pegRatio']

qv_columns= [
    'Ticker',
    'Stock Price',
    'Number of Shares to Buy',
    'Price-to-Earnings Ratio',
    'PE Percentile',
    'Price-to-Book Ratio',
    'PB Percentile',
    'Price-to-Sales Ratio',
    'PS Percentile',
    'EV/EBITDA',
    'EV/EBITDA Percentile',
    'EV/GP',
    'EV/GP Percentile',
    'PEG Ratio',
    'PEG Percentile',
    'QV Score'
]

qv_dataframe = pd.DataFrame(columns=qv_columns)

for symbol_string in symbol_strings:
    batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?types=quote,advanced-stats&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()

    for symbol in symbol_string.split(','):
        # Enterprise Value
        enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']

        # Gross-Profit
        gross_profit = data[symbol]['advanced-stats']['grossProfit']

        # Earnings Before Interest, Taxes, Depreciation and Amortization
        ebitda = data[symbol]['advanced-stats']['EBITDA']

        #fill in none type values with null in ev/ebidta
        try:
            ev_to_ebitda = enterprise_value / ebitda
        except TypeError:
            np.NaN
        # fill in none type values with null in ev/gross profit
        try:
            ev_to_grossProfit = enterprise_value / gross_profit
        except TypeError:
            np.NaN

        qv_dataframe= qv_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    'N/A',
                    data[symbol]['quote']['peRatio'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToBook'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToSales'],
                    'N/A',
                    ev_to_ebitda,
                    'N/A',
                    ev_to_grossProfit,
                    'N/A',
                    data[symbol]['advanced-stats']['pegRatio'],
                    'N/A',
                    'N/A'
                ],
                index= qv_columns
            ),
            ignore_index= True
        )

#print(qv_dataframe)

#Dealing with missing Data in dataframe

#filtering any part of the dataframe where isnull() is True
#qv_dataframe2= len(qv_dataframe[qv_dataframe.isnull().any(axis=1)].index)
#print(qv_dataframe2)

"""Fill in missing data with average non null datapoint from the column"""
#loop over every column in the dataframe
for column in ['Price-to-Earnings Ratio', 'Price-to-Book Ratio', 'Price-to-Sales Ratio','EV/EBITDA','EV/GP','PEG Ratio']:
    qv_dataframe[column].fillna(qv_dataframe[column].mean(), inplace=True)

    #print(qv_dataframe)

#Calculating Value Percentiles
metrics = {
    'Price-to-Earnings Ratio': 'PE Percentile',
    'Price-to-Book Ratio': 'PB Percentile',
    'Price-to-Sales Ratio': 'PS Percentile',
    'EV/EBITDA': 'EV/EBITDA Percentile',
    'EV/GP': 'EV/GP Percentile',
    'PEG Ratio': 'PEG Percentile'
}
for metric in metrics.keys():
    for row in qv_dataframe.index:
        qv_dataframe.loc[row, metrics[metric]] = stats.percentileofscore(qv_dataframe[metric], qv_dataframe.loc[row,metric]) / 100

#CALCULATING THE QV SCORE

from statistics import mean

for row in qv_dataframe.index:
    value_percentiles= []

    for metric in metrics.keys():
        value_percentiles.append(qv_dataframe.loc[row, metrics[metric]])
    qv_dataframe.loc[row, 'QV Score']= mean(value_percentiles)

#SELECTING 100 BEST VALUE STOCKS
qv_dataframe.sort_values('QV Score', ascending= False, inplace= True)
qv_dataframe= qv_dataframe[:100]

#reset index
qv_dataframe.reset_index(inplace=True, drop=True)

#CALCULATING NUMBER OF SHARES TO BUY
portfolio_input()

position_size = float(portfolio_size) / len(qv_dataframe)

for i in range(0, len(qv_dataframe.index)):
    qv_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/qv_dataframe.loc[i, 'Stock Price'])

print(qv_dataframe)


#FORMATTING EXCEL OUTPUT

writer = pd.ExcelWriter('Best Value Trades.xlsx', engine='xlsxwriter')
qv_dataframe.to_excel(writer, 'Best Value Trades', index=False)

#Background and font color formants
background_color = '00008B'
font_color = 'FFFFFF'

string_format= writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_format= writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

percent_format = writer.book.add_format(
    {
        'num_format': '%0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1

    }
)

decimal_format = writer.book.add_format(
    {
        'num_format': '0.0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1

    }
)
integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Number of Shares to Buy', integer_format],
    'D': ['Price-to-Earnings Ratio', decimal_format],
    'E': ['PE Percentile', percent_format],
    'F': ['Price-to-Book Ratio', decimal_format],
    'G': ['PB Percentile', percent_format],
    'H': ['Price-to-Sales Ratio', decimal_format],
    'I': ['PS Percentile', percent_format],
    'J': ['EV/EBITDA', decimal_format],
    'K': ['EV/EBITDA Percentile', percent_format],
    'L': ['EV/GP', decimal_format],
    'M': ['EV/GP Percentile', percent_format],
    'N': ['PEG Ratio', decimal_format],
    'O': ['PEG Percentile', percent_format],
    'P': ['QV Score', percent_format]
}

for column in column_formats.keys():
    writer.sheets['Best Value Trades'].set_column(f'{column}:{column}',25,column_formats[column][1])
    writer.sheets['Best Value Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()

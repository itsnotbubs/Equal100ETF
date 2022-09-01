import numpy as np
import pandas as pd
import requests
import xlsxwriter 
import math

# Load S&P Data Frame Tickers
df = pd.read_csv('Stocks in the SP 500 Index.csv')
stocks = df['Symbol']
 
# Import IEX Cloud API

from sec import IEX_CLOUD_API_TOKEN

#API Call example for single stock data

symbol = 'AAPL'
api_url = f'https://cloud.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()

#Make sure API request is accurate * remove .json
# print(data.status_code == 200)

#Retrieving price and market cap 
price = data['latestPrice']
market_cap = data['marketCap']


#Create data frame 
my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns=my_columns)

#For loop to loop through all 505 S&P companies to create our dateframe 
for stock in stocks[:100]:
    api_url = f'https://cloud.iexapis.com/stable/stock/{stock}/quote/?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(api_url).json()
    final_dataframe = final_dataframe.append(pd.Series([stock, data['latestPrice'], data['marketCap'], 'N/A'], index=my_columns), ignore_index=True)



#interactive portfolio input to determine share amount per holding 
portfolio_size = input('Enter the value of your portfolio:')
try:
    val = float(portfolio_size)
except: 
    print("That's not a number! \nPlease try again:")
    portfolio_size = input('Enter the value of your portfolio:')
    val = float(portfolio_size)


position_size = val/len(final_dataframe.index)

for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

final_dataframe

writer = pd.ExcelWriter('recommended_allocation.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, sheet_name='Recommended Allocation', index = False)


#Colors for excel file
background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )


column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Allocation'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Allocation'].write(f'{column}1', column_formats[column][0], string_format)



writer.save()
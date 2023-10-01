import numpy as np
import pandas as pd
import yfinance as yf
import requests
import xlsxwriter
import math
from bs4 import BeautifulSoup
import csv 

#Webscrape from wiki the S&P 500 stock tickers
url = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
response = requests.get(url)

soup = BeautifulSoup(response.text, 'html.parser')

table = soup.find('table', {'class': 'wikitable sortable'})
rows = table.find_all('tr')[1:]  # Skip the header row

sp_500_companies = []

for row in rows:
    cols = row.find_all('td')
    ticker = cols[0].text.strip()
    sp_500_companies.append(ticker)

# Save the list to a CSV file
with open('sp_500_companies.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Ticker'])
    for company in sp_500_companies:
        writer.writerow([company])

#read csv
stocks = pd.read_csv('sp_500_companies.csv')
stocks["Ticker"] = stocks["Ticker"].str.replace('.', '-')

# Set up dataframe
my_columns = ["Ticker", "Price", "Market Cap", "Number of Shares to Buy"]
final_dataframe = pd.DataFrame(columns = my_columns)

tickers = stocks["Ticker"].to_list()

# Loop through stocks
for i in range(0, len(tickers), 10):
    batch_tickers = tickers[i:i + 10]
    tickers_data = yf.Tickers(' '.join(batch_tickers))

    for stock in tickers_data.tickers:
        try:
            current_price = tickers_data.tickers[stock].info["currentPrice"]
            market_cap = tickers_data.tickers[stock].info["marketCap"]
            if current_price is not None and market_cap is not None:
                new_row = [stock, current_price, market_cap, 'N/A']
            else:
                new_row = [stock, 'Data not available', 'Data not available', 'N/A']
        except Exception as e:
            print(f"Error occurred for {stock}: {e}")
            new_row = [stock, "Data not available", "Data not available", "N/A"]

        final_dataframe = pd.concat([final_dataframe, pd.DataFrame([new_row], columns=my_columns)], ignore_index=True, axis=0)

#Calculating Number of Shares to Buy
portfolio_value = input("Please enter the value of your portfolio:")
try:
    value = float(portfolio_value)
except ValueError:
    print("That is not a number. \n Please input a number:")
    portfolio_value = input("Please enter the value of your portfolio:")

total_market_value = final_dataframe["Market Cap"].sum()
for i in range(0, len(final_dataframe['Ticker'])):
    ticker_market_cap = final_dataframe.loc[i, "Market Cap"]

    if ticker_market_cap != 'Data not available':
        position_size = (ticker_market_cap / total_market_value) * float(portfolio_value)
        final_dataframe.loc[i, "Number of Shares to Buy"] = position_size
    else:
        final_dataframe.loc[i, "Number of Shares to Buy"] = 'N/A'

print(final_dataframe)


#Formatting Excel Output
writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

background_colour = '#0000FF'
font_colour ='#FFFFFF'

string_format = writer.book.add_format(
    {
    'font_color': font_colour,
    'bg_color': background_colour,
    'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
    'num_format': '$0.00',
    'font_color': font_colour,
    'bg_color': background_colour,
    'border': 1
    }
)

integer_format = writer.book.add_format(
    {
    'num_format': '0',
    'font_color': font_colour,
    'bg_color': background_colour,
    'border': 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.close()
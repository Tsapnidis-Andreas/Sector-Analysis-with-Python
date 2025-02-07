import yfinance as yf
import pandas as pd
import datetime
import numpy as np
import pyodbc

def get_data(stock):
    end_date = datetime.datetime(2024, 9, 30)
    start_date = datetime.datetime(2019, 9, 30)
    data = yf.download(stock, start=start_date, end=end_date)
    print(data)
    data = pd.DataFrame(data)
    data = data['Adj Close']
    monthly_data = data.resample('ME').last()
    monthly_data.columns = ['Adj Close']
    ind=monthly_data.index
    return (pd.Series(monthly_data['Adj Close']),ind)


path=""
stocks=['AAPL','NVDA','MSFT','AMZN','GOOG','META','TSLA','AVGO','ORCL','NFLX']

data=pd.DataFrame()

for i in stocks:
    data[i],ind=get_data(i)

data.index=ind
print(data)
data.to_excel(path+'Stocks Closing Prices.xlsx')
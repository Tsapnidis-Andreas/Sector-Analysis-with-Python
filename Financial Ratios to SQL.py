import yfinance as yf
import pandas as pd
import datetime
import numpy as np
import pyodbc

def get_data(stock,path):
    microsoft=yf.Ticker(stock)
    dict=microsoft.info
    df=pd.DataFrame.from_dict(dict,orient='index')
    df=df.reset_index()
    ind=df['index']
    df=pd.DataFrame({'Value':df[0]})
    df.index=ind
    print(df)
    save_to_excel(df,stock,path)

    debt_to_equity=round(df.loc['debtToEquity','Value'],2)/100
    profit_margin=round(df.loc['profitMargins','Value'],2)
    pe=round(df.loc['trailingPE','Value'],2)
    beta=round(df.loc['beta','Value'],2)
    market_cap=round(df.loc['marketCap','Value'],2)
    print(debt_to_equity,profit_margin,pe,beta,market_cap)

    save_to_database(stock,debt_to_equity,profit_margin,pe,beta,market_cap)

def save_to_excel(df,stock,path):
    df.to_excel(path+stock+' data.xlsx')

def save_to_database(stock,debt_to_equity,profit_margin,pe,beta,market_cap):
    conn = pyodbc.connect(
        'Driver={SQL Server};' '' 'Database=Sector Analysis;' 'Trusted_connection=yes')
    conn.execute('INSERT INTO stock_data (stock_ticker,debt_to_equity,profit_margin,PE,beta,market_cap) VALUES(?,?,?,?,?,?)',
                 (stock, debt_to_equity, profit_margin, pe, beta,market_cap))
    conn.commit()
    conn.close()

path=""
stocks=['AAPL','NVDA','MSFT','AMZN','GOOG','META','TSLA','AVGO','ORCL','NFLX']
for i in stocks:
    get_data(i,path)
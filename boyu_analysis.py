import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
import talib as ta
from talib import abstract
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

 

### input ##########

stocks = ['^GSPC','AAPL','TSLA']
start = '2023-01-01'
end = '2024-01-01'

####################

# top_10_stock = ['AAPL', 'MSFT', 'NVDA', 'GOOGL', 'AMZN', 'META', 'BRK-B', 'LLY', 'AVGO', 'TSLA']
#index = ['^GSPC', '^IXIC', '^DJI', '^SOX']


if not os.path.exists('images'):
        os.makedirs('images')
workbook = xlsxwriter.Workbook('stock_data.xlsx')

for stock in stocks:
    ### setting time fetching data using yfnance
    worksheet = workbook.add_worksheet(stock)
    stockIndex = stock
    pre = datetime.strptime(start, '%Y-%m-%d') - relativedelta(months=2)
    pre = pre.strftime('%Y-%m-%d')
    df = yf.Ticker(stockIndex).history(start = pre, end = end)
    df.index = df.index.tz_localize(None)
    cdf = df['Close']

    ### function of drawing plots
    def draw_pic(inp):
        plt.figure(figsize=(12, 3))
        plt.plot(inp['x_val'], inp['y_val'][0], label= 'close', marker='.',markersize= 3,c = 'k')
        plt.plot(inp['x_val'], inp['y_val'][1], label= '20SMA', marker='.',markersize= 0,c = 'orange')
        plt.title(f'{stockIndex} {inp['name']}') 
        plt.xlabel('Date')
        plt.ylabel(f'{inp['ylabel']}')
        plt.grid()
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(os.path.join('images', f'{stockIndex} {inp['name']}.png')) 
        plt.close()  
        worksheet.insert_image(inp['loc'], os.path.join('images', f'{stockIndex} {inp['name']}.png'))
    def draw_pic_tec(inp):
        plt.figure(figsize=(12, 3))
        plt.plot(inp['x_val'], inp['y_val'][0], label = inp['lab'][0], marker='.',markersize= 0,color = 'blue')
        plt.plot(inp['x_val'], inp['y_val'][1], label = inp['lab'][1], marker='.',markersize= 0,color = 'hotpink')
        plt.bar(inp['x_val'], inp['y_val'][2], label = inp['lab'][2],color = ['green' if value > 0 else 'red' for value in inp['y_val'][2]])
        plt.legend()
        plt.title(f'{stockIndex} {inp['name']}') 
        plt.xlabel('Date')
        plt.ylabel(f'{inp['ylabel']}')
        plt.grid()
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(os.path.join('images', f'{stockIndex} {inp['name']}.png'))
        plt.close() 
        worksheet.insert_image(inp['loc'], os.path.join('images', f'{stockIndex} {inp['name']}.png'))
    def draw_pic_totval(inp):
        plt.figure(figsize=(28, 7))
        plt.plot([inp.index[0], inp.index[-1]], [100,100], label= '100%', ls = '--',linewidth = '5',c = 'red')
        plt.plot(inp.index, inp['st0'], label= 'buy and hold', marker='.',markersize= 0,c = 'navy')
        plt.plot(inp.index, inp['st1'], label= 'st1', marker='.',markersize= 0,c = 'brown')
        plt.plot(inp.index, inp['st2'], label= 'st2', marker='.',markersize= 0,c = 'cyan')
        plt.plot(inp.index, inp['st_sma'], label= 'sma', marker='.',markersize= 0,c = 'lime')
        plt.plot(inp.index, inp['st_macd'], label= 'macd', marker='.',markersize= 0,c = 'gray')
        plt.plot(inp.index, inp['st_kd'], label= 'kd', marker='.',markersize= 0,c = 'orange')
        plt.plot(inp.index, inp['st_rsi'], label= 'rsi', marker='.',markersize= 0,c = 'magenta')
        plt.title(f'{stockIndex} earning rate') 
        plt.xlabel('date')
        plt.ylabel('%')
        plt.grid()
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(os.path.join('images', f'{stockIndex} earning rate.png')) 
        plt.close()  
        worksheet.insert_image('U1', os.path.join('images', f'{stockIndex} earning rate.png'))

    ### fetching technical indicator using talib and abstract
    df.rename(columns={'High': 'high', 'Low': 'low', 'Close': 'close'}, inplace=True)
    sma_df = abstract.SMA(df, 20)
    sma_df.name = 'sma'
    macd_df = abstract.MACD(df, fastperiod=12, slowperiod=26, signalperiod=9)
    kd_df = abstract.STOCH(df,fastk_period=9, slowk_period=5,slowd_period=5, slowk_matype=1, slowd_matype=1)
    rsif_df = abstract.RSI(df, 5)
    rsif_df.name = 'rsif'
    rsis_df = abstract.RSI(df, 10)
    rsis_df.name = 'rsis'
    res = pd.concat([cdf, sma_df, macd_df, kd_df, rsif_df, rsis_df], axis=1)
    res = res.loc[res.index >= start]

    ### writing data into the excel file
    worksheet.write('A1', 'Date')
    worksheet.write('B1', 'Close Price')
    for idx in range(len(res)):
        worksheet.write(idx + 1, 0, res.index[idx].strftime('%Y-%m-%d')) 
        worksheet.write(idx + 1, 1, float(round(res['Close'].iloc[idx], 2)))

    ### drawing plots into the excel file
    draw_pic({
        'x_val': res.index,
        'y_val': [res['Close'],res['sma']],
        'name' : 'Closing Prices History',
        'ylabel': 'price',
        'loc': 'C1'
    })
    draw_pic_tec({
        'x_val': res.index,
        'y_val': [res['macd'], res['macdsignal'], res['macdhist']],
        'name' : 'MACD History',
        'ylabel': 'unit',
        'loc': 'C16',
        'lab': ['DIF', 'MACD', 'OSC'],
    })
    draw_pic_tec({
        'x_val': res.index,
        'y_val': [res['slowk'], res['slowd'], res['slowk'] - res['slowd']],
        'name' : 'KD History',
        'ylabel': 'unit',
        'loc': 'C31',
        'lab': ['K', 'D', 'K-D'],
    })
    draw_pic_tec({
        'x_val': res.index,
        'y_val': [res['rsif'], res['rsis'], res['rsif'] - res['rsis']],
        'name' : 'RSI History',
        'ylabel': 'unit',
        'loc': 'C46',
        'lab': ['RSI(5)', 'RSI(10)', 'RSI(5)-RSI(10)'],
    })
    
    ### developing strategies and visualizing the outcomes
    def strategy0(dframe):
        rate_data = pd.Series(dframe['Close'] / dframe['Close'].iloc[0]*100, name = 'st0')
        return rate_data
    def strategy1(dframe):
        rate = 100
        rate_data = [rate]  
        hold = 1
        for i in range(1,len(dframe['Close'])):
            if dframe['Close'].iloc[i-1] <= dframe['Close'].iloc[i]:
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                else:
                    hold = 1
            else: 
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                    hold = 0
            rate_data.append(rate)
        rate_data = pd.Series(rate_data,index = dframe.index, name = 'st1')
        return rate_data
    def strategy2(dframe):
        rate = 100
        rate_data = [rate]  
        hold = 1
        for i in range(1,len(dframe['Close'])):
            if dframe['Close'].iloc[i-1] <= dframe['Close'].iloc[i]:
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                    hold = 0
            else: 
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                else:
                    hold = 1
            rate_data.append(rate)
        rate_data = pd.Series(rate_data,index = dframe.index, name = 'st2')
        return rate_data
    def strategy_sma(dframe):
        rate = 100
        rate_data = [rate]  
        hold = 1
        for i in range(1,len(dframe['Close'])):
            if dframe['Close'].iloc[i] >= dframe['sma'].iloc[i]:
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                else:
                    hold = 1
            else: 
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                    hold = 0
            rate_data.append(rate)
        rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_sma')
        return rate_data
    def strategy_macd(dframe):
        rate = 100
        rate_data = [rate]  
        hold = 1
        for i in range(1,len(dframe['Close'])):
            if dframe['macdhist'].iloc[i] >= 0:
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                else:
                    hold = 1
            else: 
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                    hold = 0
            rate_data.append(rate)
        rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_macd')
        return rate_data
    def strategy_kd(dframe):
        rate = 100
        rate_data = [rate]  
        hold = 1
        for i in range(1,len(dframe['Close'])):
            if dframe['slowk'].iloc[i] >= dframe['slowd'].iloc[i]:
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                else:
                    hold = 1
            else: 
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                    hold = 0
            rate_data.append(rate)
        rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_kd')
        return rate_data
    def strategy_rsi(dframe):
        rate = 100
        rate_data = [rate]  
        hold = 1
        for i in range(1,len(dframe['Close'])):
            if dframe['rsif'].iloc[i] >= dframe['rsis'].iloc[i]:
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                else:
                    hold = 1
            else: 
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                    hold = 0
            rate_data.append(rate)
        rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_rsi')
        return rate_data

    totval = pd.concat([
        strategy0(res), 
        strategy1(res),
        strategy2(res),
        strategy_sma(res),
        strategy_macd(res),
        strategy_kd(res),
        strategy_rsi(res)
        ],axis = 1)

    draw_pic_totval(totval)
    print(f'{stock} is written in')

    ### opening the excel file
workbook.close()
os.system('start excel.exe "stock_data.xlsx"')





























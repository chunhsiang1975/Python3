# Writeten by Chun-Hsiang Chao
# Date:20250724
import twstock
#twstock.__update_codes()
real = twstock.realtime.get('1101')
if real['success']:
    print('即時股票資料：')
    print(real)  #即時資料
    print(real['success']) #確認傳輪是否成功
else:
    print('錯誤：' + real['rtmessage'])
print('目前股價：')
print(real['realtime']['latest_trade_price'])  #即時價格
stock = twstock.Stock('1101')
print(stock.sid) #回傳股票代號
#print(stock.price) #回傳各日收盤價
#print(stock.high) #回傳各日最高價
#print(stock.date) #回傳資料之對應日期
#print(stock.fetch(2025, 6))  # 獲取 2025 年 6 月之股票資料
#print(stock.fetch(2024, 4))  # 獲取 2024 年 4 月之股票資料
#print(stock.fetch_31())      # 獲取近 31 日開盤之股票資料
#print(stock.fetch_from(2025, 6))  # 獲取 2025 年 6 月至今日之股票資料
#print(stock.moving_average(stock.price, 5))  # 計算五日平均價格
#print(stock.moving_average(stock.capacity, 5))  # 計算五日平均交易量
#print(stock.ma_bias_ratio(5,10))  # 計算五日、十日乖離值
#BestFourPoint 四大買賣點判斷來自 toomore/grs 之中的一個功能， 透過四大買賣點來判斷是否要買賣股票。四個買賣點分別為：
#量大收紅 / 量大收黑
#量縮價不跌 / 量縮價跌
#三日均價由下往上 / 三日均價由上往下
#三日均價大於六日均價 / 三日均價小於六日均價
bfp = twstock.BestFourPoint(stock)
print(bfp.best_four_point_to_buy()) # 判斷是否為四大買點
print(bfp.best_four_point_to_sell())  # 判斷是否為四大賣點
print(bfp.best_four_point())        # 綜合判斷

stocks = twstock.realtime.get(['2330', '2337', '2409']) #多個股票同時查詢
print(stocks)
print('2330' in twstock.twse)	#查詢代號是否為上市股票
print('2330' in twstock.tpex)	#查詢代號是否為上櫃股票
print('2330' in twstock.codes)	#查詢代號是否為台灣股票代號


from datetime import datetime
import backtrader as bt
import yfinance as yf
from backtrader.indicators import Highest, Lowest
import pandas as pd
from pykrx import stock

class MyStrategy(bt.Strategy):
    def __init__(self):
        self.dataclose = self.datas[0].close
        self.order = None
        self.buyprice = None
        self.buycomm = None
        # self.rsi = bt.indicators.RSI_SMA(self.data.close, period=21)
        # self.rsi = bt.indicators.SMA(self.data1.close, period=3)
        self.change_daily = bt.indicators.Ichimoku2(self.data)
        self.change_daily_bfdt = bt.indicators.Ichimoku2(self.datas[1])
        self.change_weekly = bt.indicators.Ichimoku2(self.data1)
        self.change_weekly_bfdt = bt.indicators.Ichimoku2(self.data1)(1)
        self.sma5_daily = bt.indicators.SMA(self.data, period=5)
        self.sma5_weekly = bt.indicators.SMA(self.data1, period=5)
        self.atr = bt.indicators.AverageTrueRange(self.data)
        self.addWater = [True, True, True, True]

    def notify_order(self, order):
        if order.status in [order.Submitted, order.Accepted]:
            return
        if order.status in [order.Completed]:
            if order.isbuy():
                self.log(('BUY : 주가', order.executed.price, '수량', order.executed.size,
                          '수수료', order.executed.comm, '자산', cerebro.broker.getvalue()))
                self.buyprice = order.executed.price
                self.byucomm = order.executed.comm
            else:
                self.log(('SELL : 주가', order.executed.price, '수량', order.executed.size,
                          '수수료', order.executed.comm, '자산', cerebro.broker.getvalue()))
            self.bar_executed = len(self)
        elif order.status in [order.Canceled]:
            self.log('ORDER CANCELED')
        elif order.status in [order.Margin]:
            self.log('ORDER MARGIN')
        elif order.status in [order.Rejected]:
            self.log('ORDER REJECTED')
        self.order = None


    # # 1번 알고리즘
    # [일봉] 저가가 5이평선, 9전환선 미만 & 종가가 5이평선, 9전환선 초과
    # [주봉] 저가가 5이평선, 9전환선 미만 & 종가가 5이평선, 9전환선 초과
    # def next(self):
    #     if not self.position:
    #         if self.data.low < self.sma5_daily and self.data.low < self.change_daily \
    #                 and self.data.close > self.sma5_daily and self.data.close > self.change_daily \
    #                 and self.data1.low < self.sma5_weekly and self.data1.low < self.change_weekly \
    #                 and self.data1.close > self.sma5_weekly and self.data1.close > self.sma5_weekly:
    #                 # and bt.If(self.data.close >= self.data.high(-2), True, False):
    #
    #             # and self.data.close >= self.data.high(-3)\
    #             # and self.data.close >= self.data.high(-4)\
    #             # and self.data.close >= self.data.high(-5)\
    #             # and self.data.close >= self.data.high(-6)\
    #             # and self.data.close >= self.data.high(-7)\
    #             # and self.data.close >= self.data.high(-8)\
    #             # and self.data.close >= self.data.high(-9)\
    #             # and self.data.close >= self.data.high(-10):
    #             # and self.data.close >= Highest(self.data1.high, period=10):
    #
    #             # and self.data.open + (self.atr) < self.data.close:
    #             # and self.data.open + (self.atr * 2) < self.data.close:
    #             # and self.data.close >= Highest(self.data.close, period=10):   # 안되나? 수정 필요
    #             self.order = self.buy()
    #     else:
    #         if self.position.price + (self.atr / (2+(self.addWater.count(False)))) < self.data2.close:
    #             self.order = self.sell()
    #             self.addWater[0] = True
    #             self.addWater[1] = True
    #             self.addWater[2] = True
    #             self.addWater[3] = True
    #         if self.position.price - self.atr > self.data2.close:
    #             self.order = self.sell()
    #         if self.addWater[0]:
    #             # self.sma5_daily = bt.indicators.SMA(self.data, period=5)
    #             if self.data2.close <= (self.sma5_daily * 1.01):
    #                 self.order = self.buy()
    #                 self.addWater[0] = False
    #
    #                 # print(self.position.price, self.position.price + (self.atr / 2), self.data2.close)
    #         if self.addWater[1]:
    #             if self.data2.close <= (self.sma5_weekly * 1.01):
    #                 self.order = self.buy()
    #                 self.addWater[1] = False
    #                 # print(self.position.price, self.position.price + (self.atr / 2), self.data2.close)
    #         if self.addWater[2]:
    #             if self.data2.close <= (self.change_daily * 1.01):
    #                 self.order = self.buy()
    #                 self.addWater[2] = False
    #                 # print(self.position.price, self.position.price + (self.atr / 2), self.data2.close)
    #         if self.addWater[3]:
    #             if self.data2.close <= (self.change_weekly * 1.01):
    #                 self.order = self.buy()
    #                 self.addWater[3] = False
    #                 # print(self.position.price, str(self.atr / 2))
    #                 # print(f'[{dt.isoformat()}]')

    # 2번 알고리즘
    # 1번 알고리즘 + 전환선 상승추세 + 최근 고가 + ATR 대비 상승양 1.5배 이상
    # -> 매수 케이스가 없다
    # -> 전환선 상승추세 빼니까 많이 나옴
    # 물타기가 안되고있다. 메수를 자산의 90퍼 하는 중
    # 1년간 수익률 : 272,958,302.184994(2700%)

    # 여기서 물타기만 제대로 바꾼 뒤에 재검사(매수당 투자금액의 23프로로)
    # 1년간 수익률 : 631까지 해서 144,946,985

    # 그럼 물타기 + 손절선 수정
    # 손절선 수정 : min(5일이평, 9일전환, 5주이평, 9주이평) - ATR에 손절
    # 1년간 수익률 : 275,182,927.4029911 (631까지  148,358,685)

    # 물타기 제대로 수정
    # 1년간 수익률 :

    # 지난 한 달
    # atr 대비 1.5배 없을 때 :0c
    # atr 대비 1.5배 있을 때 :

    def next(self):
        if not self.position:
            if self.data.low < self.sma5_daily and self.data.low < self.change_daily \
                    and self.data.close > self.sma5_daily and self.data.close > self.change_daily \
                    and self.data1.low < self.sma5_weekly and self.data1.low < self.change_weekly \
                    and self.data1.close > self.sma5_weekly and self.data1.close > self.sma5_weekly:
                    # and self.data.close - self.data.open > self.atr * 1.5:
                    # and self.change_daily >= self.change_daily_bfdt\
                    # and self.change_weekly >= self.change_weekly_bfdt:
                # and bt.If(self.data.close >= self.data.high(-2), True, False):

                # and self.data.close >= self.data.high(-3)\
                # and self.data.close >= self.data.high(-4)\
                # and self.data.close >= self.data.high(-5)\
                # and self.data.close >= self.data.high(-6)\
                # and self.data.close >= self.data.high(-7)\
                # and self.data.close >= self.data.high(-8)\
                # and self.data.close >= self.data.high(-9)\
                # and self.data.close >= self.data.high(-10):
                # and self.data.close >= Highest(self.data1.high, period=10):

                # and self.data.open + (self.atr) < self.data.close:
                # and self.data.open + (self.atr * 2) < self.data.close:
                # and self.data.close >= Highest(self.data.close, period=10):   # 안되나? 수정 필요
                self.order = self.buy()
        else:
            if self.position.price + (self.atr / (2 + (self.addWater.count(False)))) < self.data2.close:
                #self.order = self.sell()
                self.close()
                self.addWater[0] = True
                self.addWater[1] = True
                self.addWater[2] = True
                self.addWater[3] = True
           # if self.position.price - self.atr > self.data2.close:
            if min(self.sma5_daily, self.change_daily, self.sma5_weekly, self.change_weekly) - self.atr > self.data2.close:
                #self.order = self.sell()
                self.close()
                self.addWater[0] = True
                self.addWater[1] = True
                self.addWater[2] = True
                self.addWater[3] = True
            if self.addWater[0]:
                # self.sma5_daily = bt.indicators.SMA(self.data, period=5)
                if self.data2.close <= (self.sma5_daily * 1.01):
                    self.order = self.buy()
                    self.addWater[0] = False

                    # print(self.position.price, self.position.price + (self.atr / 2), self.data2.close)
            if self.addWater[1]:
                if self.data2.close <= (self.sma5_weekly * 1.01):
                    self.order = self.buy()
                    self.addWater[1] = False
                    # print(self.position.price, self.position.price + (self.atr / 2), self.data2.close)
            if self.addWater[2]:
                if self.data2.close <= (self.change_daily * 1.01):
                    self.order = self.buy()
                    self.addWater[2] = False
                    # print(self.position.price, self.position.price + (self.atr / 2), self.data2.close)
            if self.addWater[3]:
                if self.data2.close <= (self.change_weekly * 1.01):
                    self.order = self.buy()
                    self.addWater[3] = False
                    # print(self.position.price, str(self.atr / 2))
                    # print(f'[{dt.isoformat()}]')




    def log(self, txt, dt=None):
        dt = self.datas[0].datetime.date(0)
        print(f'[{dt.isoformat()}] {txt}')


# stock_list = pd.DataFrame({'종목코드':stock.get_market_ticker_list(market='ALL')})
stock_list = pd.DataFrame({'종목코드':stock.get_market_ticker_list(market='KOSDAQ')})
stock_list['종목명'] = stock_list['종목코드'].map(lambda x: stock.get_market_ticker_name(x))
# print(stock_list)

# 1번 알고리즘 결과
# stock_from = "20211125"
# stock_to = "20220602"
# 1500%
# stock_from = "20210601"
# stock_to = "20211125"
# -200%
# stock_from = "20220301"
# stock_to = "20220601"
# 60%

# 2번 알고리즘
# 최근 1년
stock_from = "20220524"
stock_to = "20220624"
total = 0
cnt = 0


for stock_code in stock_list['종목코드']:
    cerebro = bt.Cerebro()
    cerebro.addstrategy(MyStrategy)
    print(cnt, stock_code)
    cnt += 1
    ticker = bt.indicators.Getkrxdata(stock_code, stock_from, stock_to)
    ticker.get_data()

    # data = ticker.get_data()
    data = bt.feeds.PandasData(dataname=ticker.df)
    cerebro.adddata(data)
    cerebro.resampledata(data, timeframe=bt.TimeFrame.Weeks)
    cerebro.resampledata(data, timeframe=bt.TimeFrame.Minutes)
    #print('Initial Portfolio Value : ', cerebro.broker.getvalue(), ' KRW')
    cerebro.broker.setcash(10000000)
    cerebro.broker.setcommission(commission=0.0014)
    cerebro.addsizer(bt.sizers.PercentSizer, percents=20)

    cerebro.run()
    #print('Final Portfolio Value : ', cerebro.broker.getvalue(), ' KRW')
    total += cerebro.broker.getvalue() - 10000000
    #cerebro.plot(style='candlestick')
    print('###########', total, '###########')
print(total)

from PyQt5.QtWidgets import *
import win32com.client
import Find_Target
import pandas as pd
import time
from datetime import datetime
from slacker import Slacker
from csv import writer
from csv import reader
import requests


objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
objIndex = win32com.client.Dispatch("CpIndexes.CpIndex")
objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
objAccStat = win32com.client.Dispatch('CpTrade.CpTd6033')
objCpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
objOrderAft1530 = win32com.client.Dispatch('CpTrade.CpTd0322')
objCheckNotConcluded = win32com.client.Dispatch('CpTrade.CpTd5339')

kosdaq = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥'

objCpTradeUtil.TradeInit()
acc = objCpTradeUtil.AccountNumber[0]  # 계좌번호
accFlag = objCpTradeUtil.GoodsList(acc, 1)  # -1:전체,1:주식,2:선물/옵션

moneyPerQty = 200000
orderQty = [1, 1, 2, 4, 8, 16]

def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )
# slack = Slacker('xoxb-3686270987524-3683946775731-COcMB0qZE2CHo1Mdj27f4rS6')

myToken = "xoxb-3686270987524-3683946775731-COcMB0qZE2CHo1Mdj27f4rS6"

def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""

    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken, "#trading", strbuf)
    # slack.chat.post_message('#trading', strbuf)


def GetData_Daily(code):
    cnt = 30

    # 일데이터용
    objStockChart.SetInputValue(0, code)  # 종목 코드 -
    objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
    objStockChart.SetInputValue(4, cnt)  # 최근 100일치
    objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
    objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
    objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
    objStockChart.BlockRequest()

    len = objStockChart.GetHeaderValue(3)
    CloseDaily = 0
    LowDaily = 0
    HighDaily = 0
    OpenDaily = 0

    for i in range(len):
        open = objStockChart.GetDataValue(1, len - i - 1)
        high = objStockChart.GetDataValue(2, len - i - 1)
        low = objStockChart.GetDataValue(3, len - i - 1)
        close = objStockChart.GetDataValue(4, len - i - 1)
        vol = objStockChart.GetDataValue(5, len - i - 1)
        CloseDaily = close
        LowDaily = low
        HighDaily = high
        OpenDaily = open
        objSeries.Add(close, open, high, low, vol)

    objIndex.series = objSeries

    retDaily = [OpenDaily, CloseDaily, LowDaily, HighDaily]
    # 지표들구하기
    indexNames = ['이동평균(라인1개)', '일목균형표', 'ATR']
    for indexName in indexNames:
        objIndex.put_IndexKind(indexName)
        objIndex.put_IndexDefault(indexName)
        # 지표 데이터 계산 하기
        objIndex.Calculate()
        cntofIndex = objIndex.GetCount(0)
        for i in range(cntofIndex - 2, cntofIndex):
            retDaily.append(objIndex.GetResult(0, i))
            # print(code, indexName, objIndex.GetResult(0, i))
    # print(retDaily)
    return retDaily


def GetData_Weekly(code):
    cnt = 20

    # 일데이터용
    objStockChart.SetInputValue(0, code)  # 종목 코드 -
    objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
    objStockChart.SetInputValue(4, cnt)  # 최근 100일치
    objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
    objStockChart.SetInputValue(6, ord('W'))  # '차트 주기 - 일간 차트 요청
    objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
    objStockChart.BlockRequest()

    len = objStockChart.GetHeaderValue(3)
    CloseDaily = 0
    LowDaily = 0
    HighDaily = 0
    OpenDaily = 0

    for i in range(len):
        open = objStockChart.GetDataValue(1, len - i - 1)
        high = objStockChart.GetDataValue(2, len - i - 1)
        low = objStockChart.GetDataValue(3, len - i - 1)
        close = objStockChart.GetDataValue(4, len - i - 1)
        vol = objStockChart.GetDataValue(5, len - i - 1)
        CloseDaily = close
        LowDaily = low
        HighDaily = high
        OpenDaily = open
        objSeries.Add(close, open, high, low, vol)

    objIndex.series = objSeries

    retDaily = [OpenDaily, CloseDaily, LowDaily, HighDaily]
    # 지표들구하기
    indexNames = ['이동평균(라인1개)', '일목균형표', 'ATR']
    for indexName in indexNames:
        objIndex.put_IndexKind(indexName)
        objIndex.put_IndexDefault(indexName)
        # 지표 데이터 계산 하기
        objIndex.Calculate()
        cntofIndex = objIndex.GetCount(0)
        for i in range(cntofIndex - 2, cntofIndex):
            retDaily.append(objIndex.GetResult(0, i))

    return retDaily


def GetSearchingResult():
    tmp = Find_Target.GetTarget()
    # print(tmp)
    ret = []
    for item in tmp:
        ret.append(item['code'])
    return ret


def GetFinalResult():
    codes = GetSearchingResult()
    ret1 = []
    ret2 = []
    ret3 = []
    print('1차 대상 ', len(codes), '개')
    print(codes)
    for code in codes:
        # if code not in kosdaq:
        #     continue
        dailyIndex = GetData_Daily(code)
        weeklyIndex = GetData_Weekly(code)
        # 시, 종, 저, 고, 전일5, 당일5, 전일전환, 당일전환, 전일atr, 당일atr
        # 0   1   2  3   4      5       6       7       8       9
        if ####조건부분#####:
            if ####조건부분#####:
                if ####조건부분#####:
                    # print(code)
                    # print(dailyIndex)
                    # print(weeklyIndex)
                    ret1.append(code)
                    ret2.append(dailyIndex)
                    ret3.append(weeklyIndex)
       
    print('2차 대상 ', len(ret1), '개')
    dbgout(str(ret1))
    print(ret1)
    return ret1, ret2, ret3


def UpdateCsv():
    codes, dailyIndex, weeklyIndex = GetFinalResult()
    # csv_origin = pd.read_csv('tmp.csv', index_col='종목코드', encoding='cp949')
    csv_origin = pd.read_csv('tmp.csv', encoding='cp949')

    if len(csv_origin.index) >= 10:
        return False
    elif len(csv_origin.index) + len(codes) > 10:
        tc = 10 - len(csv_origin.index)
        codes = codes[:tc]
        dailyIndex = dailyIndex[:tc]
        weeklyIndex = weeklyIndex[:tc]

    with open('tmp.csv', 'a', newline='') as f_object:
        writer_object = writer(f_object)
        # Result - a writer object
        # Pass the data in the list as an argument into the writerow() function
        for ind in range(len(codes)):
            i = codes[ind]
            if i in list(csv_origin.loc[:, '종목코드']):
                continue
            # 출력
            # 시, 종, 저, 고, 전일5, 당일5, 전일전환, 당일전환, 전일atr, 당일atr
            # 0   1   2  3   4      5       6       7       8       9
            # csv
            # 종목코드	종가	시가	5일이평	일전환선	5주이평	주전환선	당일ATR 손절가
            writer_object.writerow(
                [i, dailyIndex[ind][1], 'X', dailyIndex[ind][5], dailyIndex[ind][7],
                 weeklyIndex[ind][5], weeklyIndex[ind][7], dailyIndex[ind][9],
                 min(dailyIndex[ind][5], dailyIndex[ind][7], weeklyIndex[ind][5],
                     weeklyIndex[ind][7]) - dailyIndex[ind][9]])
        # Close the file object
        f_object.close()


def GetOrderTarget_1530():
    csv_cur = pd.read_csv('tmp.csv', index_col='종목코드', encoding='cp949')
    # cur_csv = pd.read_csv('tmp.csv', encoding='cp949')
    ret = []
    for index in csv_cur.index:
        # print(csv_cur.loc[index, :].unique())
        if 'O' in csv_cur.loc[index, :].unique():
            continue
        # print(index)

        ret.append([index, csv_cur.loc[index, '종가']])

    return ret


def OrderBuy1530(target):
    # 장후단일가 주문 넣기
    # 1. 장후단일가 주문 대상 불러오기
    targetCodes = GetOrderTarget_1530()
    dbgout('15:30분 함수 실행')
    # 2. 주문 넣기
    objOrderAft1530.SetInputValue(0, "2")
    objOrderAft1530.SetInputValue(1, acc)
    objOrderAft1530.SetInputValue(2, accFlag)
    for i in range(len(targetCodes)):
        targetCode = targetCodes[i][0]
        targetPrice = targetCodes[i][1]
        if targetPrice > 200000:
            continue
        orderCnt = moneyPerQty // targetPrice
        objOrderAft1530.SetInputValue(3, targetCode)
        objOrderAft1530.SetInputValue(4, orderCnt)
        time.sleep(5)
        dbgout('order' + targetCode + ' ' + orderCnt + 'ea')
        nRet = objOrderAft1530.BlockRequest()

        if nRet != 0:
            print("주문요청 오류", nRet)
            # 0: 정상,  그 외 오류, 4: 주문요청제한 개수 초과
            exit()

        rqStatus = objOrderAft1530.GetDibStatus()
        errMsg = objOrderAft1530.GetDibMsg1()
        if rqStatus != 0:
            print("주문 실패: ", rqStatus, errMsg)
            exit()


def CheckConclusion1530():
    # 시간외종가 체결 대응
    # objCheckNotConcluded
    # 0. 주문 했던 종목들 수신
    orderedCodes = GetOrderTarget_1530()
    notConcludedCodes = []

    # 1. 정보 입력
    objCheckNotConcluded.SetInputValue(0, acc)
    objCheckNotConcluded.SetInputValue(1, accFlag)
    objCheckNotConcluded.SetInputValue(6, "2")
    objCheckNotConcluded.SetInputValue(7, 20)

    # 2. 정보 요청
    # 16:00에 request해 체결 되었는지 확인 -> 그냥 함수 호출을 16:00에 하기
    ret = objCheckNotConcluded.BlockRequest()

    # 통신 초과 요청 방지에 의한 요류 인 경우
    while (ret == 4):  # 연속 주문 오류 임. 이 경우는 남은 시간동안 반드시 대기해야 함.
        remainTime = objCpCybos.LimitRequestRemainTime
        print("연속 통신 초과에 의해 재 통신처리 : ", remainTime / 1000, "초 대기")
        time.sleep(remainTime / 1000)
        ret = objCheckNotConcluded.BlockRequest()

    # 3. 체결 확인
    receivedCnt = objCheckNotConcluded.GetHeaderValue(5)

    for i in range(receivedCnt):
        # orderNum = objCheckNotConcluded.GetDataValue(1, i) # 주문번호
        # orderPrev = objCheckNotConcluded.GetDataValue(2, i) # 원주문번호
        code = objCheckNotConcluded.GetDataValue(3, i)  # 종목코드
        name = objCheckNotConcluded.GetDataValue(4, i)  # 종목명
        # orderDesc = objCheckNotConcluded.GetDataValue(5, i)  # 주문구분내용
        amount = objCheckNotConcluded.GetDataValue(6, i)  # 주문수량
        # price = objCheckNotConcluded.GetDataValue(7, i)  # 주문단가
        ContAmount = objCheckNotConcluded.GetDataValue(8, i)  # 체결수량
        # credit = objCheckNotConcluded.GetDataValue(9, i)  # 신용구분
        # modAvali = objCheckNotConcluded.GetDataValue(11, i)  # 정정취소 가능수량
        # buysell = objCheckNotConcluded.GetDataValue(13, i)  # 매매구분코드
        # creditdate = objCheckNotConcluded.GetDataValue(17, i)  # 대출일
        # orderFlagDesc = objCheckNotConcluded.GetDataValue(19, i)  # 주문호가구분코드내용
        # orderFlag = objCheckNotConcluded.GetDataValue(21, i)  # 주문호가구분코드

        if code in orderedCodes:
            # dbgout('시간외종가 미체결 종목 및 체결 수량(일부 체결 케이스)')
            # dbgout( code + '(' + name + ') ' + 'order qty : ' + amount + 'ea')
            # dbgout( code + '(' + name + ') ' + 'conclusion qty : ' + ContAmount + 'ea')
            notConcludedCodes.append(code)

    # 4. 체결 체크하기
    with open('tmp.csv', 'r', newline='') as f_object:
        writer_object = writer(f_object)
        rdr = reader(f_object)
        lines = []
        for line in rdr:
            if line[0] in orderedCodes:
                if line[0] not in notConcludedCodes:
                    line[1] = 'O'
            lines.append(line)
        f_object.close()

    with open('tmp.csv', 'w', newline='') as f_object:
        writer_object = writer(f_object)
        writer_object.writerows(lines)
        f_object.close()


def exe1530():
    dbgout('15:30분 함수 실행')
    UpdateCsv()
    time.sleep(5)
    OrderBuy1530()


def exe1600():
    dbgout('장후단일가 체결 여부 체크 및 csv 업데이트')
    CheckConclusion1530()


def GetOrderPrice():
    # csv 읽어오기
    csv_cur = pd.read_csv('tmp.csv', index_col='종목코드', encoding='cp949')

    # return 리스트
    targetCode = []
    targetBuySell = []
    targetPrice = []
    targetQty = []
    targetDcd = []

    # 'O' 개수에 따른 주문량
    OrderCount = [2, 4, 8, 16]

    # 컬럼명
    columns = ['5일이평', '5주이평', '일전환선', '주전환선']

    # 'O' 개수로 구분하기
    for index in csv_cur.index:
        # 하나도 없으면
        # 종가 매수 구분 필요? -> 전일 종가에 주문 넣기 위해 : 15:30~18:00(시간외단일가 포함)
        if 'O' not in csv_cur.loc[index, :].unique():
            targetCode.append(index)
            targetPrice.append(csv_cur.loc[index, '종가'])
            targetQty.append(1)
            targetDcd.append('종가')
            targetBuySell.append('매수')
        # 하나라도 있으면
        else:
            # 매수부 구현
            countO = 0
            prices = []

            for column in columns:
                cur = csv_cur.loc[index, column]
                if cur == 'O':
                    countO += 1
                else:
                    prices.append((-cur, column))

            prices.sort()

            # O의 개수가 0개면, 가장 낮은 가격에 Qty=2, 그다음에4, 그다음에 8, 그다음에 16
            # 1개면 가장낮은 가격에 4부터 하나씩
            # 즉, 0개수따라 제일 낮은 가격부터 OrderCount[Ocount] and Ocount += 1로 늘려가기
            for i in range(countO, 4):
                ind = i - countO
                targetCode.append(index)
                target_price_tmp = -prices[ind][0]
                target_price = 0
                # 호가 단위 맞추기
                if target_price_tmp < 1000:
                    target_price = round(target_price_tmp)
                elif target_price_tmp < 5000:
                    target_price = round(target_price_tmp / 5) * 5
                elif target_price_tmp < 10000:
                    target_price = round(target_price_tmp, -1)
                elif target_price_tmp < 50000:
                    target_price = round(target_price_tmp / 50) * 50
                else:
                    target_price = round(target_price_tmp, -2)
                targetPrice.append(target_price)
                # 1 당 얼마를 투자할지는 추후 매수 주문 때?
                targetQty.append(OrderCount[i])
                targetDcd.append(prices[ind][1])
                targetBuySell.append('매수')

            # 매도부 구현 (익절가로 주문 넣기)
            # O 개수에 따라 익절가 바뀐다
            # ATR * (5 - Ocount/6) --> 즉, 물 4번 다 탔으면 수익률 ATR/6에 매도 / 한번도 물 안탔으면 5/6 ATR
            # 우선, 현재 계좌에 대한 정보를 가져오도록 cp0322에 정보 입력
            # objAccStat
            # objCpTradeUtil.TradeInit()
            # acc = objCpTradeUtil.AccountNumber[0]  # 계좌번호
            # accFlag = objCpTradeUtil.GoodsList(acc, 1)  # -1:전체,1:주식,2:선물/옵션
            objAccStat.SetInputValue(0, acc)  # 계좌번호
            objAccStat.SetInputValue(1, accFlag[0])  # (string) 상품관리구분코드
            # objAccStat.SetInputValue(2, 14)  # (long) 요청건수[default:14] - 최대 50개
            objAccStat.SetInputValue(3, "1")  # (string) 수익률구분코드 - ( "1" : 100% 기준, "2": 0% 기준)

            objAccStat.BlockRequest()

            codeCnt = objAccStat.GetHeaderValue(7)

            for i in range(codeCnt):
                code = objAccStat.GetDataValue(12, i)
                time.sleep(1)
                balance = objAccStat.GetDataValue(18, i)
                time.sleep(1)
                qty = objAccStat.GetDataValue(15, i)
                time.sleep(1)
                atr = csv_cur.loc[code, '당일ATR']
                # 호가단위
                # 1000원 미만 : 1
                # ~ 5000원 미만 : 5
                # ~ 10,000원 미만 : 10
                # ~ 50,000원 미만 : 50
                # 나머지 : 100
                target_price_tmp = balance + atr
                target_price = 0
                if target_price_tmp < 1000:
                    target_price = round(target_price_tmp)
                elif target_price_tmp < 5000:
                    target_price = round(target_price_tmp / 5) * 5
                elif target_price_tmp < 10000:
                    target_price = round(target_price_tmp, -1)
                elif target_price_tmp < 50000:
                    target_price = round(target_price_tmp / 50) * 50
                else:
                    target_price = round(target_price_tmp, -2)
                targetCode.append(code)
                targetPrice.append(target_price)
                targetQty.append(qty)
                targetDcd.append('')
                targetBuySell.append('매도')

    return targetCode, targetPrice, targetQty, targetDcd, targetBuySell


# UpdateCsv()
# print(GetOrderTarget_1530())
# GetData_Daily('A007770')
# print(GetFinalResult())
# print(GetOrderPrice())
# (['A007770', 'A340440', 'A206650', 'A101670', 'A236200', 'A343510', 'A032960', 'A086520'], [21700, 4040, 17250, 1990, 21700, 2205, 14400, 88500], [1, 1, 1, 1, 1, 1, 1, 1], ['종가', '종가', '종가', '종가', '종가', '종가', '종가', '종가'], ['매수', '매수', '매수', '매수', '매수', '매수', '매수', '매수'])

UpdateCsv()


# csv 업데이트법
# with open('tmp.csv', 'r', newline='') as f_object:
#     writer_object = writer(f_object)
#     rdr = reader(f_object)
#     lines = []
#     for line in rdr:
#         if line[0] == 'A049410':
#             line[1] = 'O'
#         lines.append(line)
#     f_object.close()
#
# with open('tmp.csv', 'w', newline='') as f_object:
#     writer_object = writer(f_object)
#     writer_object.writerows(lines)
#     f_object.close()

from pykiwoom.kiwoom import *
import datetime
import time
import pandas as pd
import os

# 로그인
kiwoom = Kiwoom()
kiwoom.CommConnect()

# 전종목 종목코드
kospi = kiwoom.GetCodeListByMarket('0')
kosdaq = kiwoom.GetCodeListByMarket('10')
codes = kospi + kosdaq

# 문자열로 오늘 날짜 얻기
now = datetime.datetime.now()
today = now.strftime("%Y%m%d")

# 전 종목의 일봉 데이터
for i, code in enumerate(codes):
    print(f"{i}/{len(codes)} {code}")
    df = kiwoom.block_request("opt10081",
                              종목코드=code,
                              기준일자=today,
                              수정주가구분=1,
                              output="주식일봉차트조회",
                              next=0)

    out_name = f"{code}.xlsx"
    df.to_excel(out_name)
    time.sleep(3.6)

flist = os.listdir()
xlsx_list = [x for x in flist if x.endswith('.xlsx')]
close_data = []

for xls in xlsx_list:
    code = xls.split('.')[0]
    df = pd.read_excel(xls)
    # 예외가 발생했습니다. BadZipFile

    df2 = df[['일자', '현재가']].copy()
    df2.rename(columns={'현재가': code}, inplace=True)
    df2 = df2.set_index('일자')
    df2 = df2[::-1]
    close_data.append(df2)

# concat
df = pd.concat(close_data, axis=1)
df.to_excel("merge.xlsx")

df = pd.read_excel("momentum_list.xlsx")
df.columns = ["종목코드", "모멘텀", "순위"]

# 종목명 추가하기
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)
codes = df["종목코드"]
names = [kiwoom.GetMasterCodeName(code) for code in codes]
df['종목명'] = pd.Series(data=names)


# 매수하기
accounts = kiwoom.GetLoginInfo('ACCNO')
account = accounts[0]

for code in codes:
    ret = kiwoom.SendOrder("시장가매수", "0101", account, 1, code, 100, 0, "03", "")
    time.sleep(0.2)
    print(code, "종목 시장가 주문 완료")
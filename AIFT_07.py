from pykiwoom.kiwoom import *

#로그인
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)
print("블록킹 로그인 완료")

#사용자 정보 얻어오기
account_num = kiwoom.GetLoginInfo("ACCOUNT_CNT")        # 전체 계좌수
accounts = kiwoom.GetLoginInfo("ACCNO")                 # 전체 계좌 리스트
user_id = kiwoom.GetLoginInfo("USER_ID")                # 사용자 ID
user_name = kiwoom.GetLoginInfo("USER_NAME")            # 사용자명
keyboard = kiwoom.GetLoginInfo("KEY_BSECGB")            # 키보드보안 해지여부
firewall = kiwoom.GetLoginInfo("FIREW_SECGB")           # 방화벽 설정 여부

print(account_num)
print(accounts)
print(user_id)
print(user_name)
print(keyboard)
print(firewall)

#종목 코드 얻기
kospi = kiwoom.GetCodeListByMarket('0')
kosdaq = kiwoom.GetCodeListByMarket('10')
etf = kiwoom.GetCodeListByMarket('8')

print(len(kospi), kospi)
print(len(kosdaq), kosdaq)
print(len(etf), etf)

#종목명 얻기
name = kiwoom.GetMasterCodeName("005930")
print(name)

#연결 상태 확인
state = kiwoom.GetConnectState()
if state == 0:
    print("미연결")
elif state == 1:
    print("연결완료")

#상장 주식수 얻기
stock_cnt = kiwoom.GetMasterListedStockCnt("005930")
print("삼성전자 상장주식수: ", stock_cnt)

감리구분 = kiwoom.GetMasterConstruction("005930")
print(감리구분)

상장일 = kiwoom.GetMasterListedStockDate("005930")
print(상장일)
print(type(상장일))

전일가 = kiwoom.GetMasterLastPrice("005930")
print(int(전일가))
print(type(전일가))

종목상태 = kiwoom.GetMasterStockState("005930")
print(종목상태)

#테마그룹
import pprint

group = kiwoom.GetThemeGroupList(1)
pprint.pprint(group)

#테마별 종목코드
tickers = kiwoom.GetThemeGroupCode('330')
for ticker in tickers:
    name = kiwoom.GetMasterCodeName(ticker)
    print(ticker, name)

#매수
주식계좌
accounts = kiwoom.GetLoginInfo("ACCNO")
stock_account = accounts[0]

# 삼성전자, 10주, 시장가주문 매수
kiwoom.SendOrder("시장가매수", "0101", stock_account, 1, "005930", 10, 0, "03", "")

#매도
주식계좌
accounts = kiwoom.GetLoginInfo("ACCNO")
stock_account = accounts[0]

# 삼성전자, 10주, 시장가주문 매도
kiwoom.SendOrder("시장가매도", "0101", stock_account, 2, "005930", 10, 0, "03", "")

#키움 로그인창 실행
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self.slot_login)
        self.ocx.dynamicCall("CommConnect()")

    def slot_login(self, err_code):
        print(err_code)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_() 

#싱글데이터 TR 사용하기
df = kiwoom.block_request("opt10001",
                          종목코드="005930",
                          output="주식기본정보",
                          next=0)
print(df)

#멀티데이터 TR
import time
import pandas as pd

df = kiwoom.block_request("opt10081",
                          종목코드="005930",
                          기준일자="20200424",
                          수정주가구분=1,
                          output="주식일봉차트조회",
                          next=0)
print(df.head())

#TR 요청 (연속조회)
dfs = []
df = kiwoom.block_request("opt10081",
                          종목코드="005930",
                          기준일자="20200424",
                          수정주가구분=1,
                          output="주식일봉차트조회",
                          next=0)
print(df.head())
dfs.append(df)

while kiwoom.tr_remained:
    df = kiwoom.block_request("opt10081",
                              종목코드="005930",
                              기준일자="20200424",
                              수정주가구분=1,
                              output="주식일봉차트조회",
                              next=2)
    dfs.append(df)
    time.sleep(1)

df = pd.concat(dfs)
df.to_excel("005930.xlsx")

#전종목 일봉 데이터 다운로드
import datetime
import time

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

#전 종목 일봉 데이터 머지(merge)
import pandas as pd
import os

flist = os.listdir()
xlsx_list = [x for x in flist if x.endswith('.xlsx')]
close_data = []

for xls in xlsx_list:
    code = xls.split('.')[0]
    df = pd.read_excel(xls)
    df2 = df[['일자', '현재가']].copy()
    df2.rename(columns={'현재가': code}, inplace=True)
    df2 = df2.set_index('일자')
    df2 = df2[::-1]
    close_data.append(df2)

# concat
df = pd.concat(close_data, axis=1)
df.to_excel("merge.xlsx")

#모멘텀 전략 종목 선정
import pandas as pd

df = pd.read_excel("merge.xlsx")
df['일자'] = pd.to_datetime(df['일자'], format="%Y%m%d")
df = df.set_index('일자')

# 60 영업일 수익률
return_df = df.pct_change(60)

return_df.tail()

s = return_df.loc["2020-06-22"]
momentum_df = pd.DataFrame(s)
momentum_df.columns = ["모멘텀"]

momentum_df.head(n=10)

momentum_df['순위'] = momentum_df['모멘텀'].rank(ascending=False)
momentum_df.head(n=10)

momentum_df = momentum_df.sort_values(by='순위')
momentum_df[:30]

momentum_df[:30].to_excel("momentum_list.xlsx")

#선종 종목 매수하기
import pandas as pd
from pykiwoom.kiwoom import *
import time

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

#장시작시간 테스트 프로그램
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import datetime


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real")
        self.setGeometry(300, 300, 300, 400)

        btn = QPushButton("Register", self)
        btn.move(20, 20)
        btn.clicked.connect(self.btn_clicked)

        btn2 = QPushButton("DisConnect", self)
        btn2.move(20, 100)
        btn2.clicked.connect(self.btn2_clicked)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommmConnect()

    def btn_clicked(self):
        #self.SetRealReg("1000", "005930", "20;10", 0)
        self.SetRealReg("2000", "", "215;20;214", 0)
        print("called\n")

    def btn2_clicked(self):
        self.DisConnectRealData("1000")

    def CommmConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.statusBar().showMessage("login 중 ...")

    def _handler_login(self, err_code):
        if err_code == 0:
            self.statusBar().showMessage("login 완료")


    def _handler_real_data(self, code, real_type, data):
        print(code, real_type, data)
        if real_type == "장시작시간":
            gubun =  self.GetCommRealData(code, 215)
            remained_time =  self.GetCommRealData(code, 214)
            print(gubun, remained_time)            


    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)

    def DisConnectRealData(self, screen_no):
        self.ocx.dynamicCall("DisConnectRealData(QString)", screen_no)

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#주식체결
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import datetime


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real")
        self.setGeometry(300, 300, 300, 400)

        btn = QPushButton("Register", self)
        btn.move(20, 20)
        btn.clicked.connect(self.btn_clicked)

        btn2 = QPushButton("DisConnect", self)
        btn2.move(20, 100)
        btn2.clicked.connect(self.btn2_clicked)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommmConnect()

    def btn_clicked(self):
        self.SetRealReg("1000", "005930", "20;10", 0)

    def btn2_clicked(self):
        self.DisConnectRealData("1000")

    def CommmConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.statusBar().showMessage("login 중 ...")

    def _handler_login(self, err_code):
        if err_code == 0:
            self.statusBar().showMessage("login 완료")


    def _handler_real_data(self, code, real_type, data):
        if real_type == "주식체결":
            # 체결 시간 
            time =  self.GetCommRealData(code, 20)
            date = datetime.datetime.now().strftime("%Y-%m-%d ")
            time =  datetime.datetime.strptime(date + time, "%Y-%m-%d %H%M%S")
            print(time, end=" ")

            # 현재가 
            price =  self.GetCommRealData(code, 10)
            print(int(price))


    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)

    def DisConnectRealData(self, screen_no):
        self.ocx.dynamicCall("DisConnectRealData(QString)", screen_no)

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#주식우선호가
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import datetime


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real")
        self.setGeometry(300, 300, 300, 400)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)

        btn = QPushButton("구독", self)
        btn.move(20, 20)
        btn.clicked.connect(self.btn_clicked)

        btn2 = QPushButton("해지", self)
        btn2.move(180, 20)
        btn2.clicked.connect(self.btn2_clicked)

        # 2초 후에 로그인 진행
        QTimer.singleShot(1000 * 2, self.CommmConnect)


    def btn_clicked(self):
        self.SetRealReg("1000", "005930", "27;28", 0)

    def btn2_clicked(self):
        self.DisConnectRealData("1000") 

    def CommmConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.statusBar().showMessage("login 중 ...")

    def _handler_login(self, err_code):
        if err_code == 0:
            self.statusBar().showMessage("login 완료")


    def _handler_real_data(self, code, real_type, data):
        if real_type == "주식우선호가":
            now = datetime.datetime.now()
            ask01 =  self.GetCommRealData(code, 27)         
            bid01 =  self.GetCommRealData(code, 28)         

            print(f"현재시간 {now} | 최우선매도호가: {ask01} 최우선매수호가: {bid01}")

    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)
        self.statusBar().showMessage("구독 신청 완료")

    def DisConnectRealData(self, screen_no):
        self.ocx.dynamicCall("DisConnectRealData(QString)", screen_no)
        self.statusBar().showMessage("구독 해지 완료")

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#주식호가잔량
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import datetime


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("주식호가잔량")
        self.setGeometry(300, 300, 300, 400)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)

        btn = QPushButton("구독", self)
        btn.move(20, 20)
        btn.clicked.connect(self.btn_clicked)

        btn2 = QPushButton("해지", self)
        btn2.move(180, 20)
        btn2.clicked.connect(self.btn2_clicked)

        # 2초 후에 로그인 진행
        QTimer.singleShot(1000 * 2, self.CommmConnect)


    def btn_clicked(self):
        self.SetRealReg("1000", "005930", "41", 0)

    def btn2_clicked(self):
        self.DisConnectRealData("1000") 

    def CommmConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.statusBar().showMessage("login 중 ...")

    def _handler_login(self, err_code):
        if err_code == 0:
            self.statusBar().showMessage("login 완료")

    def _handler_real_data(self, code, real_type, data):
        if real_type == "주식호가잔량":
            hoga_time =  self.GetCommRealData(code, 21)         
            ask01_price =  self.GetCommRealData(code, 41)         
            ask01_volume =  self.GetCommRealData(code, 61)         
            bid01_price =  self.GetCommRealData(code, 51)         
            bid01_volume =  self.GetCommRealData(code, 71)         
            print(hoga_time)
            print(f"매도호가: {ask01_price} - {ask01_volume}")
            print(f"매수호가: {bid01_price} - {bid01_volume}")

    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)
        self.statusBar().showMessage("구독 신청 완료")

    def DisConnectRealData(self, screen_no):
        self.ocx.dynamicCall("DisConnectRealData(QString)", screen_no)
        self.statusBar().showMessage("구독 해지 완료")

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#VI발동/해제
import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 200)
        self.setWindowTitle("Kiwoom VI 발동 테스트")

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveTrData.connect(self._handler_tr)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommConnect()

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def request_vi(self):
        self.SetInputValue("시장구분", "000")
        self.SetInputValue("장전구분", "1")
        self.SetInputValue("종목코드", "")
        self.SetInputValue("발동구분", "1")
        self.SetInputValue("제외종목", "111111011")
        self.SetInputValue("거래량구분", "0")
        #self.SetInputValue("최소거래량", "0")
        #self.SetInputValue("최대거래량", "0")
        self.SetInputValue("거래대금구분", "0")
        #self.SetInputValue("최소거래대금", "0")
        #self.SetInputValue("최대거래대금", "0")
        self.SetInputValue("발동방향", "0")
        self.CommRqData("opt10054", "opt10054", 0, "1000")

    def _handler_login(self, err_code):
        print("handler login", err_code);
        self.request_vi()

    def _handler_tr(self, screen, rqname, trcode, record, next):
        print("OnReceiveTrData: ", screen, rqname, trcode, record, next)

    def _handler_real_data(self, code, real_type, real_data):
        print("OnReceiveRealData", code, real_type, real_data)

    def SetInputValue(self, id, value):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    def CommRqData(self, rqname, trcode, next, screen):
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen)

    def GetCommData(self, trcode, rqname, index, item):
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, index, item)
        return data.strip()

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#업종지수
import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 200)
        self.setWindowTitle("Kiwoom Sector Index 테스트")

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveTrData.connect(self._handler_tr)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommConnect()

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def request_sector_index(self):
        self.SetInputValue("시장구분", "0")
        self.SetInputValue("업종코드", "001")

        self.CommRqData("opt20001", "opt20001", 0, "1000")

    def _handler_login(self, err_code):
        print("handler login", err_code);
        self.request_sector_index()

    def _handler_tr(self, screen, rqname, trcode, record, next):
        print("OnReceiveTrData: ", screen, rqname, trcode, record, next)

    def _handler_real_data(self, code, real_type, real_data):
        if real_type == "업종지수":
            code = self.GetCommRealData(code, 20) 
            open = self.GetCommRealData(code, 16) 
            high = self.GetCommRealData(code, 17) 
            low  = self.GetCommRealData(code, 18) 
            close= self.GetCommRealData(code, 10)
            print(code, open, high, low, close)

    def SetInputValue(self, id, value):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    def CommRqData(self, rqname, trcode, next, screen):
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen)

    def GetCommData(self, trcode, rqname, index, item):
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, index, item)
        return data.strip()

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#업종등락
import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 200)
        self.setWindowTitle("Kiwoom Sector Fluctuation 테스트")

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveTrData.connect(self._handler_tr)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommConnect()

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def request_sector_index(self):
        self.SetInputValue("시장구분", "0")
        self.SetInputValue("업종코드", "001")

        self.CommRqData("opt20001", "opt20001", 0, "1000")

    def _handler_login(self, err_code):
        print("handler login", err_code);
        self.request_sector_index()

    def _handler_tr(self, screen, rqname, trcode, record, next):
        print("OnReceiveTrData: ", screen, rqname, trcode, record, next)

    def _handler_real_data(self, code, real_type, real_data):
        if real_type == "업종지수":
            code = self.GetCommRealData(code, 20) 
            open = self.GetCommRealData(code, 16) 
            high = self.GetCommRealData(code, 17) 
            low  = self.GetCommRealData(code, 18) 
            close= self.GetCommRealData(code, 10)
            print('SectorIndex', code, open, high, low, close)
        elif real_type == "업종등락":
            time = self.GetCommRealData(code, 20) 
            up_cnt = self.GetCommRealData(code, 252) 
            down_cnt = self.GetCommRealData(code, 255) 
            print("SectorFluctuation", time, up_cnt, down_cnt)

    def SetInputValue(self, id, value):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    def CommRqData(self, rqname, trcode, next, screen):
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen)

    def GetCommData(self, trcode, rqname, index, item):
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, index, item)
        return data.strip()

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#실시간 조건식
import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 200)
        self.setWindowTitle("Kiwoom 실시간 조건식 테스트")

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveConditionVer.connect(self._handler_condition_load)
        self.ocx.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.CommConnect()

        btn1 = QPushButton("condition down")
        btn2 = QPushButton("condition list")
        btn3 = QPushButton("condition send")

        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(btn1)
        layout.addWidget(btn2)
        layout.addWidget(btn3)
        self.setCentralWidget(widget)

        # event
        btn1.clicked.connect(self.GetConditionLoad)
        btn2.clicked.connect(self.GetConditionNameList)
        btn3.clicked.connect(self.send_condition)

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def _handler_login(self, err_code):
        print("handler login", err_code);

    def _handler_condition_load(self, ret, msg):
        print("handler condition load", ret, msg)

    def _handler_real_condition(self, code, type, cond_name, cond_index):
        print(cond_name, code, type) 

    def GetConditionLoad(self):
        self.ocx.dynamicCall("GetConditionLoad()")

    def GetConditionNameList(self):
        data = self.ocx.dynamicCall("GetConditionNameList()")
        conditions = data.split(";")[:-1]
        for condition in conditions:
            index, name = condition.split('^')
            print(index, name)

    def SendCondition(self, screen, cond_name, cond_index, search):
        ret = self.ocx.dynamicCall("SendCondition(QString, QString, int, int)", screen, cond_name, cond_index, search)

    def SendConditionStop(self, screen, cond_name, cond_index):
        ret = self.ocx.dynamicCall("SendConditionStop(QString, QString, int)", screen, cond_name, cond_index)

    def send_condition(self):
        self.SendCondition("100", "test", "000", 1)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#코스닥 KODEX ETF150 변동성 돌파
import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import datetime


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 400, 300)
        self.setWindowTitle("코스닥 ETF150 변동성 돌파")

        self.range = None
        self.target = None
        self.account = None
        self.amount = None
        self.hold = None

        self.previous_day_hold = False
        self.previous_day_quantity = False

        self.plain_text_edit = QPlainTextEdit(self)
        self.plain_text_edit.setReadOnly(True)
        self.plain_text_edit.move(10, 10)
        self.plain_text_edit.resize(380, 280)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveTrData.connect(self._handler_tr_data)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.ocx.OnReceiveChejanData.connect(self._handler_chejan_data)

        self.login_event_loop = QEventLoop()
        self.CommConnect()          # 로그인이 될 때까지 대기
        self.run()

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.login_event_loop.exec()

    def run(self):
        accounts = self.GetLoginInfo("ACCNO")
        self.account = accounts.split(';')[0]
        print(self.account)

        # TR 요청 
        self.request_opt10081()
        self.request_opw00001()
        self.request_opw00004()

        # 주식체결 (실시간)
        self.subscribe_market_time('1')
        self.subscribe_stock_conclusion('2')

    def GetLoginInfo(self, tag):
        data = self.ocx.dynamicCall("GetLoginInfo(QString)", tag)
        return data

    def _handler_login(self, err_code):
        if err_code == 0:
            self.plain_text_edit.appendPlainText("로그인 완료")
        self.login_event_loop.exit()

    def _handler_tr_data(self, screen_no, rqname, trcode, record, next):
        if rqname == "KODEX일봉데이터":
            now = datetime.datetime.now()
            today = now.strftime("%Y%m%d")
            일자 = self.GetCommData(trcode, rqname, 0, "일자")

            # 장시작 후 TR 요청하는 경우 0번째 row에 당일 일봉 데이터가 존재함
            if 일자 != today:
                고가 = self.GetCommData(trcode, rqname, 0, "고가")
                저가 = self.GetCommData(trcode, rqname, 0, "저가")
            else:
                일자 = self.GetCommData(trcode, rqname, 1, "일자")
                고가 = self.GetCommData(trcode, rqname, 1, "고가")
                저가 = self.GetCommData(trcode, rqname, 1, "저가")

            self.range = int(고가) - int(저가)
            info = f"일자: {일자} 고가: {고가} 저가: {저가} 전일변동: {self.range}"
            self.plain_text_edit.appendPlainText(info)

        elif rqname == "예수금조회":
            주문가능금액 = self.GetCommData(trcode, rqname, 0, "주문가능금액")
            주문가능금액 = int(주문가능금액)
            self.amount = int(주문가능금액 * 0.2)
            self.plain_text_edit.appendPlainText(f"주문가능금액: {주문가능금액} 투자금액: {self.amount}")

        elif rqname == "계좌평가현황":
            rows = self.GetRepeatCnt(trcode, rqname)
            for i in range(rows):
                종목코드 = self.GetCommData(trcode, rqname, i, "종목코드")
                보유수량 = self.GetCommData(trcode, rqname, i, "보유수량")

                if 종목코드[1:] == "229200":
                    self.previous_day_hold = True
                    self.previous_day_quantity = int(보유수량)

    def GetRepeatCnt(self, trcode, rqname):
        ret = self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def request_opt10081(self):
        now = datetime.datetime.now()
        today = now.strftime("%Y%m%d")
        self.SetInputValue("종목코드", "229200")
        self.SetInputValue("기준일자", today)
        self.SetInputValue("수정주가구분", 1)
        self.CommRqData("KODEX일봉데이터", "opt10081", 0, "9000")

    def request_opw00001(self):
        self.SetInputValue("계좌번호", self.account)
        self.SetInputValue("비밀번호", "")
        self.SetInputValue("비밀번호입력매체구분", "00")
        self.SetInputValue("조회구분", 2)
        self.CommRqData("예수금조회", "opw00001", 0, "9001")

    def request_opw00004(self):
        self.SetInputValue("계좌번호", self.account)
        self.SetInputValue("비밀번호", "")
        self.SetInputValue("상장폐지조회구분", 0)
        self.SetInputValue("비밀번호입력매체구분", "00")
        self.CommRqData("계좌평가현황", "opw00004", 0, "9002")

    # 실시간 타입을 위한 메소드
    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data

    def DisConnectRealData(self, screen_no):
        self.ocx.dynamicCall("DisConnectRealData(QString)", screen_no)

    # 실시간 이벤트 처리 핸들러 
    def _handler_real_data(self, code, real_type, real_data):
        if real_type == "장시작시간":
            장운영구분 = self.GetCommRealData(code, 215)
            if 장운영구분 == '3':
                if self.previous_day_hold:
                    self.previous_day_hold = False
                    # 매도 (시장가)
                    self.SendOrder("매도", "8001", self.account, 2, "229200", self.previous_day_quantity, 0, "03", "")
            elif 장운영구분 == '4':
                QCoreApplication.instance().quit()
                print("메인 윈도우 종료")

        elif real_type == "주식체결": 
            # 현재가 
            현재가 = self.GetCommRealData(code, 10)
            현재가 = abs(int(현재가))          # +100, -100
            체결시간 = self.GetCommRealData(code, 20)

            # 목표가 계산
            # TR 요청을 통한 전일 range가 계산되었고 아직 당일 목표가가 계산되지 않았다면 
            if self.range is not None and self.target is None:
                시가 = self.GetCommRealData(code, 16)
                시가 = abs(int(시가))          # +100, -100
                self.target = int(시가 + (self.range * 0.5))      
                self.plain_text_edit.appendPlainText(f"목표가 계산됨: {self.target}")

            # 매수시도
            # 당일 매수하지 않았고
            # TR 요청과 Real을 통한 목표가가 설정되었고 
            # TR 요청을 통해 잔고조회가 되었고 
            # 현재가가 목표가가 이상이면
            if self.hold is None and self.target is not None and self.amount is not None and 현재가 > self.target:
                self.hold = True 
                quantity = int(self.amount / 현재가)
                self.SendOrder("매수", "8000", self.account, 1, "229200", quantity, 0, "03", "")
                self.plain_text_edit.appendPlainText(f"시장가 매수 진행 수량: {quantity}")

            # 로깅
            self.plain_text_edit.appendPlainText(f"시간: {체결시간} 목표가: {self.target} 현재가: {현재가} 보유여부: {self.hold}")

    def _handler_chejan_data(self, gubun, item_cnt, fid_list):
        if 'gubun' == '1':      # 잔고통보
            예수금 = self.GetChejanData('951')
            예수금 = int(예수금)
            self.amount = int(예수금 * 0.2)
            self.plain_text_edit.appendPlainText(f"투자금액 업데이트 됨: {self.amount}")

    def subscribe_stock_conclusion(self, screen_no):
        self.SetRealReg(screen_no, "229200", "20", 0)
        self.plain_text_edit.appendPlainText("주식체결 구독신청")

    def subscribe_market_time(self, screen_no):
        self.SetRealReg(screen_no, "", "215", 0)
        self.plain_text_edit.appendPlainText("장시작시간 구독신청")

    # TR 요청을 위한 메소드
    def SetInputValue(self, id, value):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    def CommRqData(self, rqname, trcode, next, screen_no):
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", 
                              rqname, trcode, next, screen_no)

    def GetCommData(self, trcode, rqname, index, item):
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", 
                                     trcode, rqname, index, item)
        return data.strip()

    def SendOrder(self, rqname, screen, accno, order_type, code, quantity, price, hoga, order_no):
        self.ocx.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                             [rqname, screen, accno, order_type, code, quantity, price, hoga, order_no])

    def GetChejanData(self, fid):
        data = self.ocx.dynamicCall("GetChejanData(int)", fid)
        return data

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#로깅
import logging

#1단계
logging.debug("debug")
logging.info("info")
logging.warning("warning")
logging.error("error")
logging.critical("critical")

#2단계
logging.basicConfig(level=logging.INFO)

logging.debug("debug")
logging.info("info")
logging.warning("warning")
logging.error("error")
logging.critical("critical")

#파일로 로깅
logging.basicConfig(filename = "mylog.txt", level=logging.INFO)

logging.debug("debug")
logging.info("info")

#Logger 객체 사용하기
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

stream_handler = logging.StreamHandler() # 콘솔 핸들러
file_handler = logging.FileHandler("mylog2.txt") # 파일 핸들러

logger.addHandler(stream_handler)
logger.addHandler(file_handler)

logger.info("this is info")

#로깅 포맷 설정
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s | %(name)s | %(levelname)s | %(message)s')

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)

file_handler = logging.FileHandler("mylog2.txt")
file_handler.setFormatter(formatter)

# add handler to logger
logger.addHandler(stream_handler)
logger.addHandler(file_handler)

logger.info("this is info")

#Handlers
import logging.handlers
import time

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s | %(name)s | %(levelname)s | %(message)s')

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)

file_handler = logging.handlers.TimedRotatingFileHandler(filename="log",when = "M", interval=1)
file_handler.setFormatter(formatter)

# add handler to logger
logger.addHandler(stream_handler)
logger.addHandler(file_handler)

while True:
    logger.info("this is info")
    time.sleep(10)

#GUI와 키움 API 처리 코드 분리하기
import sys
from PyQt5.QtWidgets import *
from pykiwoom import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 300)

        self.kiwoom = Kiwoom()
        self.kiwoom.CommConnect(self.callback_login)

    def callback_login(self, *args, **kwargs):
        if kwargs['err_code'] == 0:
            self.statusBar().showMessage("로그인 완료")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#QTimer
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import threading


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timer_slot)

    def timer_slot(self):
        name = threading.currentThread().getName()
        print(f"timer slot is called by {name}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec()

#타이머와 키움 OCX
import sys 
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import threading
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.login_status = False

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timer_slot)

        # login
        self.CommConnect()

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def _handler_login(self, err_code):
        if err_code == 0:
            self.statusBar().showMessage("login 완료")
            self.login_status = True

    def timer_slot(self):
        thread_name = threading.currentThread().getName()
        print(f"timer slot is called by {thread_name}")
        if self.login_status:
            name = self.GetMasterCodeName("005930")
            print(name)

    def GetMasterCodeName(self, code):
        name = self.ocx.dynamicCall("GetMasterCodeName(QString)", code)
        return name


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec()

#멀티 스레드 기반의 전략과 주문 분리
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import datetime
from multiprocessing import Queue


# 실시간으로 들어오는 데이터를 보고 주문 여부를 판단하는 스레드
class Worker(QThread):
    # argument는 없는 단순 trigger
    # 데이터는 queue를 통해서 전달됨
    trigger = pyqtSignal()

    def __init__(self, data_queue, order_queue):
        super().__init__()
        self.data_queue = data_queue                # 데이터를 받는 용
        self.order_queue = order_queue              # 주문 요청용
        self.timestamp = None
        self.limit_delta = datetime.timedelta(seconds=2)

    def run(self):
        while True:
            if not self.data_queue.empty():
                data = self.data_queue.get()
                result = self.process_data(data)
                if result:
                    self.order_queue.put(data)                      # 주문 Queue에 주문을 넣음
                    self.timestamp = datetime.datetime.now()        # 이전 주문 시간을 기록함
                    self.trigger.emit()

    def process_data(self, data):
        # 시간 제한을 충족하는가?
        time_meet = False
        if self.timestamp is None:
            time_meet = True
        else:
            now = datetime.datetime.now()                           # 현재시간
            delta = now - self.timestamp                            # 현재시간 - 이전 주문 시간
            if delta >= self.limit_delta:
                time_meet = True

        # 알고리즘을 충족하는가?
        algo_meet = False
        if data % 2 == 0:
            algo_meet = True

        # 알고리즘과 주문 가능 시간 조건을 모두 만족하면
        if time_meet and algo_meet:
            return True
        else:
            return False


class MyWindow(QMainWindow):
    def __init__(self, data_queue, order_queue):
        super().__init__()

        # queue
        self.data_queue = data_queue
        self.order_queue = order_queue

        # thread start
        self.worker = Worker(data_queue, order_queue)
        self.worker.trigger.connect(self.pop_order)
        self.worker.start()

        # 데이터가 들어오는 속도는 주문보다 빠름
        self.timer1 = QTimer()
        self.timer1.start(1000)
        self.timer1.timeout.connect(self.push_data)

    def push_data(self):
        now = datetime.datetime.now()
        self.data_queue.put(now.second)

    @pyqtSlot()
    def pop_order(self):
        if not self.order_queue.empty():
            data = self.order_queue.get()
            print(data)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    data_queue = Queue()
    order_queue = Queue()
    window = MyWindow(data_queue, order_queue)
    window.show()
    app.exec_()
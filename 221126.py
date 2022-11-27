import sys
from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import pandas as pd
import os
import threading

class btl_system():
    def __init__(self):
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        print("로그인 시작!")

        self.kiwoom.OnEventConnect.connect(self.login_Connect)
        self.kiwoom.OnReceiveTrData.connect(self.trdata_get)

        self.kiwoom.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

        self.minute_data = {'date': [], 'time': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': [], 'pattern': []}

        global interesting_codes
        interesting_codes = ["005930", "009450", "067160", "214420"]

    def login_Connect(self, err_code):
        if err_code == 0:
            print('로그인 성공했습니다')
        else:
            print('로그인 실패했습니다')
        self.login_event_loop.exit()

    def rq_data_opt10080(self, stock_code, tik, jugagubun):
        self.kiwoom.dynamicCall("SetInputValue(QString,QString)", "종목코드", stock_code)
        self.kiwoom.dynamicCall("SetInputValue(QString,QString)", "틱범위", tik)
        self.kiwoom.dynamicCall("SetInputValue(QString,QString)", "수정주가구분", jugagubun)
        self.kiwoom.dynamicCall("CommRqData(QString,QString,int,QString)", "opt_10080", "opt_10080", 0, "0102")
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def trdata_get(self, sScrNo, rqname, strcode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSplmMsg):
        if rqname == 'opt_10080':
            for i in range(1, 2):
                date = int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "체결시간").strip()[0:8])
                time = int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "체결시간").strip()[8:])
                open = abs(int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "시가").strip()))
                high = abs(int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "고가").strip()))
                low = abs(int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "저가").strip()))
                close = abs(int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "현재가").strip()))
                volume = abs(int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10080", "주식분봉차트조회요청", i, "거래량").strip()))
                pattern = str(high - open) + str(low - open) + str(close - open)

                self.minute_data['date'].append(date)
                self.minute_data['time'].append(time)
                self.minute_data['open'].append(open)
                self.minute_data['high'].append(high)
                self.minute_data['low'].append(low)
                self.minute_data['close'].append(close)
                self.minute_data['volume'].append(volume)
                self.minute_data['pattern'].append(pattern)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    btl = btl_system()

    def rq_data_minute():
        for interesting_code in interesting_codes:
            btl.rq_data_opt10080(interesting_code, "1", 0)
            df_minute_data = pd.DataFrame(btl.minute_data, columns=['date', 'time', 'open', 'high', 'low', 'close', 'volume', 'pattern'])
            df_minute_data = df_minute_data.tail(n=1)
            print(df_minute_data)

            df2 = df_minute_data
            df3 = df_minute_data.head(n=0)

            dir = r'C:\Users\dahyu\OneDrive\문서\복수전공\인공지능 금융투자\minute_data.xlsx'

            if os.path.exists(dir):
                with pd.ExcelWriter(dir, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                    df2.to_excel(writer, header=False, index=False, sheet_name=interesting_code, startrow=writer.sheets[interesting_code].max_row)

            else:
                with pd.ExcelWriter(dir, mode='w', engine='openpyxl') as writer:
                    df2.to_excel(writer, index=False, sheet_name=interesting_codes[0])
                    df3.to_excel(writer, index=False, sheet_name=interesting_codes[1])
                    df3.to_excel(writer, index=False, sheet_name=interesting_codes[2])
                    df3.to_excel(writer, index=False, sheet_name=interesting_codes[3])

        threading.Timer(10, rq_data_minute).start()

rq_data_minute()

app.exec()


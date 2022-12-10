import pandas as pd
import datetime
import time
import sys
import sqlite3
from collections import deque
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5 import uic


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python 로그인")
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)

        # login
        self.ocx.dynamicCall("CommConnect()")
        self.login_loop = QEventLoop()
        self.login_loop.exec()

        self.ocx.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")

    def _handler_login(self):
        try:
            self.login_loop.exit()
        except:
            pass



# 조건검색식 불러오기

    def auto_trading(self):
        
        # callback fn 등록
        self.kw.notify_fn["_on_receive_real_condition"] = self.search_condi

        screen_no = "4000"
        condi_info = self.kw.get_condition_load()
        
        for condi_name, condi_id in condi_info.items():
            # 화면번호, 조건식이름, 조건식ID, 실시간조건검색(1)
            self.kw.send_condition(screen_no, condi_name, int(condi_id), 1)
            time.sleep(0.2)



# 실시간 조건검색식 종목 검출하기

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.logger = KWlog()
        self.tr_mgr = TrManager(self)
        self.evt_loop = QEventLoop()  # lock/release event loop
        self.ret_data = None
        self.req_queue = deque(maxlen=10)
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.tr_controller = TrController(self)
        self.notify_fn = {}

    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._on_event_connect)  # 로긴 이벤트
        self.OnReceiveTrData.connect(self.tr_mgr._on_receive_tr_data)  # tr 수신 이벤트
        self.OnReceiveRealData.connect(self._on_receive_real_data)  # 실시간 시세 이벤트
        self.OnReceiveRealCondition.connect(self._on_receive_real_condition)  # 조건검색 실시간 편입, 이탈종목 이벤트
        self.OnReceiveTrCondition.connect(self._on_receive_tr_condition)  # 조건검색 조회응답 이벤트
        self.OnReceiveConditionVer.connect(self._on_receive_condition_ver)  # 로컬에 사용자조건식 저장 성공여부 응답 이벤트
        self.OnReceiveChejanData.connect(self._on_receive_chejan_data)  # 주문 접수/확인 수신시 이벤트
        self.OnReceiveMsg.connect(self._on_receive_msg)  # 수신 메시지 이벤트


    def _on_receive_real_condition(self, code, event_type, condi_name, condi_index):

        try:
            self.logger.info("_on_receive_real_condition")
            max_char_cnt = 60
            self.logger.info("[실시간 조건 검색 결과]".center(max_char_cnt, '-'))
            data = [
                ("code", code),
                ("event_type", event_type),
                ("condi_name", condi_name),
                ("condi_index", condi_index)
            ]
            max_key_cnt = max(len(d[0]) for d in data) + 3
            for d in data:
                key = ("* " + d[0]).rjust(max_key_cnt)
                self.logger.info("{0}: {1}".format(key, d[1]))
            self.logger.info("-" * max_char_cnt)
            data = dict(data)
            data["kw_event"] = "OnReceiveRealCondition"
            if '_on_receive_real_condition' in self.notify_fn:
                self.notify_fn['_on_receive_real_condition'](data)

        except Exception as e:
            self.logger.error(e)
        finally:
            self.real_condition_search_result = []



# 종목 매수하기

def search_condi(self, event_data): 
    event_data = {
            "code": code, # "066570"
            "event_type": event_type, # "I"(종목편입), "D"(종목이탈)
            "condi_name": condi_name, # "스켈핑"
            "condi_index": condi_index # "004"
        }


    if event_data["event_type"] == "I":
        if self.stock_account["계좌정보"]["예수금"] < 100000: 
            return
        curr_price = self.kw.get_curr_price(event_data["code"])
        quantity = int(100000/curr_price)
        self.kw.reg_callback("OnReceiveChejanData", ("조건식매수", "5000"), self.update_account)
        self.kw.send_order("조건식매수", "5000", self.acc_no, 1, event_data["code"], quantity, 0, "03", "")  



# 계좌 업데이트

def set_account(self):
    self.acc_no = self.kw.get_login_info("ACCNO")
    self.acc_no = self.acc_no.strip(";")  # 계좌 1개를 가정함.
    self.stock_account = self.kw.계좌평가현황요청("계좌평가현황요청", self.acc_no, "", "1", "6001")

def update_account(self):
    self.stock_account = self.kw.계좌평가현황요청("계좌평가현황요청", self.acc_no, "", "1", "6001")        



# 종목 매도하기

def start_timer(self):
    if self.timer:
        self.timer.stop()
        self.timer.deleteLater()
    self.timer = QTimer()
    self.timer.timeout.connect(self.sell)
    # self.timer.setSingleShot(True)
    self.timer.start(30000) # 30 sec interval

def sell(self):
    self.update_account()
    print("=" * 50)
    print("현재 계좌 현황입니다...")
    for data in self.stock_account["종목정보"]:
        stock_name, code, quantity = data["종목코드"], data["종목명"], data["보유수량"]
        print("* 종목: {}, 손익율: {}%, 보유수량: {}, 평가금액: {}원".format(
            data["종목명"], ("%.2f" % data["손익율"]), int(data["보유수량"]), format(int(data["평가금액"]), ',')
        ))
        if data["손익율"] > 3.0:
            print("시장가로 물량 전부 익절합니다. [{}, {}주]".format(stock_name, quantity))
            self.kw.send_order("익절매도", "5001", self.acc_no, 2, code, quantity, 0, "03", "")
        elif data["손익율"] < -2.0:
            print("시장가로 물량 전부 손절합니다. [{}, {}주]".format(stock_name, quantity))
            self.kw.send_order("손절매도", "5002", self.acc_no, 2, code, quantity, 0, "03", "")



 # 트레이딩 이력 분석하기       

def search_condi(self, event_data):
    event_data = {
            "code": code, # "066570"
            "event_type": event_type, # "I"(종목편입), "D"(종목이탈)
            "condi_name": condi_name, # "스켈핑"
            "condi_index": condi_index # "004"
        }
    curr_time = datetime.today()
    # 실시간 조건검색 이력정보
    self.tt_db.real_condi_search.insert({
        'date': curr_time,
        'code': event_data["code"],
        'stock_name': self.stock_dict[event_data["code"]]["stock_name"],
        'market': self.stock_dict[event_data["code"]]["market"],
        'event': event_data["event_type"],
        'condi_name': event_data["condi_name"]
    })

    if event_data["event_type"] == "I":
        if self.stock_account["계좌정보"]["예수금"] < 100000:  # 잔고가 10만원 미만이면 매수 안함
            return
        curr_price = self.kw.get_curr_price(event_data["code"])
        quantity = int(100000/curr_price)
        self.kw.reg_callback("OnReceiveChejanData", ("조건식매수", "5000"), self.update_account)
        self.tt_db.trading_history.insert({
            'date': curr_time,
            'code': event_data["code"],
            'stock_name': self.stock_dict[event_data["code"]]["stock_name"],
            'market': self.stock_dict[event_data["code"]]["market"],
            'event': event_data["event_type"],
            'condi_name': event_data["condi_name"],
            'trade': 'buy',
            'quantity': quantity,
            'hoga_gubun': '시장가',
            'account_no': self.acc_no
        })
        self.kw.send_order("조건식매수", "5000", self.acc_no, 1, event_data["code"], quantity, 0, "03", "")    


def sell(self):
    self.update_account()
    curr_time = datetime.today()
    print("=" * 50)
    print("현재 계좌 현황입니다.")
    for data in self.stock_account["종목정보"]:
        stock_name, code, quantity = data["종목코드"], data["종목명"], data["보유수량"]
        print("* 종목: {}, 손익율: {}%, 보유수량: {}, 평가금액: {}원".format(
            data["종목명"], ("%.2f" % data["손익율"]), int(data["보유수량"]), format(int(data["평가금액"]), ',')
        ))

        if data["손익율"] > 3.0 or data["손익율"] < -2.0:
            if data["손익율"] > 0:
                print("시장가로 물량 전부 익절합니다. [{}, {}주]".format(stock_name, quantity))
            else:
                print("시장가로 물량 전부 손절합니다. [{}, {}주]".format(stock_name, quantity))

            self.kw.reg_callback("OnReceiveChejanData", ("시장가매도", "5001"), self.update_account)
            self.kw.send_order("시장가매도", "5001", self.acc_no, 2, code, quantity, 0, "03", "")
            self.tt_db.trading_history.insert({
                'date': curr_time,
                'code': code,
                'stock_name': self.stock_dict[code]["stock_name"],
                'market': self.stock_dict[code]["market"],
                'event': '',
                'condi_name': '',
                'trade': 'sell',
                'profit': data["손익율"],
                'quantity': quantity,
                'hoga_gubun': '시장가',
                'account_no': self.acc_no
            })

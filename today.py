import sys
from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import time
import pandas
 
 
class system_trading():
	def __init__(self):
		self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
		print("연결되었습니다.")
 
		self.kiwoom.OnEventConnect.connect(self.OnEventConnect)
        
		self.kiwoom.dynamicCall("CommConnect()")
		self.login_event_loop = QEventLoop()
		self.login_event_loop.exec_()
        
		self.kiwoom.OnReceiveTrData.connect(self.OnReceiveTrData)
 
		self.day_data = {'date':[], 'open':[], 'high':[], 'low':[], 'close':[], 'volume':[], 'trade_volume':[]}
 
	def OnEventConnect(self, err_code):
		if err_code == 0:
			print("로그인에 성공하였습니다.")
		else:
			print("로그인에 실패하였습니다.")
		self.login_event_loop.exit()
 
	def OnReceiveTrData(self, scrno, rqname, trcode, recordname, prenext, unused1, unused2, unused3, unused4):
		if prenext == '2':
			self.remained_data = True
		elif prenext != '2':
			self.remained_data = False
 
		if rqname == 'hello':
			self.opt10080()
 
		try:
			self.tr_event_loop.exit()
		except AttributeError:
			pass
 
	def opt10080(self):
		getrepeatcnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", "opt10080", "주식일봉차트조회요청")
 
		for i in range(getrepeatcnt):
			item_code = self.GetCommData("opt10080", "주식일봉차트조회요청", 0, "종목코드").strip()
			date = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "일자").strip()
			open = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "시가").strip()
			high = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "고가").strip()
			low = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "저가").strip()
			close = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "현재가").strip()
			volume = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "거래량").strip()
			trade_volume = self.GetCommData("opt10080", "주식일봉차트조회요청", i, "거래대금").strip()
 
			self.day_data['date'].append(date)
			self.day_data['open'].append(open)
			self.day_data['high'].append(high)
			self.day_data['low'].append(low)
			self.day_data['close'].append(close)
			self.day_data['volume'].append(volume)
			self.day_data['trade_volume'].append(trade_volume)
 
	def rq_chart_data(self, itemcode, date, justify):
		print("차트를 조회합니다.")
		self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", itemcode)
		self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)
		self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", justify)
		self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "hello", "opt10080", 0, "0101")
		self.tr_event_loop = QEventLoop()
		self.tr_event_loop.exec_()
 
		while self.remained_data == True:
			time.sleep(0.3)
			self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", itemcode)
			self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)
			self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", justify)
			self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "hello", "opt10080", 2, "0101")
			self.tr_event_loop = QEventLoop()
			self.tr_event_loop.exec_()
 
	def GetCommData(self, trcode, recordname, index, itemname):
		result = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, recordname, index, itemname)
		return result
 
if __name__ == "__main__":
	app = QApplication(sys.argv)
	trade = system_trading()
 
	trade.rq_chart_data("005930", "20210530", 1)
	df_day_data = pandas.DataFrame(trade.day_data, columns=['date', 'open', 'high', 'low', 'close', 'volume', 'trade_volume'])
	print(df_day_data)
import sys
import logging
import logging.config
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop
from PyQt5.QtWidgets import QApplication
from pandas import DataFrame


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()

        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        # Loop 변수
        self.loginLoop = None
        self.requestLoop = None
        self.orderLoop = None
        self.conditionLoop = None

        # 서버구분
        self.server = None

        # 조건식
        self.condition = None

        # 에러
        self.error = None

        # 주문번호
        self.orderNo = ""

        # 조회
        self.inquiry = 0

        # 서버에서 받은 메시지
        self.msg = ""

        # 예수금 d+2
        self.opw00001Data = 0

        # 보유종목 정보
        self.opw00018Data = {'accountEvaluation': [], 'stocks': []}

        # signal & slot
        self.OnEventConnect.connect(self.eventConnect)
        self.OnReceiveTrData.connect(self.receiveTrData)
        self.OnReceiveChejanData.connect(self.receiveChejanData)
        self.OnReceiveRealData.connect(self.receiveRealData)
        self.OnReceiveMsg.connect(self.receiveMsg)
        self.OnReceiveConditionVer.connect(self.receiveConditionVer)
        self.OnReceiveTrCondition.connect(self.receiveTrCondition)
        self.OnReceiveRealCondition.connect(self.receiveRealCondition)

        # 로깅용 설정파일
        logging.config.fileConfig('logging.conf')
        self.log = logging.getLogger('Kiwoom')

    # 로깅용 메서드 정의                                               

    def logger(origin):
        def wrapper(*args, **kwargs):
            args[0].log.debug('{} args - {}, kwargs - {}'.format(origin.__name__, args, kwargs))
            return origin(*args, **kwargs)

        return wrapper

    # 이벤트 정의                                                    

    def eventConnect(self, returnCode):
        """
        통신 연결 상태 변경시 이벤트

        returnCode가 0이면 로그인 성공
        그 외에는 ReturnCode 클래스 참조.

        :param returnCode: int
        """

        try:
            if returnCode == ReturnCode.OP_ERR_NONE:

                self.server = self.getLoginInfo("GetServerGubun", True)

                if len(self.server) == 0 or self.server != "1":
                    self.msg += "실서버 연결 성공" + "\r\n\r\n"

                else:
                    self.msg += "모의투자서버 연결 성공" + "\r\n\r\n"

            else:
                self.msg += "연결 끊김: 원인 - " + ReturnCode.CAUSE[returnCode] + "\r\n\r\n"

        except Exception as error:
            self.log.error('eventConnect {}'.format(error))

        finally:
            try:
                self.loginLoop.exit()
            except AttributeError:
                pass

    def receiveMsg(self, screenNo, requestName, trCode, msg):

        self.msg += requestName + ": " + msg 

    def receiveTrData(self, screenNo, requestName, trCode, recordName, inquiry,
                      deprecated1, deprecated2, deprecated3, deprecated4):
        print("receiveTrData 실행: ", screenNo, requestName, trCode, recordName, inquiry)

        # 주문번호와 주문루프
        self.orderNo = self.commGetData(trCode, "", requestName, 0, "주문번호")

        try:
            self.orderLoop.exit()
        except AttributeError:
            pass

        self.inquiry = inquiry

        if requestName == "관심종목정보요청":
            data = self.getCommDataEx(trCode, "관심종목정보")
            print(type(data))
            print(data)
        elif requestName == "주식일봉차트조회요청":
            data = self.getCommDataEx(trCode, "주식일봉차트조회")

            colName = ['종목코드', '현재가', '거래량', '거래대금', '일자', '시가', '고가', '저가',
                       '수정주가구분', '수정비율', '대업종구분', '소업종구분', '종목정보', '수정주가이벤트', '전일종가']

            data = DataFrame(data, columns=colName)

            print(type(data))
            print(data.head(5))

        elif requestName == "예수금상세현황요청":
            deposit = self.commGetData(trCode, "", requestName, 0, "d+2추정예수금")
            deposit = self.changeFormat(deposit)
            self.opw00001Data = deposit

        elif requestName == "계좌평가잔고내역요청":
            # 계좌 평가 정보
            accountEvaluation = []
            keyList = ["총매입금액", "총평가금액", "총평가손익금액", "총수익률(%)", "추정예탁자산"]

            for key in keyList:
                value = self.commGetData(trCode, "", requestName, 0, key)

                if key.startswith("총수익률"):
                    value = self.changeFormat(value, 1)
                else:
                    value = self.changeFormat(value)

                accountEvaluation.append(value)

            self.opw00018Data['accountEvaluation'] = accountEvaluation

            # 보유 종목 정보
            cnt = self.getRepeatCnt(trCode, requestName)
            keyList = ["종목명", "보유수량", "매입가", "현재가", "평가손익", "수익률(%)"]

            for i in range(cnt):
                stock = []

                for key in keyList:
                    value = self.commGetData(trCode, "", requestName, i, key)

                    if key.startswith("수익률"):
                        value = self.changeFormat(value, 2)
                    elif key != "종목명":
                        value = self.changeFormat(value)

                    stock.append(value)

                self.opw00018Data['stocks'].append(stock)

        try:
            self.requestLoop.exit()
        except AttributeError:
            pass

    def receiveRealData(self, code, realType, realData):
       
        try:
            self.log.debug("[receiveRealData]")
            self.log.debug("({})".format(realType))

            if realType not in RealType.REALTYPE:
                return

            data = []

            if code != "":
                data.append(code)
                codeOrNot = code
            else:
                codeOrNot = realType

            for fid in sorted(RealType.REALTYPE[realType].keys()):
                value = self.getCommRealData(codeOrNot, fid)
                data.append(value)

            # TODO: DB에 저장
            self.log.debug(data)

        except Exception as e:
            self.log.error('{}'.format(e))

    def receiveChejanData(self, gubun, itemCnt, fidList):
       
        fids = fidList.split(';')
        print("[receiveChejanData]")
        print("gubun: ", gubun, "itemCnt: ", itemCnt, "fidList: ", fidList)
        print("========================================")
        print("[ 구분: ", self.getChejanData(913) if '913' in fids else '잔고통보', "]")
        for fid in fids:
            print(FidList.CHEJAN[int(fid)] if int(fid) in FidList.CHEJAN else fid, ": ", self.getChejanData(int(fid)))
        print("========================================")

    # 메서드 정의: 로그인 관련 메서드                                   
    def commConnect(self):
        self.dynamicCall("CommConnect()")
        self.loginLoop = QEventLoop()
        self.loginLoop.exec_()

    def getConnectState(self):
        state = self.dynamicCall("GetConnectState()")
        return state

    def getLoginInfo(self, tag, isConnectState=False):
        if not isConnectState:
            if not self.getConnectState():
                raise KiwoomConnectError()

        if not isinstance(tag, str):
            raise ParameterTypeError()

        if tag not in ['ACCOUNT_CNT', 'ACCNO', 'USER_ID', 'USER_NAME', 'GetServerGubun']:
            raise ParameterValueError()

        if tag == "GetServerGubun":
            info = self.getServerGubun()
        else:
            cmd = 'GetLoginInfo("%s")' % tag
            info = self.dynamicCall(cmd)

        return info

    def getServerGubun(self):
        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret

    #################################################################
    # 메서드 정의: 조회 관련 메서드                                        #
    # 시세조회, 관심종목 조회, 조건검색 등 이들의 합산 조회 횟수가 1초에 5회까지 허용 #
    #################################################################

    def setInputValue(self, key, value):
        if not (isinstance(key, str) and isinstance(value, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetInputValue(QString, QString)", key, value)

    def commRqData(self, requestName, trCode, inquiry, screenNo): #TR요청

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(requestName, str)
                and isinstance(trCode, str)
                and isinstance(inquiry, int)
                and isinstance(screenNo, str)):

            raise ParameterTypeError()

        returnCode = self.dynamicCall("CommRqData(QString, QString, int, QString)", requestName, trCode, inquiry, screenNo)

        if returnCode != ReturnCode.OP_ERR_NONE:
            raise KiwoomProcessingError("commRqData(): " + ReturnCode.CAUSE[returnCode])

        # 루프 생성: receiveTrData() 메서드에서 루프를 종료시킨다.
        self.requestLoop = QEventLoop()
        self.requestLoop.exec_()

    def commGetData(self, trCode, realType, requestName, index, key): #데이터획득

        return self.getCommData(trCode, requestName, index, key)

    def getCommData(self, trCode, requestName, index, key): #데이터획득

        if not (isinstance(trCode, str)
                and isinstance(requestName, str)
                and isinstance(index, int)
                and isinstance(key, str)):
            raise ParameterTypeError()

        data = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                trCode, requestName, index, key)
        return data.strip()

    def getRepeatCnt(self, trCode, requestName): #데이터 개수

        if not (isinstance(trCode, str)
                and isinstance(requestName, str)):
            raise ParameterTypeError()

        count = self.dynamicCall("GetRepeatCnt(QString, QString)", trCode, requestName)
        return count

    def getCommDataEx(self, trCode, multiDataName):

        if not (isinstance(trCode, str)
                and isinstance(multiDataName, str)):
            raise ParameterTypeError()

        data = self.dynamicCall("GetCommDataEx(QString, QString)", trCode, multiDataName)
        return data

    def commKwRqData(self, codes, inquiry, codeCount, requestName, screenNo, typeFlag=0):
       
        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(codes, str)
                and isinstance(inquiry, int)
                and isinstance(codeCount, int)
                and isinstance(requestName, str)
                and isinstance(screenNo, str)
                and isinstance(typeFlag, int)):

            raise ParameterTypeError()

        returnCode = self.dynamicCall("CommKwRqData(QString, QBoolean, int, int, QString, QString)",
                                      codes, inquiry, codeCount, typeFlag, requestName, screenNo)

        if returnCode != ReturnCode.OP_ERR_NONE:
            raise KiwoomProcessingError("commKwRqData(): " + ReturnCode.CAUSE[returnCode])

        # 루프 생성: receiveTrData() 메서드에서 루프를 종료시킨다.
        self.requestLoop = QEventLoop()
        self.requestLoop.exec_()

    ###############################################################
    # 메서드 정의: 실시간 데이터 처리 관련 메서드                           #
    ###############################################################

    def disconnectRealData(self, screenNo):
        if not self.getConnectState():
            raise KiwoomConnectError()

        if not isinstance(screenNo, str):
            raise ParameterTypeError()

        self.dynamicCall("DisconnectRealData(QString)", screenNo)

    def getCommRealData(self, code, fid):
       
        if not (isinstance(code, str)
                and isinstance(fid, int)):
            raise ParameterTypeError()

        value = self.dynamicCall("GetCommRealData(QString, int)", code, fid)

        return value

    def setRealReg(self, screenNo, codes, fids, realRegType):
       
        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(codes, str)
                and isinstance(fids, str)
                and isinstance(realRegType, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetRealReg(QString, QString, QString, QString)",
                         screenNo, codes, fids, realRegType)

    def setRealRemove(self, screenNo, code):
        
        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(code, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetRealRemove(QString, QString)", screenNo, code)

  
    # 메서드 정의: 조건검색 관련 메서드와 이벤트                          

    def receiveConditionVer(self, receive, msg):
       
        try:
            if not receive:
                return

            self.condition = self.getConditionNameList()
            print("조건식 개수: ", len(self.condition))

            for key in self.condition.keys():
                print("조건식: ", key, ": ", self.condition[key])
                print("key type: ", type(key))

        except Exception as e:
            print(e)

        finally:
            self.conditionLoop.exit()

    def receiveTrCondition(self, screenNo, codes, conditionName, conditionIndex, inquiry):
        
        print("[receiveTrCondition]")

        try:
            if codes == "":
                return

            codeList = codes.split(';')
            del codeList[-1]

            print(codeList)
            print("종목개수: ", len(codeList))

        finally:
            self.conditionLoop.exit()

    def receiveRealCondition(self, code, event, conditionName, conditionIndex):
       
        print("[receiveRealCondition]")

        print("종목코드: ", code)
        print("이벤트: ", "종목편입" if event == "I" else "종목이탈")

    def getConditionLoad(self):
        """ 조건식 목록 요청 메서드 """

        if not self.getConnectState():
            raise KiwoomConnectError()

        isLoad = self.dynamicCall("GetConditionLoad()")

        # 요청 실패시
        if not isLoad:
            raise KiwoomProcessingError("getConditionLoad(): 조건식 요청 실패")

        # receiveConditionVer() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()

    def getConditionNameList(self):
       
        data = self.dynamicCall("GetConditionNameList()")

        if data == "":
            raise KiwoomProcessingError("getConditionNameList(): 사용자 조건식이 없습니다.")

        conditionList = data.split(';')
        del conditionList[-1]

        conditionDictionary = {}

        for condition in conditionList:
            key, value = condition.split('^')
            conditionDictionary[int(key)] = value

        return conditionDictionary

    def sendCondition(self, screenNo, conditionName, conditionIndex, isRealTime):
        

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(conditionName, str)
                and isinstance(conditionIndex, int)
                and isinstance(isRealTime, int)):
            raise ParameterTypeError()

        isRequest = self.dynamicCall("SendCondition(QString, QString, int, int",
                                     screenNo, conditionName, conditionIndex, isRealTime)

        if not isRequest:
            raise KiwoomProcessingError("sendCondition(): 조건검색 요청 실패")

        # receiveTrCondition() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()

    def sendConditionStop(self, screenNo, conditionName, conditionIndex):
        """ 종목 조건검색 중지 메서드 """

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(conditionName, str)
                and isinstance(conditionIndex, int)):
            raise ParameterTypeError()

        self.dynamicCall("SendConditionStop(QString, QString, int)", screenNo, conditionName, conditionIndex)

  
    # 메서드 정의: 주문과 잔고처리 관련 메서드                              
    # 1초에 5회까지 주문 허용                                          

    def sendOrder(self, requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo):
        #주문
    
        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(requestName, str)
                and isinstance(screenNo, str)
                and isinstance(accountNo, str)
                and isinstance(orderType, int)
                and isinstance(code, str)
                and isinstance(qty, int)
                and isinstance(price, int)
                and isinstance(hogaType, str)
                and isinstance(originOrderNo, str)):

            raise ParameterTypeError()

        returnCode = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                      [requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo])

        if returnCode != ReturnCode.OP_ERR_NONE:
            raise KiwoomProcessingError("sendOrder(): " + ReturnCode.CAUSE[returnCode])

        # receiveTrData() 에서 루프종료
        self.orderLoop = QEventLoop()
        self.orderLoop.exec_()

    def getChejanData(self, fid):
      
        if not isinstance(fid, int):
            raise ParameterTypeError()

        cmd = 'GetChejanData("%s")' % fid
        data = self.dynamicCall(cmd)
        return data


    def getCodeListByMarket(self, market):
      
        if not self.getConnectState():
            raise KiwoomConnectError()

        if not isinstance(market, str):
            raise ParameterTypeError()

        if market not in ['0', '3', '4', '5', '6', '8', '9', '10', '30']:
            raise ParameterValueError()

        cmd = 'GetCodeListByMarket("%s")' % market
        codeList = self.dynamicCall(cmd)
        return codeList.split(';')

    def getCodeList(self, *market):
       

        codeList = []

        for m in market:
            tmpList = self.getCodeListByMarket(m)
            codeList += tmpList

        return codeList

    def getMasterCodeName(self, code):
        

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not isinstance(code, str):
            raise ParameterTypeError()

        cmd = 'GetMasterCodeName("%s")' % code
        name = self.dynamicCall(cmd)
        return name

    def changeFormat(self, data, percent=0):

        if percent == 0:
            d = int(data)
            formatData = '{:-,d}'.format(d)

        elif percent == 1:
            f = int(data) / 100
            formatData = '{:-,.2f}'.format(f)

        elif percent == 2:
            f = float(data)
            formatData = '{:-,.2f}'.format(f)

        return formatData

    def opwDataReset(self):
        """ 잔고 및 보유종목 데이터 초기화 """
        self.opw00001Data = 0
        self.opw00018Data = {'accountEvaluation': [], 'stocks': []}


class ParameterTypeError(Exception):
    """ 파라미터 타입이 일치하지 않을 경우 발생하는 예외 """

    def __init__(self, msg="파라미터 타입이 일치하지 않습니다."):
        self.msg = msg

    def __str__(self):
        return self.msg


class ParameterValueError(Exception):
    """ 파라미터로 사용할 수 없는 값을 사용할 경우 발생하는 예외 """

    def __init__(self, msg="파라미터로 사용할 수 없는 값 입니다."):
        self.msg = msg

    def __str__(self):
        return self.msg


class KiwoomProcessingError(Exception):
    """ 키움에서 처리실패에 관련된 리턴코드를 받았을 경우 발생하는 예외 """

    def __init__(self, msg="처리 실패"):
        self.msg = msg

    def __str__(self):
        return self.msg

    def __repr__(self):
        return self.msg


class KiwoomConnectError(Exception):
    """ 키움서버에 로그인 상태가 아닐 경우 발생하는 예외 """

    def __init__(self, msg="로그인 여부를 확인하십시오"):
        self.msg = msg

    def __str__(self):
        return self.msg


class ReturnCode(object):
    """ 키움 OpenApi+ 함수들이 반환하는 값 """

    OP_ERR_NONE = 0 # 정상처리
    OP_ERR_FAIL = -10   # 실패
    OP_ERR_LOGIN = -100 # 사용자정보교환실패
    OP_ERR_CONNECT = -101   # 서버접속실패
    OP_ERR_VERSION = -102   # 버전처리실패
    OP_ERR_FIREWALL = -103  # 개인방화벽실패
    OP_ERR_MEMORY = -104    # 메모리보호실패
    OP_ERR_INPUT = -105 # 함수입력값오류
    OP_ERR_SOCKET_CLOSED = -106 # 통신연결종료
    OP_ERR_SISE_OVERFLOW = -200 # 시세조회과부하
    OP_ERR_RQ_STRUCT_FAIL = -201    # 전문작성초기화실패
    OP_ERR_RQ_STRING_FAIL = -202    # 전문작성입력값오류
    OP_ERR_NO_DATA = -203   # 데이터없음
    OP_ERR_OVER_MAX_DATA = -204 # 조회가능한종목수초과
    OP_ERR_DATA_RCV_FAIL = -205 # 데이터수신실패
    OP_ERR_OVER_MAX_FID = -206  # 조회가능한FID수초과
    OP_ERR_REAL_CANCEL = -207   # 실시간해제오류
    OP_ERR_ORD_WRONG_INPUT = -300   # 입력값오류
    OP_ERR_ORD_WRONG_ACCTNO = -301  # 계좌비밀번호없음
    OP_ERR_OTHER_ACC_USE = -302 # 타인계좌사용오류
    OP_ERR_MIS_2BILL_EXC = -303 # 주문가격이20억원을초과
    OP_ERR_MIS_5BILL_EXC = -304 # 주문가격이50억원을초과
    OP_ERR_MIS_1PER_EXC = -305  # 주문수량이총발행주수의1%초과오류
    OP_ERR_MIS_3PER_EXC = -306  # 주문수량이총발행주수의3%초과오류
    OP_ERR_SEND_FAIL = -307 # 주문전송실패
    OP_ERR_ORD_OVERFLOW = -308  # 주문전송과부하
    OP_ERR_MIS_300CNT_EXC = -309    # 주문수량300계약초과
    OP_ERR_MIS_500CNT_EXC = -310    # 주문수량500계약초과
    OP_ERR_ORD_WRONG_ACCTINFO = -340    # 계좌정보없음
    OP_ERR_ORD_SYMCODE_EMPTY = -500 # 종목코드없음

    CAUSE = {
        0: '정상처리',
        -10: '실패',
        -100: '사용자정보교환실패',
        -102: '버전처리실패',
        -103: '개인방화벽실패',
        -104: '메모리보호실패',
        -105: '함수입력값오류',
        -106: '통신연결종료',
        -200: '시세조회과부하',
        -201: '전문작성초기화실패',
        -202: '전문작성입력값오류',
        -203: '데이터없음',
        -204: '조회가능한종목수초과',
        -205: '데이터수신실패',
        -206: '조회가능한FID수초과',
        -207: '실시간해제오류',
        -300: '입력값오류',
        -301: '계좌비밀번호없음',
        -302: '타인계좌사용오류',
        -303: '주문가격이20억원을초과',
        -304: '주문가격이50억원을초과',
        -305: '주문수량이총발행주수의1%초과오류',
        -306: '주문수량이총발행주수의3%초과오류',
        -307: '주문전송실패',
        -308: '주문전송과부하',
        -309: '주문수량300계약초과',
        -310: '주문수량500계약초과',
        -340: '계좌정보없음',
        -500: '종목코드없음'
    }


class FidList(object):
    """ receiveChejanData() 이벤트 메서드로 전달되는 FID 목록 """

    CHEJAN = {
        9201: '계좌번호',
        9203: '주문번호',
        9205: '관리자사번',
        9001: '종목코드',
        912: '주문업무분류',
        913: '주문상태',
        302: '종목명',
        900: '주문수량',
        901: '주문가격',
        902: '미체결수량',
        903: '체결누계금액',
        904: '원주문번호',
        905: '주문구분',
        906: '매매구분',
        907: '매도수구분',
        908: '주문/체결시간',
        909: '체결번호',
        910: '체결가',
        911: '체결량',
        10: '현재가',
        27: '(최우선)매도호가',
        28: '(최우선)매수호가',
        914: '단위체결가',
        915: '단위체결량',
        938: '당일매매수수료',
        939: '당일매매세금',
        919: '거부사유',
        920: '화면번호',
        921: '921',
        922: '922',
        923: '923',
        949: '949',
        10010: '10010',
        917: '신용구분',
        916: '대출일',
        930: '보유수량',
        931: '매입단가',
        932: '총매입가',
        933: '주문가능수량',
        945: '당일순매수수량',
        946: '매도/매수구분',
        950: '당일총매도손일',
        951: '예수금',
        307: '기준가',
        8019: '손익율',
        957: '신용금액',
        958: '신용이자',
        959: '담보대출수량',
        924: '924',
        918: '만기일',
        990: '당일실현손익(유가)',
        991: '당일신현손익률(유가)',
        992: '당일실현손익(신용)',
        993: '당일실현손익률(신용)',
        397: '파생상품거래단위',
        305: '상한가',
        306: '하한가'
    }


class RealType(object):

    REALTYPE = {
        '주식시세': {
            10: '현재가',
            11: '전일대비',
            12: '등락율',
            27: '최우선매도호가',
            28: '최우선매수호가',
            13: '누적거래량',
            14: '누적거래대금',
            16: '시가',
            17: '고가',
            18: '저가',
            25: '전일대비기호',
            26: '전일거래량대비',
            29: '거래대금증감',
            30: '거일거래량대비',
            31: '거래회전율',
            32: '거래비용',
            311: '시가총액(억)'
        },

        '주식체결': {
            20: '체결시간(HHMMSS)',
            10: '체결가',
            11: '전일대비',
            12: '등락율',
            27: '최우선매도호가',
            28: '최우선매수호가',
            15: '체결량',
            13: '누적체결량',
            14: '누적거래대금',
            16: '시가',
            17: '고가',
            18: '저가',
            25: '전일대비기호',
            26: '전일거래량대비',
            29: '거래대금증감',
            30: '전일거래량대비',
            31: '거래회전율',
            32: '거래비용',
            228: '체결강도',
            311: '시가총액(억)',
            290: '장구분',
            691: 'KO접근도'
        },

        '주식호가잔량': {
            21: '호가시간',
            41: '매도호가1',
            61: '매도호가수량1',
            81: '매도호가직전대비1',
            51: '매수호가1',
            71: '매수호가수량1',
            91: '매수호가직전대비1',
            42: '매도호가2',
            62: '매도호가수량2',
            82: '매도호가직전대비2',
            52: '매수호가2',
            72: '매수호가수량2',
            92: '매수호가직전대비2',
            43: '매도호가3',
            63: '매도호가수량3',
            83: '매도호가직전대비3',
            53: '매수호가3',
            73: '매수호가수량3',
            93: '매수호가직전대비3',
            44: '매도호가4',
            64: '매도호가수량4',
            84: '매도호가직전대비4',
            54: '매수호가4',
            74: '매수호가수량4',
            94: '매수호가직전대비4',
            45: '매도호가5',
            65: '매도호가수량5',
            85: '매도호가직전대비5',
            55: '매수호가5',
            75: '매수호가수량5',
            95: '매수호가직전대비5',
            46: '매도호가6',
            66: '매도호가수량6',
            86: '매도호가직전대비6',
            56: '매수호가6',
            76: '매수호가수량6',
            96: '매수호가직전대비6',
            47: '매도호가7',
            67: '매도호가수량7',
            87: '매도호가직전대비7',
            57: '매수호가7',
            77: '매수호가수량7',
            97: '매수호가직전대비7',
            48: '매도호가8',
            68: '매도호가수량8',
            88: '매도호가직전대비8',
            58: '매수호가8',
            78: '매수호가수량8',
            98: '매수호가직전대비8',
            49: '매도호가9',
            69: '매도호가수량9',
            89: '매도호가직전대비9',
            59: '매수호가9',
            79: '매수호가수량9',
            99: '매수호가직전대비9',
            50: '매도호가10',
            70: '매도호가수량10',
            90: '매도호가직전대비10',
            60: '매수호가10',
            80: '매수호가수량10',
            100: '매수호가직전대비10',
            121: '매도호가총잔량',
            122: '매도호가총잔량직전대비',
            125: '매수호가총잔량',
            126: '매수호가총잔량직전대비',
            23: '예상체결가',
            24: '예상체결수량',
            128: '순매수잔량(총매수잔량-총매도잔량)',
            129: '매수비율',
            138: '순매도잔량(총매도잔량-총매수잔량)',
            139: '매도비율',
            200: '예상체결가전일종가대비',
            201: '예상체결가전일종가대비등락율',
            238: '예상체결가전일종가대비기호',
            291: '예상체결가',
            292: '예상체결량',
            293: '예상체결가전일대비기호',
            294: '예상체결가전일대비',
            295: '예상체결가전일대비등락율',
            13: '누적거래량',
            299: '전일거래량대비예상체결률',
            215: '장운영구분'
        },

        '장시작시간': {
            215: '장운영구분(0:장시작전, 2:장종료전, 3:장시작, 4,8:장종료, 9:장마감)',
            20: '시간(HHMMSS)',
            214: '장시작예상잔여시간'
        },

        '업종지수': {
            20: '체결시간',
            10: '현재가',
            11: '전일대비',
            12: '등락율',
            15: '거래량',
            13: '누적거래량',
            14: '누적거래대금',
            16: '시가',
            17: '고가',
            18: '저가',
            25: '전일대비기호',
            26: '전일거래량대비(계약,주)'
        },

        '업종등락': {
            20: '체결시간',
            252: '상승종목수',
            251: '상한종목수',
            253: '보합종목수',
            255: '하락종목수',
            254: '하한종목수',
            13: '누적거래량',
            14: '누적거래대금',
            10: '현재가',
            11: '전일대비',
            12: '등락율',
            256: '거래형성종목수',
            257: '거래형성비율',
            25: '전일대비기호'
        },

        '주문체결': {
            9201: '계좌번호',
            9203: '주문번호',
            9205: '관리자사번',
            9001: '종목코드',
            912: '주문분류(jj:주식주문)',
            913: '주문상태(10:원주문, 11:정정주문, 12:취소주문, 20:주문확인, 21:정정확인, 22:취소확인, 90,92:주문거부)',
            302: '종목명',
            900: '주문수량',
            901: '주문가격',
            902: '미체결수량',
            903: '체결누계금액',
            904: '원주문번호',
            905: '주문구분(+:현금매수, -:현금매도)',
            906: '매매구분(보통, 시장가등)',
            907: '매도수구분(1:매도, 2:매수)',
            908: '체결시간(HHMMSS)',
            909: '체결번호',
            910: '체결가',
            911: '체결량',
            10: '체결가',
            27: '최우선매도호가',
            28: '최우선매수호가',
            914: '단위체결가',
            915: '단위체결량',
            938: '당일매매수수료',
            939: '당일매매세금'
        },

        '잔고': {
            9201: '계좌번호',
            9001: '종목코드',
            302: '종목명',
            10: '현재가',
            930: '보유수량',
            931: '매입단가',
            932: '총매입가',
            933: '주문가능수량',
            945: '당일순매수량',
            946: '매도매수구분',
            950: '당일총매도손익',
            951: '예수금',
            27: '최우선매도호가',
            28: '최우선매수호가',
            307: '기준가',
            8019: '손익율'
        },

        '주식시간외호가': {
            21: '호가시간(HHMMSS)',
            131: '시간외매도호가총잔량',
            132: '시간외매도호가총잔량직전대비',
            135: '시간외매수호가총잔량',
            136: '시간외매수호가총잔량직전대비'
        }
    }


if __name__ == "__main__":
    """ 조건검색 테스트 코드 """

    app = QApplication(sys.argv)

    try:
        kiwoom = Kiwoom()
        kiwoom.commConnect()

        server = kiwoom.getServerGubun()
        print("server: ", server)
        print("type: ", type(server))
        print("len: ", len(server))

        if len(server) == 0 or server != "1":
            print("실서버 입니다.")

        else:
            print("모의투자 서버입니다.")

    except Exception as e:
        print(e)

    sys.exit(app.exec_())

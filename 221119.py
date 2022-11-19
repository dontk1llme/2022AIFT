from pykiwoom.kiwoom import *
import pprint

kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

# 3.OpenAPI 이용하기 ---------------------

#사용자 정보
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

# 종목 코드
kospi = kiwoom.GetCodeListByMarket('0')
kosdaq = kiwoom.GetCodeListByMarket('10')
etf = kiwoom.GetCodeListByMarket('8')

print(len(kospi), kospi)
print(len(kosdaq), kosdaq)
print(len(etf), etf)

#종목명
name = kiwoom.GetMasterCodeName("005930")
print(name)

# 연결 상태 확인
state = kiwoom.GetConnectState()
if state == 0:
    print("미연결")
elif state == 1:
    print("연결완료")

#상장 주식 수
stock_cnt = kiwoom.GetMasterListedStockCnt("005930")
print("삼성전자 상장주식수: ", stock_cnt)

# 감리구분
감리구분 = kiwoom.GetMasterConstruction("005930")
print(감리구분)

# 상장일
상장일 = kiwoom.GetMasterListedStockDate("005930")
print(상장일)
print(type(상장일))   

# 전일가
전일가 = kiwoom.GetMasterLastPrice("005930")
print(int(전일가))
print(type(전일가))

# 종목상태
종목상태 = kiwoom.GetMasterStockState("005930")
print(종목상태)

# 테마
group = kiwoom.GetThemeGroupList(1)
pprint.pprint(group)


# 4. 조건검색 -------------------------------------------------

#조건검색 조회
# 조건식을 PC로 다운로드
kiwoom.GetConditionLoad()

# 전체 조건식 리스트 얻기
conditions = kiwoom.GetConditionNameList()

# 0번 조건식에 해당하는 종목 리스트 출력
condition_index = conditions[0][0]
condition_name = conditions[0][1]
codes = kiwoom.SendCondition("0101", condition_name, condition_index, 0)

print(codes) #IndexError: list index out of range

#5. 매매 -------------------------------------------------

# 주식계좌
#로그인에러 어케 해결하는 거징
accounts = kiwoom.GetLoginInfo("ACCNO")
stock_account = accounts[0]

# 삼성전자, 10주, 시장가주문 매수
kiwoom.SendOrder("시장가매수", "0101", stock_account, 1, "005930", 10, 0, "03", "")

#매도
kiwoom.SendOrder("시장가매도", "0101", stock_account, 2, "005930", 10, 0, "03", "")
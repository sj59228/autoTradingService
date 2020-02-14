from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()  # 부모 클래스의 함수 상속

        print('Kikooom 클래스입니다.')

        ##################### 이벤트 루프 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        #####################################

        ##################### 변수 모음
        self.screen_my_info = "2000"
        #####################################


        ##################### 변수 모음
        self.account_num = None
        self.account_stock_dict = {}        # 현재 소유한 주식 목록
        self.not_account_stock_dict = {}    # 아직 미체결된 주식 목록
        #####################################

        ##################### 계좌 관련 변수
        self.use_money = 0
        self.use_money_percent = 0.5
        #####################################



        self.get_ocx_instance()
        self.event_slots()


        self.signal_login_CommConnect() #(시그널 보냄)
        self.get_account_info()
        self.detail_account_info()     # 예수금을 가져오는 함수(시그널 보냄)
        self.detail_account_mystock()  # 계좌평가 잔고내역 요청하는 함수(시그널 보냄)
        self.not_concluded_account()   # 미체결 요청(시그널 보냄)


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOPENAPICtrl.1") # 키움 open Api 사용 가능

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)    # 로그인 반환 값은 받는 곳
        self.OnReceiveTrData.connect(self.trdata_slot)  # tr 데이터를 받는 곳

    def login_slot(self, errCode):
        print(errors(errCode))                          # 0이면 성공
        self.login_event_loop.exit()            # 이벤트 루프 종료

    def signal_login_CommConnect(self):
        self.dynamicCall("CommConnect()")   # dynamicCall 다른 네트워크 서버에 데이터 전송(PyQt5 에서 제공)
                                            # 키움 API에 CommConnect함수로 시그널을 보냄

        ## 로그인이 완료 될때 까지 루프 시작 / QEventLoop (Qtcore 에서 제공)
        self.login_event_loop = QEventLoop()    # 시그널을 보낸 후 이벤트 루프 생성
        self.login_event_loop.exec_()           # 다음 것이 실행 안되게 막는 것


    def get_account_info(self):
        # 인자 형식을 String으로 지정해주고 옆에 해당 값을 넣어줌
        # 하나씩 밖에 못 넣음
        # return : 은 ; 기준으로 반환나옴
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]   # ; 별로 끊기
        print("나의 보유 계좌번호 %s" % self.account_num)   # 8130311311


    def detail_account_info(self):
        print('예수금 요청하는 부분')
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "")   # 공백으로 해도 됨
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)

        self.detail_account_info_event_loop = QEventLoop()  # 데이터를 요청하고 event loop에 들어간다
        self.detail_account_info_event_loop.exec_()         # 다음 코드를 막음

    def detail_account_mystock(self, sPrevNext="0"):
        print('계좌평가 잔고내역 요청하기 연속조회 %s' % sPrevNext)

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")  # 공백으로 해도 됨
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)


        # 전역 변수에서
        # self.detail_account_info_event_loop = QEventLoop() # 데이터를 요청하고 event loop에 들어간다
        # 먼저 선언해줌
        self.detail_account_info_event_loop.exec_()  # 다음 코드를 막음



    def not_concluded_account(self, sPrevNext="0"):
        print("미체결 요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결부분", "1")  # 공백으로 해도 됨
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()  # 다음 코드를 막음


    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr 요청한 데이터를 받는 구역
        :param sScrNo:스크린 번호
        :param sRQName:내가 요청했을 때 지은 이름
        :param sTrCode:요청 id, tr코드
        :param sRecordName:사용 안함
        :param sPrevNext:다음 페이지가 있는지
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print( "예수금 : %s" % int(deposit))

            self.use_money= int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4


            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            int(ok_deposit)
            print("출금예수금 : %s" % int(ok_deposit))

            self.detail_account_info_event_loop.exit()  # 이벤트 loop 종료



        if sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)
            print( "총매입금액 : %s" % total_buy_money_result)


            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)
            print('총수익률(%%) : %s' % total_profit_loss_rate_result)

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            # GetRepeatCnt 멀티데이터 조회 용도 (총 20개까지만 조회 가능)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]
                # A 장내주식 / B ELW 종목 / Q ETN 종목
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")

                if code in self.self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({code : {}})

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명" : code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt += 1    # 총 종목 개수

            print ("계좌에 가지고 있는 종목 : %s" % self.account_stock_dict)
            print ("계좌에 보유종목 카운트 : %s" % cnt)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()  # 이벤트 loop_2 종료


        if sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태") # 접수 -> 확인 -> 체결
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분") # -매도, +매수, 매도정
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,  "체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}


                nasd = self.not_account_stock_dict[order_no]
                nasd.update ({"종목코드" : code})
                nasd.update({"종목명": code_nm})
                nasd.update({"주문번호": order_no})
                nasd.update({"주문상태": order_status})
                nasd.update({"주문수량": order_quantity})
                nasd.update({"주문가격": order_price})
                nasd.update({"주문구분": order_gubun})
                nasd.update({"미체결수량": not_quantity})
                nasd.update({"체결량": ok_quantity})

                print("미체결 종목 : %s" % self.not_account_stock_dict[order_no])


            self.detail_account_info_event_loop.exit()

















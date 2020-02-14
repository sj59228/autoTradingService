from kiwoom.kiwoom import *
from  PyQt5.QtWidgets import *

import sys

class Ui_class():
    def __init__(self):
        print('Ui 클래스 입니다.')
        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom()
        self.app.exec_()        # 다음 것이 실행 안되게 막는 것 / 종료를 막음


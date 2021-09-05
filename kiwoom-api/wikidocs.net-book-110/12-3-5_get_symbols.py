# 종목 코드 및 한글 종목명을 출력하는 프로그램
# https://wikidocs.net/4244
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import  *
from PyQt5.QAxContainer import *
from PyQt5 import QtWidgets

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        # OpenAPI+ Event
        self.kiwoom.OnEventConnect.connect(self.on_event_connect)
        self.kiwoom.dynamicCall("CommConnect()")

        self.setWindowTitle("종목 코드")
        self.setGeometry(300, 300, 300, 300)

        btn1 = QPushButton("종목코드 얻기", self)
        btn1.move(190, 10)
        btn1.clicked.connect(self.btn1_clicked)

        self.combo = QComboBox(self)
        self.combo.addItem('0: 장내', '0')
        self.combo.addItem('10: 코스닥', '10')
        self.combo.move(190, 45)

        self.listWidget = QListWidget(self)
        self.listWidget.setGeometry(10, 10, 170, 280)

    def message_box(self, text, icon=QtWidgets.QMessageBox.Information):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(icon)
        msg.setText(text)
        msg.exec()

    def btn1_clicked(self):
        if self.kiwoom.dynamicCall("GetConnectState()") == 0:
            self.message_box('서버 연결이 안되었습니다.')
            return
        self.listWidget.clear()
        # 전체 심볼을 얻어온다
        get_type = self.combo.currentData()
        print(f'get_type: {get_type}')
        ret = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", [get_type])
        #print(f'GetCodeListByMarket: {ret}')
        kospi_code_list = ret.split(';')
        kospi_code_name_list = []
        ## 여기에 조회된 종목 코드는 거래정지 종목도 포함되어 있음으로
        ## 걸러내는 기능이 필요함.
        for x in kospi_code_list:
            if x:
                name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [x])
                print(f'{x},{name}')
                kospi_code_name_list.append(x + " : " + name)
        self.listWidget.addItems(kospi_code_name_list)
        print(f'get_type={get_type}, num_symbols={len(kospi_code_name_list)}')

    def on_event_connect(self, err_code):
        """
        - 연결성공하면 0
        - 연결도중에 끊긴 경우 화면에 메시지 박스 띄우고 오류코드 -106으로 이벤트가 발생됨
        - CommConnect() 호출후 아예 접속이 안되는 경우 이벤트 발생되지 않음
        """
        print(f'err_code={err_code}')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(app.exec_())
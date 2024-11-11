import sys
from PyQt5.QtWidgets import QTextEdit, QWidget, QApplication, QPushButton, QLabel, QLineEdit, QGridLayout, QVBoxLayout
from PyQt5.QAxContainer import QAxWidget

import pandas_datareader.data as web
import pandas as pd

import datetime

import math

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")

        self.kiwoom.OnEventConnect.connect(self.event_connect)


        self.setupUI()

        self.money = 0

        self.f_PER = ""
        self.s_PER = ""
        self.t_PER = ""

        self.ff_PER = 0
        self.fs_PER = 0
        self.ft_PER = 0

        self.f_w1 = 0
        self.f_w2 = 0
        self.f_w3 = 0
        self.s_w1 = 0
        self.s_w2 = 0
        self.s_w3 = 0
        self.t_w1 = 0
        self.t_w2 = 0
        self.t_w3 = 0

        self.f_e = 0
        self.f_r = 0
        self.s_e = 0
        self.s_r = 0
        self.t_e = 0
        self.t_r = 0


        self.f_pe = 0
        self.f_pr = 0
        self.s_pe = 0
        self.s_pr = 0
        self.t_pe = 0
        self.t_pr = 0

        self.invest1 = ""
        self.invest2 = ""
        self.invest3 = ""

        self.c1 = 0
        self.c2 = 0
        self.c3 = 0


        self.start = datetime.datetime(2019, 1, 1)
        self.end = datetime.datetime(2020, 1, 1)

    def setupUI(self):
        self.setWindowTitle("주식포트폴리오")
        self.setGeometry(300,300,400,200)

        self.label1 = QLabel("자산: ", self)
        self.label2 = QLabel("투자항목(종목코드): ", self)
        self.label3_0 = QLabel("1번투자비율(합100): ", self)
        self.label3_1 = QLabel("2번투자비율(합100): ", self)
        self.label3_2 = QLabel("3번투자비율(합100): ", self)

        self.label4_0 = QLabel("1번투자비율", self)
        self.label4_1 = QLabel("2번투자비율", self)
        self.label4_2 = QLabel("3번투자비율", self)
        self.label5 = QLabel("포트폴리오기대수익률(%): ", self)
        self.label6 = QLabel("포트폴리오기대수익: ", self)
        self.label7 = QLabel("포트폴리오위험(%): ", self)



        self.lineEdit1 = QLineEdit("", self)
        self.lineEdit2 = QLineEdit("", self)
        self.lineEdit3 = QLineEdit("", self)
        self.lineEdit4 = QLineEdit("", self)
        self.lineEdit5_1 = QLineEdit("", self)
        self.lineEdit5_2 = QLineEdit("", self)
        self.lineEdit5_3 = QLineEdit("", self)
        self.lineEdit6_1 = QLineEdit("", self)
        self.lineEdit6_2 = QLineEdit("", self)
        self.lineEdit6_3 = QLineEdit("", self)
        self.lineEdit7_1 = QLineEdit("", self)
        self.lineEdit7_2 = QLineEdit("", self)
        self.lineEdit7_3 = QLineEdit("", self)

        self.btn1 = QPushButton("기대수익률", self)
        self.btn2 = QPushButton("기대수익률", self)
        self.btn3 = QPushButton("기대수익률", self)

        self.btn4 = QPushButton("확인", self)

        self.btn5 = QPushButton("위험", self)
        self.btn6 = QPushButton("위험", self)
        self.btn7 = QPushButton("위험", self)


        self.lineEdit8_1 = QLineEdit("", self)
        self.lineEdit8_2 = QLineEdit("", self)
        self.lineEdit8_3 = QLineEdit("", self)
        self.lineEdit9_1 = QLineEdit("", self)
        self.lineEdit9_2 = QLineEdit("", self)
        self.lineEdit9_3 = QLineEdit("", self)
        self.lineEdit92_1 = QLineEdit("", self)
        self.lineEdit92_2 = QLineEdit("", self)
        self.lineEdit92_3 = QLineEdit("", self)

        self.label10 = QLabel("기대수익률(%): ", self)
        self.lineEdit10_1 = QLineEdit("", self)
        self.lineEdit10_2 = QLineEdit("", self)
        self.lineEdit10_3 = QLineEdit("", self)

        self.label11 = QLabel("위험(%): ", self)
        self.lineEdit11_1 = QLineEdit("", self)
        self.lineEdit11_2 = QLineEdit("", self)
        self.lineEdit11_3 = QLineEdit("", self)




        layout = QVBoxLayout()

        inputPad = QGridLayout()

        inputPad.addWidget(self.label1, 0, 0)
        inputPad.addWidget(self.lineEdit1, 0, 1)

        inputPad.addWidget(self.label2, 1, 0)
        inputPad.addWidget(self.lineEdit2, 1, 1)
        inputPad.addWidget(self.lineEdit3, 1, 2)
        inputPad.addWidget(self.lineEdit4, 1, 3)

        inputPad.addWidget(self.btn1, 2, 1)
        inputPad.addWidget(self.btn2, 2, 2)
        inputPad.addWidget(self.btn3, 2, 3)

        inputPad.addWidget(self.btn5, 3, 1)
        inputPad.addWidget(self.btn6, 3, 2)
        inputPad.addWidget(self.btn7, 3, 3)


        inputPad.addWidget(self.label10, 4, 0)
        inputPad.addWidget(self.lineEdit10_1, 4, 1)
        inputPad.addWidget(self.lineEdit10_2, 4, 2)
        inputPad.addWidget(self.lineEdit10_3, 4, 3)

        inputPad.addWidget(self.label11, 5, 0)
        inputPad.addWidget(self.lineEdit11_1, 5, 1)
        inputPad.addWidget(self.lineEdit11_2, 5, 2)
        inputPad.addWidget(self.lineEdit11_3, 5, 3)

        inputPad.addWidget(self.label3_0, 6, 0)
        inputPad.addWidget(self.lineEdit5_1, 6, 1)
        inputPad.addWidget(self.lineEdit5_2, 6, 2)
        inputPad.addWidget(self.lineEdit5_3, 6, 3)

        inputPad.addWidget(self.label3_1, 7, 0)
        inputPad.addWidget(self.lineEdit6_1, 7, 1)
        inputPad.addWidget(self.lineEdit6_2, 7, 2)
        inputPad.addWidget(self.lineEdit6_3, 7, 3)

        inputPad.addWidget(self.label3_2, 8, 0)
        inputPad.addWidget(self.lineEdit7_1, 8, 1)
        inputPad.addWidget(self.lineEdit7_2, 8, 2)
        inputPad.addWidget(self.lineEdit7_3, 8, 3)

        inputPad.addWidget(self.btn4, 8, 4)

        outputPad = QGridLayout()
        outputPad.addWidget(self.label4_0, 0, 1)
        outputPad.addWidget(self.label4_1, 0, 2)
        outputPad.addWidget(self.label4_2, 0, 3)

        outputPad.addWidget(self.label5, 1, 0)
        outputPad.addWidget(self.lineEdit8_1, 1, 1)
        outputPad.addWidget(self.lineEdit8_2, 1, 2)
        outputPad.addWidget(self.lineEdit8_3, 1, 3)

        outputPad.addWidget(self.label6, 2, 0)
        outputPad.addWidget(self.lineEdit9_1, 2, 1)
        outputPad.addWidget(self.lineEdit9_2, 2, 2)
        outputPad.addWidget(self.lineEdit9_3, 2, 3)

        outputPad.addWidget(self.label7, 3, 0)
        outputPad.addWidget(self.lineEdit92_1, 3, 1)
        outputPad.addWidget(self.lineEdit92_2, 3, 2)
        outputPad.addWidget(self.lineEdit92_3, 3, 3)

        self.text_edit = QTextEdit(self)
        self.text_edit.setEnabled(False)


        layout.addLayout(inputPad)
        layout.addLayout(outputPad)
        layout.addWidget(self.text_edit)
        self.setLayout(layout)


        self.btn1.clicked.connect(self.btn1Click)
        self.btn2.clicked.connect(self.btn2Click)
        self.btn3.clicked.connect(self.btn3Click)
        self.btn4.clicked.connect(self.btn4Click)

        self.btn5.clicked.connect(self.btn5Click)
        self.btn6.clicked.connect(self.btn6Click)
        self.btn7.clicked.connect(self.btn7Click)

    def event_connect(self, err_code):
        if err_code == 0:
            self.text_edit.append("로그인 성공")




    def btn1Click(self):
        self.invest1 = self.lineEdit2.text()

        self.kiwoom.dynamicCall("SetInputValue(S, S)", "종목코드", self.invest1)
        self.kiwoom.dynamicCall("CommRqData(S, S, int, S)", "opt10001_req", "opt10001", 0, "0600")

        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata1)


    def btn2Click(self):
        self.invest2 = self.lineEdit3.text()

        self.kiwoom.dynamicCall("SetInputValue(S, S)", "종목코드", self.invest2)
        self.kiwoom.dynamicCall("CommRqData(S, S, int, S)", "opt10001_req", "opt10001", 0, "0600")

        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata2)


    def btn3Click(self):
        self.invest3 = self.lineEdit4.text()

        self.kiwoom.dynamicCall("SetInputValue(S, S)", "종목코드", self.invest3)
        self.kiwoom.dynamicCall("CommRqData(S, S, int, S)", "opt10001_req", "opt10001", 0, "0600")

        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata3)




    def btn5Click(self):
        weekly_close1 = pd.DataFrame()

        self.invest1 = self.lineEdit2.text() + ".KS"

        fp = web.DataReader(self.invest1, "yahoo", self.start, self.end)

        close = fp['Adj Close']

        weekly_close1 = close.resample('W').mean()

        self.pct1 = weekly_close1.pct_change()
        self.f_r = round(self.pct1.std() * 100, 2)

        self.lineEdit11_1.setText(str(self.f_r) + "%")

    def btn6Click(self):
        weekly_close2 = pd.DataFrame()
        self.invest2 = self.lineEdit3.text() + ".KS"

        sp = web.DataReader(self.invest2, "yahoo", self.start, self.end)

        close = sp['Adj Close']

        weekly_close2 = close.resample('W').mean()

        self.pct2 = weekly_close2.pct_change()
        self.s_r = round(self.pct2.std() * 100, 2)

        self.lineEdit11_2.setText(str(self.s_r) + "%")

    def btn7Click(self):
        weekly_close3 = pd.DataFrame()
        self.invest3 = self.lineEdit4.text() + ".KS"

        tp = web.DataReader(self.invest3, "yahoo", self.start, self.end)

        close = tp['Adj Close']

        weekly_close3 = close.resample('W').mean()

        self.pct3 = weekly_close3.pct_change()
        self.t_r = round(self.pct3.std() * 100, 2)

        self.lineEdit11_3.setText(str(self.t_r) + "%")











    def btn4Click(self):
        self.money = float(self.lineEdit1.text())

        self.f_w1 = float(self.lineEdit5_1.text())
        self.f_w2 = float(self.lineEdit5_2.text())
        self.f_w3 = float(self.lineEdit5_3.text())
        self.s_w1 = float(self.lineEdit6_1.text())
        self.s_w2 = float(self.lineEdit6_2.text())
        self.s_w3 = float(self.lineEdit6_3.text())
        self.t_w1 = float(self.lineEdit7_1.text())
        self.t_w2 = float(self.lineEdit7_2.text())
        self.t_w3 = float(self.lineEdit7_3.text())


        self.f_pe = round(((self.f_e * self.f_w1) + (self.s_e * self.f_w2) +(self.t_e + self.f_w3)) / 100, 2)
        self.s_pe = round(((self.f_e * self.s_w1) + (self.s_e * self.s_w2) + (self.t_e + self.s_w3)) / 100, 2)
        self.t_pe = round(((self.f_e * self.t_w1) + (self.s_e * self.t_w2) + (self.t_e + self.t_w3)) / 100, 2)

        self.lineEdit8_1.setText(str(self.f_pe) + "%")
        self.lineEdit8_2.setText(str(self.s_pe) + "%")
        self.lineEdit8_3.setText(str(self.t_pe) + "%")

        self.lineEdit9_1.setText(str(round((self.money * self.f_pe / 100))))
        self.lineEdit9_2.setText(str(round((self.money * self.s_pe / 100))))
        self.lineEdit9_3.setText(str(round((self.money * self.t_pe / 100))))



        df1 = pd.merge(self.pct1, self.pct2, on = "Date")
        c1 = df1.corr(method='pearson')
        self.c1 = c1.iat[0, 1]

        df2 = pd.merge(self.pct1, self.pct3, on = "Date")
        c2 = df2.corr(method='pearson')
        self.c2 = c2.iat[0, 1]

        df3 = pd.merge(self.pct2, self.pct3, on = "Date")
        c3 = df3.corr(method='pearson')
        self.c3 = c3.iat[0, 1]

        self.f_pr = round(math.sqrt((self.f_w1**2 * self.f_r**2) + (self.f_w2**2 * self.s_r**2) + (self.f_w3**2 * self.t_r**2) +
                     (2 * ((self.f_w1 * self.f_w2 * self.c1 * self.f_r * self.s_r) + (self.f_w1 * self.f_w3 * self.c2 * self.f_r * self.t_r)
                           + (self.f_w2 * self.f_w3 * self.c3 * self.s_r * self.t_r)))) / 100 , 2)

        self.s_pr = round(math.sqrt((self.s_w1**2 * self.f_r**2) + (self.s_w2**2 * self.s_r**2) + (self.s_w3**2 * self.t_r**2) +
                     (2 * ((self.s_w1 * self.s_w2 * self.c1 * self.f_r * self.s_r) + (self.s_w1 * self.s_w3 * self.c2 * self.f_r * self.t_r)
                           + (self.s_w2 * self.s_w3 * self.c3 * self.s_r * self.t_r)))) / 100 , 2)

        self.t_pr = round(math.sqrt((self.t_w1**2 * self.f_r**2) + (self.t_w2**2 * self.s_r**2) + (self.t_w3**2 * self.t_r**2) +
                     (2 * ((self.t_w1 * self.t_w2 * self.c1 * self.f_r * self.s_r) + (self.t_w1 * self.t_w3 * self.c2 * self.f_r * self.t_r)
                           + (self.t_w2 * self.t_w3 * self.c3 * self.s_r * self.t_r)))) / 100 , 2)



        self.lineEdit92_1.setText(str(self.f_pr) + '%')
        self.lineEdit92_2.setText(str(self.s_pr) + '%')
        self.lineEdit92_3.setText(str(self.t_pr) + '%')





    def receive_trdata1(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "opt10001_req":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",rqname, 0, "종목코드")

            if name.strip() == self.invest1:

                self.f_PER = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",rqname, 0, "PER")

                self.ff_PER = float(self.f_PER)
                self.f_e = round((1 / self.ff_PER) * 100, 2)

                self.lineEdit10_1.setText(str(self.f_e) + "%")

    def receive_trdata2(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "opt10001_req":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목코드")

            if name.strip() == self.invest2:

                self.s_PER = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",rqname, 0, "PER")

                self.fs_PER = float(self.s_PER)
                self.s_e = round((1 / self.fs_PER) * 100, 2)

                self.lineEdit10_2.setText(str(self.s_e) + "%")

    def receive_trdata3(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "opt10001_req":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목코드")

            if name.strip() == self.invest3:

                self.t_PER = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",rqname, 0, "PER")

                self.ft_PER = float(self.t_PER)
                self.t_e = round((1 / self.ft_PER) * 100, 2)

                self.lineEdit10_3.setText(str(self.t_e) + "%")












app = QApplication(sys.argv)
mymy = MyWindow()
mymy.show()
app.exec_()
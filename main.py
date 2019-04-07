import sys
from PyQt5.QtWidgets import QWidget, QPushButton, QDesktopWidget, QApplication, QLabel, QFileDialog, QCheckBox, QMessageBox
from PyQt5.QtCore import QCoreApplication, Qt
from PyQt5.QtGui import QColor, QPainter
import reports
import os
import mysql.connector
import export


class ImportWindow(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.path = ''
        self.table = 0
        self.branch = 0
        self.parent = parent
        self.initUI()
        self.center()

    def initUI(self):
        dir_lbl = QLabel(self)
        dir_lbl.setText('Выберите путь до файла с данными')
        dir_lbl.setGeometry(20, 20, 200, 20)

        self.path_lbl = QLabel(self)
        self.path_lbl.setText('')
        self.path_lbl.setGeometry(20, 40, 280, 20)

        self.file_btn = QPushButton('Обзор', self)
        self.file_btn.clicked.connect(self.open_dir)
        self.file_btn.resize(self.file_btn.sizeHint())
        self.file_btn.move(305, 39)
        self.file_btn.setEnabled(False)

        dir_lbl = QLabel(self)
        dir_lbl.setText('Выберите филиал')
        dir_lbl.setGeometry(20, 70, 200, 20)

        self.check1 = QCheckBox('Филиал №1 (Excel)', self)
        self.check1.stateChanged.connect(self.setBranch)
        self.check1.move(20,120)

        self.check2 = QCheckBox('Филиал №2 (Access)', self)
        self.check2.stateChanged.connect(self.setBranch)
        self.check2.move(150, 120)

        dir_lbl = QLabel(self)
        dir_lbl.setText('Выберите таблицу в БД')
        dir_lbl.setGeometry(20, 170, 200, 20)

        self.check3 = QCheckBox('Поставщики', self)
        self.check3.stateChanged.connect(self.setTable)
        self.check3.move(20, 200)

        self.check4 = QCheckBox('Детали', self)
        self.check4.stateChanged.connect(self.setTable)
        self.check4.move(150, 200)

        self.check5 = QCheckBox('Поставки', self)
        self.check5.move(250, 200)
        self.check5.stateChanged.connect(self.setTable)

        self.pr_btn = QPushButton('Внести данные в базу', self)
        self.pr_btn.clicked.connect(self.process)
        self.pr_btn.setGeometry(30, 240, 180, 40)
        self.pr_btn.setEnabled(False)

        qbtn = QPushButton('Вернуться', self)
        qbtn.clicked.connect(self.close)
        qbtn.resize(140, 40)
        qbtn.move(250, 240)

        self.setGeometry(400, 400, 400, 300)
        self.setWindowTitle('Импорт данных')
        self.show()

    def close(self):
        self.hide()
        self.parent.show()

    def setBranch(self, state):
        source = self.sender()
        if state == Qt.Checked:
            if source.text() == "Филиал №1 (Excel)":
                self.branch = 1
                self.check2.setCheckState(Qt.Unchecked)
            elif source.text() == "Филиал №2 (Access)":
                self.branch = 2
                self.check1.setCheckState(Qt.Unchecked)
            self.file_btn.setEnabled(True)

    def setTable(self, state):
        source = self.sender()
        if state == Qt.Checked:
            if source.text() == "Поставщики":
                self.table = 1
                self.check4.setCheckState(Qt.Unchecked)
                self.check5.setCheckState(Qt.Unchecked)
            elif source.text() == "Детали":
                self.table = 2
                self.check3.setCheckState(Qt.Unchecked)
                self.check5.setCheckState(Qt.Unchecked)
            elif source.text() == "Поставки":
                self.table = 3
                self.check3.setCheckState(Qt.Unchecked)
                self.check4.setCheckState(Qt.Unchecked)
            if self.path:
                self.pr_btn.setEnabled(True)

    def open_dir(self):
        if self.branch == 1:
            self.path = QFileDialog.getOpenFileName(self, "Выберите файл с данными", "", "Файлы Excel (*.xlsx)")[0]
        else:
            self.path = QFileDialog.getOpenFileName(self, "Выберите файл с данными", "", "Файлы MS Access (*.accdb)")[0]
        if self.path:
            self.path_lbl.setText(self.path)
            if self.table != 0:
                self.pr_btn.setEnabled(True)

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def paintEvent(self, e):
        qp = QPainter()
        qp.begin(self)
        col = QColor(0, 0, 0)
        col.setNamedColor('#d4d4d4')
        qp.setPen(col)
        qp.setBrush(QColor(255, 255, 255))
        qp.drawRect(18, 40, 280, 20)
        qp.end()

    def process(self):
        try:
            if self.table == 1:
                import ETL_Suppliers
                mess = ETL_Suppliers.fillSuppliersTable(self.path, self.branch)
                QMessageBox.warning(self, 'Info', mess, QMessageBox.Ok)
            elif self.table ==2:
                import ETL_Parts
                mess = ETL_Parts.fillPartsTable(self.path, self.branch)
                QMessageBox.warning(self, 'Info', mess, QMessageBox.Ok)
            else:
                import ETL_Deliveries
                mess = ETL_Deliveries.fillDeliveriesTable(self.path, self.branch)
                QMessageBox.warning(self, 'Info', mess, QMessageBox.Ok)
        except mysql.connector.errors.InterfaceError:
            QMessageBox.warning(self, 'Warning', 'Невозможно подключиться к БД', QMessageBox.Ok)


class ExportWindow(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.initUI()
        self.center()

    def initUI(self):
        dir_lbl = QLabel(self)
        dir_lbl.setText('Выберите данные для экспорта')
        dir_lbl.setGeometry(20, 20, 200, 20)

        exp_btn = QPushButton('Данные о поставках', self)
        exp_btn.clicked.connect(self.process)
        exp_btn.setGeometry(20, 50, 150, 35)

        supp_btn = QPushButton('Данные о поставщиках', self)
        supp_btn.setGeometry(200, 50, 150, 35)

        qbtn = QPushButton('Вернуться', self)
        qbtn.clicked.connect(self.close)
        qbtn.resize(140, 25)
        qbtn.move(20, 120)

        self.setGeometry(400, 400, 400, 160)
        self.setWindowTitle('Экспорт данных')
        self.show()

    def close(self):
        self.hide()
        self.parent.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def process(self):
        try:
            path = export.export_deliveries_by_dates()
            os.startfile(os.curdir + '//' + path)
            QMessageBox.warning(self, 'Info', 'Выгрузку данных о поставках можно найти по адресу  ' + path,
                                QMessageBox.Ok)
        except mysql.connector.errors.InterfaceError:
            QMessageBox.warning(self, 'Warning', 'Невозможно подключиться к БД', QMessageBox.Ok)



class ClientWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.center()

    def initUI(self):
        lbl = QLabel(self)
        lbl.setText('Выберите интересующий вас отчет')
        lbl.setGeometry(140, 10, 200, 20)

        r1_btn = QPushButton('Вес поставок в зависимости от времени и поставщика', self)
        r1_btn.setGeometry(80, 40, 300, 40)
        r1_btn.clicked.connect(self.report_weight_city)

        r2_btn = QPushButton('Стоимость поставок в зависимости от времени и поставщика', self)
        r2_btn.setGeometry(65, 95, 330, 40)
        r2_btn.clicked.connect(self.report_price_city)

        r3_btn = QPushButton('Стоимость поставок в зависимости от времени и весовой категории поставки', self)
        r3_btn.setGeometry(15, 150, 420, 40)
        r3_btn.clicked.connect(self.report_price_weight)

        r4_btn = QPushButton('Вес поставок в зависимости от времени и ценовой категории детали', self)
        r4_btn.setGeometry(40, 205, 380, 40)
        r4_btn.clicked.connect(self.peport_weight_price)

        pr_btn = QPushButton('Импорт данных', self)
        pr_btn.clicked.connect(self.import_data)
        pr_btn.setGeometry(30, 290, 140, 40)

        qbtn = QPushButton('Экспорт данных', self)
        qbtn.clicked.connect(self.export_data)
        qbtn.resize(140, 40)
        qbtn.move(280, 290)

        self.setGeometry(400, 400, 450, 350)
        self.setWindowTitle('Клиент')
        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def paintEvent(self, e):
        qp = QPainter()
        qp.begin(self)
        col = QColor(0, 0, 0)
        col.setNamedColor('#d4d4d4')
        qp.setPen(col)
        qp.setBrush(QColor(0, 0, 0))
        qp.drawRect(225, 270, 2, 120)
        qp.drawRect(0, 270, 450, 2)
        qp.end()

    def import_data(self):
        self.import_window = ImportWindow(self)
        self.import_window.show()
        self.hide()

    def report_weight_city(self):
        try:
            path = reports.report_weight_city()
            os.startfile(os.curdir + '//' + path)
            QMessageBox.warning(self, 'Info', 'Отчет успешно сгенерирован. Вы можете найти его по адресу ' + path,
                                QMessageBox.Ok)
        except mysql.connector.errors.InterfaceError:
            QMessageBox.warning(self, 'Warning', 'Невозможно подключиться к БД', QMessageBox.Ok)

    def peport_weight_price(self):
        try:
            path = reports.report_weight_price()
            os.startfile(os.curdir + '//' + path)
            QMessageBox.warning(self, 'Info', 'Отчет успешно сгенерирован. Вы можете найти его по адресу ' + path,
                                QMessageBox.Ok)
        except mysql.connector.errors.InterfaceError:
            QMessageBox.warning(self, 'Warning', 'Невозможно подключиться к БД', QMessageBox.Ok)

    def report_price_weight(self):
        try:
            path = reports.report_price_weight()
            os.startfile(os.curdir + '//' + path)
            QMessageBox.warning(self, 'Info', 'Отчет успешно сгенерирован. Вы можете найти его по адресу ' + path,
                                QMessageBox.Ok)
        except mysql.connector.errors.InterfaceError:
            QMessageBox.warning(self, 'Warning', 'Невозможно подключиться к БД', QMessageBox.Ok)

    def report_price_city(self):
        try:
            path = reports.report_price_city()
            os.startfile(os.curdir + '//' + path)
            QMessageBox.warning(self, 'Info', 'Отчет успешно сгенерирован. Вы можете найти его по адресу ' + path,
                                QMessageBox.Ok)
        except mysql.connector.errors.InterfaceError:
            QMessageBox.warning(self, 'Warning', 'Невозможно подключиться к БД', QMessageBox.Ok)

    def export_data(self):
        self.export_window = ExportWindow(self)
        self.export_window.show()
        self.hide()


app = QApplication(sys.argv)
ex = ClientWindow()
sys.exit(app.exec_())

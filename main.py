import sys
from PyQt5.QtWidgets import QWidget, QPushButton, QDesktopWidget, QApplication, QLabel, QFileDialog, QCheckBox, QMessageBox
from PyQt5.QtCore import QCoreApplication, Qt
from PyQt5.QtGui import QColor, QPainter


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.path = ''
        self.table = 0
        self.branch = 0
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

        qbtn = QPushButton('Завершить работу', self)
        qbtn.clicked.connect(QCoreApplication.instance().quit)
        qbtn.resize(140, 40)
        qbtn.move(250, 240)

        self.setGeometry(400, 400, 400, 300)
        self.setWindowTitle('Клиент')
        self.show()

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
        import mysql.connector
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


app = QApplication(sys.argv)
ex = Window()
sys.exit(app.exec_())
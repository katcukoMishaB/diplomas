from PyQt5 import QtCore, QtGui, QtWidgets
from functools import partial
from ui_work_monitoring import Ui_WorkMonitoring
from sqlalchemy import create_engine, Table, Column, String, MetaData
from PyQt5.QtWidgets import QHeaderView


class WindowWorkMonitoring(QtWidgets.QMainWindow):
    def __init__(self):
        super(WindowWorkMonitoring, self).__init__()
        self.ui = Ui_WorkMonitoring()
        self.ui.setupUi(self)
        self.databased()
        self.ui.pushButton.clicked.connect(self.exit)
        self.ui.pushButton_2.clicked.connect(self.path_diploma)
        self.ui.pushButton_3.clicked.connect(self.path_sertificate)

        self.ui.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.tableWidget_2.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        self.ui.tableWidget.setRowCount(len(self.result_diploma))
        self.ui.tableWidget.setColumnCount(len(self.result_diploma[0]))
        for i, row in enumerate(self.result_diploma):
            for j, cell in enumerate(row):
                item = QtWidgets.QTableWidgetItem(str(cell))
                self.ui.tableWidget.setItem(i, j, item)
        self.ui.tableWidget_2.setRowCount(len(self.result_sertificate))
        self.ui.tableWidget_2.setColumnCount(len(self.result_sertificate[0]))
        for i, row in enumerate(self.result_sertificate):
            for j, cell in enumerate(row):
                item = QtWidgets.QTableWidgetItem(str(cell))
                self.ui.tableWidget_2.setItem(i, j, item)
        self.ui.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.ui.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.ui.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.ui.tableWidget_2.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    def databased(self):
        engine = create_engine('sqlite:///diploma.db', echo=True)
        metadata = MetaData()
        diploma = Table('diploma', metadata,
                                Column('last_name', String),
                                Column('first_name', String),
                                Column('patronymic', String),
                                Column('place', String),
                                Column('email', String))
        sertificate = Table('sertificate', metadata,
                                Column('second_name', String),
                                Column('first_name', String),
                                Column('patronymic', String),
                                Column('email', String))
        connection = engine.connect()
        self.result_diploma = connection.execute(diploma.select()).fetchall()
        self.result_sertificate = connection.execute(sertificate.select()).fetchall()
        connection.close()
        
    def exit(self):
        from main_window import MainWindow
        self.main = MainWindow()
        self.main.show()
        self.close()

    def path_diploma(self):
        from diploma import WindowDiploma
        self.window_diploma = WindowDiploma()
        self.window_diploma.show()
        self.close()
    def path_sertificate(self):
        from sertificate import WindowSertificate
        self.window_sertificate = WindowSertificate()
        self.window_sertificate.show()
        self.close()
    
        
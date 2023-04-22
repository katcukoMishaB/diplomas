from PyQt5 import QtWidgets, QtCore
from ui_main_window import Ui_MainWindow
from diploma import Window3
from letter import Window2

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.create_win_path_to_diploma)
        self.ui.pushButton_2.clicked.connect(self.create_win_path_to_letter)
    def create_win_path_to_diploma(self):
        from diploma import Window3
        self.window2 = Window3()
        self.window2.show()
    def create_win_path_to_letter(self):
        self.window3 = Window2()
        self.window3.show() 
       

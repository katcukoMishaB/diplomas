from PyQt5 import QtWidgets
from ui_mainwindow import Ui_MainWindow


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton_7.clicked.connect(self.create_win_path_to_diploma)
        self.ui.pushButton_3.clicked.connect(self.create_win_path_to_letter)
        self.ui.pushButton_6.clicked.connect(self.create_win_path_to_certificate)
        self.ui.pushButton_5.clicked.connect(self.create_win_path_to_work_monitoring)
        self.ui.pushButton_8.clicked.connect(self.exit)
        
    def create_win_path_to_diploma(self):
        from diploma import WindowDiploma
        self.window_diploma = WindowDiploma()
        self.window_diploma.show()
        self.close()
    def create_win_path_to_letter(self):
        from letter import WindowLetter 
        self.window_letter = WindowLetter()
        self.window_letter.show()
        self.close()
    def create_win_path_to_certificate(self):
        from sertificate import WindowSertificate
        self.window_sertificate = WindowSertificate()
        self.window_sertificate.show()
        self.close()
    def create_win_path_to_work_monitoring(self):
        from work_monitoring import WindowWorkMonitoring
        self.window_work_monitoring = WindowWorkMonitoring()
        self.window_work_monitoring.show()
        self.close()
    def exit(self):
        from log_in import WindowLogIn
        self.log_in = WindowLogIn()
        self.log_in.show()
        self.close()
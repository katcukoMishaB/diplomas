from ui_log_in_registration import Ui_LogInWindow
from PyQt5 import QtWidgets

class WindowLogIn(QtWidgets.QMainWindow):
    def __init__(self):
        super(WindowLogIn, self).__init__()
        self.ui = Ui_LogInWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.log_in)
        self.ui.pushButton_2.clicked.connect(self.sign_in)

    def log_in(self):
        from main_window import MainWindow
        self.main = MainWindow()
        self.main.show()
        self.close()
    def sign_in(self):
        from sign_in import WindowSignIn
        self.main = WindowSignIn()
        self.main.show()
        self.close()
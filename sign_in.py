from ui_sign_in_registration import Ui_SignInWindow
from PyQt5 import QtWidgets

class WindowSignIn(QtWidgets.QMainWindow):
    def __init__(self):
        super(WindowSignIn, self).__init__()
        self.ui = Ui_SignInWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.log_in)
        self.ui.pushButton_2.clicked.connect(self.sign_in)

    def log_in(self):
        from log_in import WindowLogIn
        self.main = WindowLogIn()
        self.main.show()
        self.close()
    def sign_in(self):
        from main_window import MainWindow
        self.main = MainWindow()
        self.main.show()
        self.close()
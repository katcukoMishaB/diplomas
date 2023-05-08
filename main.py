from PyQt5.QtWidgets import QApplication
import sys
from log_in import WindowLogIn

if __name__ == '__main__':
    app = QApplication(sys.argv)
    log_in = WindowLogIn()
    log_in.show()
    sys.exit(app.exec_())
    
from PyQt5.QtWidgets import QApplication, QMainWindow
import sys
from Ui_MainWindow import MainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
    
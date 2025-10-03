import sys
# from ui import Ui_MainUI
from PyQt5.QtWidgets import QMainWindow, QApplication

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # self.main_ui = Ui_MainUI()
        # self.main_ui.setupUi(self)

if __name__ == "__main__":
        app = QApplication(sys.argv)
        application = MyWindow()
        application.show()
        sys.exit(app.exec_())
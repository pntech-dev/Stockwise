import sys

from ui.mainUI import Ui_MainWindow
from mvc import Model, View, Controller

from PyQt5.QtWidgets import QMainWindow, QApplication


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.main_ui = Ui_MainWindow()
        self.main_ui.setupUi(self)

        # MVC Initialization
        self.model = Model()
        self.view = View(ui=self.main_ui)
        self.controller = Controller(model=self.model, view=self.view)

if __name__ == "__main__":
        app = QApplication(sys.argv)
        application = MyWindow()
        application.show()
        sys.exit(app.exec_())
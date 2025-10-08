import sys

from resources import resources_rc
from ui.mainUI import Ui_MainWindow
from mvc import MainModel, MainView, MainController

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QApplication


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.main_ui = Ui_MainWindow()
        self.main_ui.setupUi(self)

        # Устанавливаем иконку приложения
        icon = QIcon(":/icons/icon.ico")
        self.setWindowIcon(icon)

        # Инициализация архитектуры MVC главного окна приложения
        self.main_model = MainModel()
        self.main_view = MainView(ui=self.main_ui)
        self.main_controller = MainController(model=self.main_model, view=self.main_view)

if __name__ == "__main__":
        app = QApplication(sys.argv)
        application = MyWindow()
        application.show()
        sys.exit(app.exec_())
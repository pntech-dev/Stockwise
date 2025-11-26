import sys

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow

from mvc import MainController, MainModel, MainView
from resources import resources_rc
from ui.mainUI import Ui_MainWindow


class MyWindow(QMainWindow):
    """The main window of the application.

    This class initializes the main UI, sets up the window icon, and
    assembles the Model-View-Controller (MVC) components for the application's
    main screen.
    """

    def __init__(self) -> None:
        """Initializes the main window and its components.

        Sets up the user interface from the generated Ui_MainWindow class,
        configures the window icon, and initializes the MainModel, MainView,
        and MainController to establish the application's core architecture.
        """
        super().__init__()

        self.main_ui: Ui_MainWindow = Ui_MainWindow()
        self.main_ui.setupUi(self)

        # Set the application icon
        icon: QIcon = QIcon(":/icons/icon.ico")
        self.setWindowIcon(icon)

        # Initialize the MVC architecture for the main application window
        self.main_model: MainModel = MainModel()
        self.main_view: MainView = MainView(ui=self.main_ui)
        self.main_controller: MainController = MainController(
            model=self.main_model, view=self.main_view
        )


if __name__ == "__main__":
    app: QApplication = QApplication(sys.argv)
    application: MyWindow = MyWindow()
    application.show()
    sys.exit(app.exec_())

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMessageBox

from resources import resources_rc


class Notification:
    """A utility class for displaying various types of message boxes.

    This class provides standardized methods to show simple notifications
    (info, warning, error) and confirmation dialogs with user actions.
    """

    def show_notification_message(self, msg_type: str, text: str) -> None:
        """Displays a simple notification message box.

        Args:
            msg_type: The type of notification. Can be "info", "warning",
                      or "error".
            text: The message text to display.
        """
        if not msg_type:
            return

        msg_box: QMessageBox = QMessageBox()
        msg_box.setWindowIcon(QIcon(":/icons/icon.ico"))
        msg_box.setText(text)

        if msg_type == "info":
            msg_box.setWindowTitle("Information")
            msg_box.setIcon(QMessageBox.Information)
        elif msg_type == "warning":
            msg_box.setWindowTitle("Warning")
            msg_box.setIcon(QMessageBox.Warning)
        elif msg_type == "error":
            msg_box.setWindowTitle("Error")
            msg_box.setIcon(QMessageBox.Critical)
        else:
            return

        msg_box.exec_()

    def show_action_message(
        self,
        msg_type: str = "error",
        title: str = "",
        text: str = "",
        buttons: list[str] = None,
    ) -> bool:
        """Displays a confirmation message box with action buttons.

        Args:
            msg_type: The type of icon to display. Can be "info", "warning",
                      or "error". Defaults to "error".
            title: The window title for the message box.
            text: The message text to display.
            buttons: A list of two strings for the button labels.
                     Defaults to ["Yes", "No"].

        Returns:
            True if the first button is clicked, False otherwise.
        """
        if buttons is None:
            buttons = ["Yes", "No"]

        msg: QMessageBox = QMessageBox()
        msg.setWindowIcon(QIcon(":/icons/icon.ico"))
        msg.setWindowTitle(title)
        msg.setText(text)

        if msg_type == "info":
            msg.setIcon(QMessageBox.Information)
        elif msg_type == "warning":
            msg.setIcon(QMessageBox.Warning)
        elif msg_type == "error":
            msg.setIcon(QMessageBox.Critical)

        yes_button = msg.addButton(buttons[0], QMessageBox.YesRole)
        msg.addButton(buttons[1], QMessageBox.NoRole)

        msg.exec_()

        return msg.clickedButton() == yes_button
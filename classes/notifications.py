from PyQt5.QtWidgets import QMessageBox

class Notification:
    def show_notification_message(self, msg_type, text):
        if not msg_type:
            return

        msg_box = QMessageBox()
        msg_box.setText(text)

        if msg_type == "info":
            msg_box.setWindowTitle("Инфомация")
            msg_box.setIcon(QMessageBox.Information)
        elif msg_type == "warning":
            msg_box.setWindowTitle("Предупреждение")
            msg_box.setIcon(QMessageBox.Warning)
        elif msg_type == "error":
            msg_box.setWindowTitle("Ошибка")
            msg_box.setIcon(QMessageBox.Critical)
        else:
            return
        
        msg_box.exec_()

    def show_action_message(self, msg_type="error", title="", text="", buttons=["Да", "Нет"]):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(text)
        
        if msg_type == "info":
            msg.setIcon(QMessageBox.Information)
        elif msg_type == "warning":
            msg.setIcon(QMessageBox.Warning)
        elif msg_type == "error":
            msg.setIcon(QMessageBox.Critical)

        yes_button = msg.addButton(buttons[0], QMessageBox.YesRole)
        no_button = msg.addButton(buttons[1], QMessageBox.NoRole)

        msg.exec_()

        if msg.clickedButton() == yes_button:
            return True
        else:
            return False

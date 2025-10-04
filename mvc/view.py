class View:
    def __init__(self, ui):
        self.ui = ui

    def get_search_field_text(self):
        """Функция возвращает текст из поля поиска."""
        return self.ui.search_line_lineEdit.text()

    def clear_search_field(self):
        """Функция очищает поле поиска."""
        self.ui.search_line_lineEdit.clear()

    def set_save_folder_path(self, path):
        """Функция устанавливает путь к папке для сохранения."""
        self.ui.save_file_path_line_lineEdit.setText(path)

    def update_clear_button_state(self, enabled):
        """Функция обновляет состояние кнопки очистки поля поиска."""
        self.ui.search_line_clear_pushButton.setEnabled(enabled)

    def search_field_changed(self, handler):
        """Функция вызывает обработчик при изменении поля поиска."""
        self.ui.search_line_lineEdit.textChanged.connect(handler)
    
    def clear_button_clicked(self, handler):
        """Функция вызывает обработчик при нажатии кнопки очистки поля поиска."""
        self.ui.search_line_clear_pushButton.clicked.connect(handler)

    def choose_save_folder_path(self, handler):
        """Функция вызывает обработчик при нажатии кнопки выбора папки для сохранения."""
        self.ui.save_file_path_line_choose_pushButton.clicked.connect(handler)
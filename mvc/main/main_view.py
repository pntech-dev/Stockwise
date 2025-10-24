from PyQt5.QtWidgets import QCompleter, QTableWidgetItem, QHeaderView


class MainView:
    def __init__(self, ui):
        self.ui = ui

        # Настройки таблицы
        headers = ['Номенклатура', 'Количество', 'Ед. изм.'] # Заголовки таблицы
        self.ui.data_tableWidget.setColumnCount(len(headers)) # Устанавливаем количество столбцов
        self.ui.data_tableWidget.setHorizontalHeaderLabels(headers) # Устанавливаем заголовки столбцов
        header = self.ui.data_tableWidget.horizontalHeader() # Получаем заголовок
        header.setSectionResizeMode(QHeaderView.ResizeToContents) # Растяжение колонок по данным в колонке

    def get_search_field_text(self):
        """Функция возвращает текст из поля поиска."""
        return self.ui.search_line_lineEdit.text()
    
    def get_search_line_edit(self):
        """Возвращает виджет QLineEdit для поиска."""
        return self.ui.search_line_lineEdit

    def set_search_completer(self, model):
        """Функция устанавливает QCompleter для поля поиска."""
        completer = QCompleter(model, self.ui.search_line_lineEdit)
        completer.setCaseSensitivity(0)
        completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        self.ui.search_line_lineEdit.setCompleter(completer)
        return completer
    
    def set_window_enabled_state(self, enabled):
        """Функция устанавливает состояние окна."""
        self.ui.centralwidget.setEnabled(enabled)

    def set_search_field_text(self, text):
        """Функция устанавливает текст в поле поиска."""
        self.ui.search_line_lineEdit.setText(text)

    def set_search_in_materials_checkbox_text(self, text):
        """Функция устанавливает текст в поле поиска."""
        checkbox_text = "Поиск по материалам изделия" if not text else f"Поиск по материалам изделия: {text}"
        self.ui.search_in_materials_checkBox.setText(checkbox_text)

    def update_clear_button_state(self, enabled):
        """Функция обновляет состояние кнопки очистки поля поиска."""
        self.ui.search_line_clear_pushButton.setEnabled(enabled)
    
    def update_table_widget_data(self, data):
        """Функция обновляет данные в таблице."""
        self.ui.data_tableWidget.setRowCount(0) # Очищаем только строки

        # Заполняем таблицу данными
        for row_index, item in enumerate(data): # Перебираем все строки данных
            self.ui.data_tableWidget.insertRow(row_index) # Добавляем строку в таблицу

            # Устанавливаем дааные в колонку "Номенклатура"
            nomenclature_item = QTableWidgetItem(item.get('Номенклатура', ''))
            self.ui.data_tableWidget.setItem(row_index, 0, nomenclature_item)

            # Устанавливаем данные в колонку "Количество"
            quantity_item = QTableWidgetItem(str(item.get('Количество', '')))
            self.ui.data_tableWidget.setItem(row_index, 1, quantity_item)

            # Устанавливаем данные в клонку "Ед. изм."
            unit_item = QTableWidgetItem(item.get('Ед. изм.', ''))
            self.ui.data_tableWidget.setItem(row_index, 2, unit_item)

    def update_create_document_button_state(self, enabled):
        """Функция обновляет состояние кнопки экспорта."""
        self.ui.create_document_pushButton.setEnabled(enabled)

    def update_export_button_state(self, enabled):
        """Функция обновляет состояние кнопки экспорта."""
        self.ui.export_pushButton.setEnabled(enabled)

    def clear_search_field(self):
        """Функция очищает поле поиска."""
        self.ui.search_line_lineEdit.clear()

    def search_field_changed(self, handler):
        """Функция вызывает обработчик при изменении поля поиска."""
        self.ui.search_line_lineEdit.textChanged.connect(handler)
    
    def clear_button_clicked(self, handler):
        """Функция вызывает обработчик при нажатии кнопки очистки поля поиска."""
        self.ui.search_line_clear_pushButton.clicked.connect(handler)

    def create_document_button_clicked(self, handler):
        """Функция вызывает обработчик при нажатии кнопки экспорта."""
        self.ui.create_document_pushButton.clicked.connect(handler)

    def norms_calculations_changed(self, handler):
        """Функция вызывает обработчик при изменении нормы расчета."""
        self.ui.norms_calculations_spinBox.valueChanged.connect(handler)

    def export_button_clicked(self, handler):
        """функция вызывает обработчик при нажатии кнопки экспорта"""
        self.ui.export_pushButton.clicked.connect(handler)

    def search_in_materials_checkbox_state_changed(self, handler):
        """Функция вызывает обработчик при изменении состояния чекбокса поиска в перечне материалов"""
        self.ui.search_in_materials_checkBox.stateChanged.connect(lambda state: handler(state == 2))
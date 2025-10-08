from PyQt5.QtWidgets import QCompleter, QTableWidgetItem, QHeaderView


class MainView:
    def __init__(self, ui):
        self.ui = ui

        # Настройки таблицы
        headers = ['Номенклатура', 'Количество', 'Ед. изм.', 'Изделие'] # Заголовки таблицы
        self.ui.data_tableWidget.setColumnCount(len(headers)) # Устанавливаем количество столбцов
        self.ui.data_tableWidget.setHorizontalHeaderLabels(headers) # Устанавливаем заголовки столбцов
        header = self.ui.data_tableWidget.horizontalHeader() # Получаем заголовок
        header.setSectionResizeMode(QHeaderView.ResizeToContents) # Растяжение колонок по данным в колонке

    def get_search_field_text(self):
        """Функция возвращает текст из поля поиска."""
        return self.ui.search_line_lineEdit.text()

    def clear_search_field(self):
        """Функция очищает поле поиска."""
        self.ui.search_line_lineEdit.clear()

    def set_search_completer(self, model):
        """Функция устанавливает QCompleter для поля поиска."""
        completer = QCompleter(model, self.ui.search_line_lineEdit)
        completer.setCaseSensitivity(0)
        completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        self.ui.search_line_lineEdit.setCompleter(completer)
        return completer

    def get_search_line_edit(self):
        """Возвращает виджет QLineEdit для поиска."""
        return self.ui.search_line_lineEdit

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

            # Устанавливаем данные в колонку "Изделие"
            product_item = QTableWidgetItem(", ".join(item.get('Изделие', '')))
            self.ui.data_tableWidget.setItem(row_index, 3, product_item)

    def update_export_button_state(self, enabled):
        """Функция обновляет состояние кнопки экспорта."""
        self.ui.export_pushButton.setEnabled(enabled)

    def search_field_changed(self, handler):
        """Функция вызывает обработчик при изменении поля поиска."""
        self.ui.search_line_lineEdit.textChanged.connect(handler)
    
    def clear_button_clicked(self, handler):
        """Функция вызывает обработчик при нажатии кнопки очистки поля поиска."""
        self.ui.search_line_clear_pushButton.clicked.connect(handler)

    def export_button_clicked(self, handler):
        """Функция вызывает обработчик при нажатии кнопки экспорта."""
        self.ui.export_pushButton.clicked.connect(handler)

    def norms_calculations_changed(self, handler):
        """Функция вызывает обработчик при изменении нормы расчета."""
        self.ui.norms_calculations_spinBox.valueChanged.connect(handler)
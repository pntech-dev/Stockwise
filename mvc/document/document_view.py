from PyQt5.QtWidgets import QCompleter


class DocumentView:
    def __init__(self, ui):
        self.ui = ui

    def get_save_folder_path(self):
        """Функция возвращает путь к папке сохранения файла."""
        return self.ui.save_file_path_line_lineEdit.text()

    def get_selected_document_type(self):
        """Функция возвращает выбранный тип документа."""
        if self.ui.document_radioButton.isChecked():
            return "document"
        elif self.ui.bid_radioButton.isChecked():
            return "bid"
        else:
            return None
    
    def get_export_format(self):
        """Функция возвращает формат экспорта."""
        if self.ui.export_formats_excel_radioButton.isChecked():
            return "excel"
        elif self.ui.export_formats_word_radioButton.isChecked():
            return "word"
        elif self.ui.export_formats_pdf_radioButton.isChecked():
            return "pdf"
        else:
            return None

    def set_product_nomenclature_lineedit(self, product_nomenclature):
        """Функция устанавливает значение поля номенклатуры изделия."""
        self.ui.product_nomenclature_lineEdit.setText(product_nomenclature)
    
    def set_quantity_spinbox(self, quantity):
        """Функция устанавливает значение поля количества изделий."""
        self.ui.product_count_spinBox.setValue(quantity)

    def set_completer(self, line_edit, data):
        """Функция устанавливает QCompleter для поля ввода."""
        completer = QCompleter(data, line_edit)
        completer.setCaseSensitivity(0)
        line_edit.setCompleter(completer)

    def set_current_date(self, date):
        """Функция устанавливает текущую дату в поле даты."""
        self.ui.date_dateEdit.setDate(date)

    def set_save_folder_path(self, path):
        """Функция устанавливает путь к папке сохранения файла."""
        self.ui.save_file_path_line_lineEdit.setText(path)

    def set_export_button_state(self, enabled):
        """Функция устанавливает состояние кнопки экспорта."""
        self.ui.export_pushButton.setEnabled(enabled)

    def choose_save_file_path_button_clicked(self, handler):
        """Функция устанавливает обработчик нажатия кнопки выбора пути сохранения файла."""
        self.ui.save_file_path_line_choose_pushButton.clicked.connect(handler)

    def document_type_radiobutton_clicked(self, handler):
        """Функция устанавливает обработчик на изменение типа документа."""
        self.ui.document_radioButton.clicked.connect(handler)
        self.ui.bid_radioButton.clicked.connect(handler)

    def export_button_clicked(self, handler):
        """Функция устанавливает обработчик нажатия кнопки экспорта."""
        self.ui.export_pushButton.clicked.connect(handler)
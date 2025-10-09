from PyQt5.QtWidgets import QCompleter


class DocumentView:
    def __init__(self, ui):
        self.ui = ui

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
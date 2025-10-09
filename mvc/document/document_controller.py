from PyQt5.QtWidgets import QFileDialog 


class DocumentController:
    def __init__(self, model, view):
        self.model = model
        self.view = view
        
        self.__set_lineedits() # Вызываем автоматическое заполнение полей ввода
        self.__set_completers() # Вызываем установку комплитеров
        self.__set_dateedit() # Вызываем установку текущей даты

        # Записываем сегодняшнюю дату, как дату по умолчанию
        self.model.current_date = self.view.get_date().toString("dd.MM.yyyy")

        # Обработчики
        self.view.choose_save_file_path_button_clicked(self.on_choose_save_file_path_button_clicked) # Нажатие кнопки выбора пути сохранения файла
        self.view.document_type_radiobutton_clicked(self.on_document_type_radiobutton_clicked) # Изменение типа документа
        self.view.export_button_clicked(self.export_button_clicked) # Нажатие кнопки экспорта
        self.view.outgoing_number_lineedit_text_changed(self.on_outgoing_number_lineedit_text_changed) # Изменение номера исходящего документа
        self.view.date_dateedit_changed(self.on_date_dateedit_changed) # Изменение даты
        self.view.from_position_lineedit_text_changed(self.on_from_position_lineedit_text_changed) # Изменение должности от кого
        self.view.from_fio_lineedit_text_changed(self.on_from_fio_lineedit_text_changed) # Изменение ФИО от кого
        self.view.whom_position_lineedit_text_changed(self.on_whom_position_lineedit_text_changed) # Изменение должности кому
        self.view.whom_fio_lineedit_text_changed(self.on_whom_fio_lineedit_text_changed) # Изменение ФИО

    def __set_lineedits(self):
        """Функция автоматического заполнения полей ввода."""
        self.__set_nomenclature_lineedit() # Поле номенклатура
        self.__set_quantity_spinbox() # Поле количества изделий

    def __set_nomenclature_lineedit(self):
        """Функция автоматического заполнения поля номенклатуры изделия."""
        self.view.set_product_nomenclature_lineedit(self.model.product_name)

    def __set_quantity_spinbox(self):
        """Функция автоматического заполнения поля количества изделий."""
        self.view.set_quantity_spinbox(self.model.quantity)

    def __set_completers(self):
        """Функция устанавливает комплитеры для полей ввода."""
        self.view.set_completer(self.view.ui.from_position_lineEdit, self.model.signature_from_position)
        self.view.set_completer(self.view.ui.from_fio_lineEdit, self.model.signature_from_human)
        self.view.set_completer(self.view.ui.whom_position_lineEdit, self.model.signature_whom_position)
        self.view.set_completer(self.view.ui.whom_fio_lineEdit, self.model.signature_whom_human)
    
    def __set_dateedit(self):
        """Функция устанавливает текущую дату в поле даты."""
        self.view.set_current_date(date=self.model.get_current_date())

    def on_choose_save_file_path_button_clicked(self):
        """Функция обрабатывает нажатие кнопки выбора пути сохранения файла."""
        folder_path = QFileDialog.getExistingDirectory() # Получаем путь к папке

        if folder_path: # Если путь получен
            self.view.set_save_folder_path(path=folder_path) # Записываем путь в QLineEdit

    def export_button_clicked(self):
        """Функция обрабатывает нажатие кнопки экспорта."""
        # Получаем состояния выбора типа документа
        # Первый объект - Докладная записка (Цех)
        # Второй объект - Заявка (ПДС)
        document_type = self.view.get_selected_document_type()

        # Вызываем экспорт документа в потоке
        self.model.export_in_thread(document_type=document_type, 
                                    save_folder_path=self.view.get_save_folder_path())

    def on_document_type_radiobutton_clicked(self):
        """Функция обрабатывает изменение типа документа."""
        self.view.set_export_button_state(True)
        
    def on_outgoing_number_lineedit_text_changed(self):
        """Функция обрабатывает изменение номера исходящего документа."""
        self.model.outgoing_number = self.view.get_outgoing_number()

    def on_date_dateedit_changed(self):
        """Функция обрабатывает изменение даты."""
        self.model.current_date = self.view.get_date().toString("dd.MM.yyyy")

    def on_from_position_lineedit_text_changed(self):
        """Функция обрабатывает изменение должности от кого."""
        self.model.from_position = self.view.get_from_position()

    def on_from_fio_lineedit_text_changed(self):
        """Функция обрабатывает изменение ФИО от кого."""
        self.model.from_fio = self.view.get_from_fio()

    def on_whom_position_lineedit_text_changed(self):
        """Функция обрабатывает изменение должности кому."""
        self.model.whom_position = self.view.get_whom_position()

    def on_whom_fio_lineedit_text_changed(self):
        """Функция обрабатывает изменение ФИО кому."""
        self.model.whom_fio = self.view.get_whom_fio()
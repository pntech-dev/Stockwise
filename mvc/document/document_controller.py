class DocumentController:
    def __init__(self, model, view):
        self.model = model
        self.view = view
        
        self.__set_lineedits() # Вызываем автоматическое заполнение полей ввода
        self.__set_completers() # Вызываем установку комплитеров
        self.__set_dateedit() # Вызываем установку текущей даты

    def __set_lineedits(self):
        """Функция автоматического заполнения полей ввода."""
        self.__set_nomenclature_lineedit() # Поле номенклатура
        self.__set_quantity_spinbox() # Поле количества изделий

    def __set_nomenclature_lineedit(self):
        """Функция автоматического заполнения поля номенклатуры изделия."""
        self.view.set_product_nomenclature_lineedit(self.model.current_product_name)

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
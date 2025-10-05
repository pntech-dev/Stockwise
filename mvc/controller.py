import sys

from PyQt5.QtWidgets import QFileDialog 


class Controller:
    def __init__(self, model, view):
        self.model = model
        self.view = view

        self.__check_available_products_folder() # Проверяем доступность папки изделий

        # Настраиваем QCompleter
        self.model.update_product_names()
        suggestions_lst = [" ".join(name) for name in self.model.products_names]
        self.view.set_search_completer(suggestions_lst)

        # Обработчики
        self.view.search_field_changed(self.on_search_field_changed) # Изменение текста в поле поиска
        self.view.clear_button_clicked(self.on_clear_button_clicked) # Нажатие кнопки очистки
        self.view.choose_save_folder_path(self.on_choose_save_folder_path) # Нажатие кнопки выбора папки для сохранения

    def __check_available_products_folder(self):
        """Функция проверяет доступность папки изделий на сервере."""
        if not self.model.is_products_folder_available:
            pass # ДОБАВИТЬ УВЕДОМЛЕНИЕ!!!
            sys.exit()

    def on_search_field_changed(self):
        """Функция обрабатывает изменение поля поиска."""
        # Обновляем состояние кнопки очистки если в поле поиска есть текст
        self.view.update_clear_button_state(enabled=len(self.view.get_search_field_text()) > 0)

        # Проверяем существует ли введёное название изделия в списке изделий
        product_name = self.view.get_search_field_text()
        
        semi_finished_products = []
        for name in self.model.products_names:
            if product_name == " ".join(name):
                self.model.current_product = " ".join(name)
                semi_finished_products = self.model.get_semi_finished_products(name)
                break
        
        if semi_finished_products:
            product_materials = self.model.get_product_materials(semi_finished_products)
            
            self.view.update_table_widget_data(data=product_materials)

    def on_clear_button_clicked(self):
        """Функция обрабатывает нажатие кнопки очистки поля поиска."""
        self.view.clear_search_field() # Очищаем поле поиска
        self.view.update_clear_button_state(enabled=False) # Изменяем состояние кнопки очистки

    def on_choose_save_folder_path(self):
        """Функция обрабатывает нажатие кнопки выбора папки для сохранения."""
        folder_path = QFileDialog.getExistingDirectory() # Получаем путь к папке

        if folder_path: # Если путь получен
            self.view.set_save_folder_path(path=folder_path) # Записываем путь в QLineEdit

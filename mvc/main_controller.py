import sys

from classes import Notification

from PyQt5.QtWidgets import QFileDialog 
from PyQt5.QtCore import QStringListModel, QSortFilterProxyModel, Qt


class CustomFilterProxyModel(QSortFilterProxyModel):
    def filterAcceptsRow(self, source_row, source_parent):
        index = self.sourceModel().index(source_row, 0, source_parent)
        item_text = self.sourceModel().data(index, Qt.DisplayRole)
        
        filter_string = self.filterRegExp().pattern()
        search_words = filter_string.lower().split()
        
        if not search_words:
            return True
            
        item_text_lower = item_text.lower()
        
        return all(word in item_text_lower for word in search_words)


class MainController:
    def __init__(self, model, view):
        self.model = model
        self.view = view
        self.is_highlighting = False

        self.__check_program_version() # Проверяем версию программы

        self.__check_available_products_folder() # Проверяем доступность папки изделий

        # Настраиваем QCompleter
        self.model.update_products_names()
        suggestions_lst = [" ".join(name) for name in self.model.products_names]
        
        string_list_model = QStringListModel(suggestions_lst)
        
        self.proxy_model = CustomFilterProxyModel()
        self.proxy_model.setSourceModel(string_list_model)
        
        self.completer = self.view.set_search_completer(self.proxy_model)
        self.completer.highlighted.connect(self.on_completer_highlighted)
        
        self.view.get_search_line_edit().textChanged.connect(self.on_text_changed_for_filter)

        # Обработчики
        self.view.search_field_changed(self.on_search_field_changed) # Изменение текста в поле поиска
        self.view.clear_button_clicked(self.on_clear_button_clicked) # Нажатие кнопки очистки
        self.view.export_button_clicked(self.on_export_button_clicked) # Нажатие кнопки экспорта
        self.view.norms_calculations_changed(self.on_norms_calculations_changed) # Изменение нормы расчета

        # Сигналы
        self.model.show_notification.connect(self.show_notification) # Сигнал показа уведомления

    def __check_program_version(self):
        """Функция проверяет версию программы"""
        is_version = self.model.check_program_version() # Проверяем версию программы

        # Если произошла ошибка во время проверки версии
        if is_version is None: 
            action = Notification().show_action_message(msg_type="error", 
                                                        title="Ошибка проверки версии", 
                                                        text="Во время проверки версии произошла ошибка!\nОшибка связана с путём к файлу конфигурации\nЖелаете открыть файл конфигурации?", 
                                                        buttons=["Да", "Нет"])

            if action:
                self.model.open_config_file()
                exit()
            else:
                exit()

        # Если версия не совпадает (требуется обновление)
        elif not is_version:
            action = Notification().show_action_message(msg_type="warning", 
                                                        title="Обновление", 
                                                        text="Обнаружена новая версия программы\nЖелаете обновить?", 
                                                        buttons=["Обновить", "Закрыть"])

            if action:
                self.model.update_program()
                sys.exit() # Закрываем основное приложение после запуска обновления
            else:
                sys.exit()


    def __check_available_products_folder(self):
        """Функция проверяет доступность папки изделий на сервере."""
        if not self.model.is_products_folder_available:
            sys.exit()

    def show_notification(self, msg_type, text):
        """Функция показывает уведомление."""
        Notification().show_notification_message(msg_type=msg_type, text=text)

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
            self.model.current_product_materials = product_materials
            self.view.update_table_widget_data(data=product_materials)
            self.view.update_export_button_state(enabled=True) # Изменяем состояние кнопки экспорта
        else:
            self.view.update_table_widget_data(data=[])
            self.view.update_export_button_state(enabled=False)

    def on_clear_button_clicked(self):
        """Функция обрабатывает нажатие кнопки очистки поля поиска."""
        self.view.clear_search_field() # Очищаем поле поиска
        self.view.update_clear_button_state(enabled=False) # Изменяем состояние кнопки очистки

    def on_export_button_clicked(self):
        """Функция обрабатывает нажатие кнопки экспорта."""
        print("Экспорт")

    def on_norms_calculations_changed(self, value):
        """Функция обрабатывает изменение нормы расчета."""
        self.model.norms_calculations_value = value
        self.on_search_field_changed() # Обновляем данные в таблице

    def on_completer_highlighted(self, text):
        self.is_highlighting = True

    def on_text_changed_for_filter(self, text):
        if self.is_highlighting:
            self.is_highlighting = False
            return
        self.proxy_model.setFilterRegExp(text)
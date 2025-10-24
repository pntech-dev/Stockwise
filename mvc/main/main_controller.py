import sys

from classes import Notification
from mvc.document import create_document_window

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

        self.document_window = None # Ссылка на окно документов

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
        self.view.create_document_button_clicked(self.on_create_document_button_clicked) # Нажатие кнопки создания документа
        self.view.norms_calculations_changed(self.on_norms_calculations_changed) # Изменение нормы расчета
        self.view.export_button_clicked(self.on_export_button_clicked) # Нажатие кнопки экспорта
        self.view.search_in_materials_checkbox_state_changed(self.on_search_in_materials_checkbox_state_changed) # Изменение состояния чекбокса поиска по материалам

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
        self.view.set_window_enabled_state(enabled=True) # Включаем окно

    def on_search_field_changed(self):
        """Функция обрабатывает изменение поля поиска."""
        search_text = self.view.get_search_field_text()
        self.view.update_clear_button_state(enabled=bool(search_text))

        data_for_view = []
        
        if not self.model.search_in_materials:
            product_name = search_text
            
            product_tuple = next((name for name in self.model.products_names if " ".join(name) == product_name), None)

            if product_tuple:
                self.model.current_product = product_name
                semi_finished_products = self.model.get_semi_finished_products(product_tuple)
                if semi_finished_products:
                    product_materials = self.model.get_product_materials(semi_finished_products)
                    self.model.current_product_materials = product_materials
                    data_for_view = product_materials
        else:
            search_words = search_text.strip().lower().split()
            
            data_for_view = [
                item for item in self.model.current_product_materials 
                if all(word in item['Номенклатура'].lower() for word in search_words)
            ]
            self.model.search_in_materials_data = data_for_view

        self.view.update_table_widget_data(data=data_for_view)
        buttons_enabled = bool(data_for_view)
        self.view.update_create_document_button_state(enabled=buttons_enabled)
        self.view.update_export_button_state(enabled=buttons_enabled)

    def on_clear_button_clicked(self):
        """Функция обрабатывает нажатие кнопки очистки поля поиска."""
        self.view.clear_search_field() # Очищаем поле поиска
        self.view.update_clear_button_state(enabled=False) # Изменяем состояние кнопки очистки

    def on_create_document_button_clicked(self):
        """Функция обрабатывает нажатие кнопки экспорта, вызывая создание окна создания документов."""
        # Проверка, выбрано ли изделие
        if not self.model.current_product_materials:
            self.show_notification("error", "Нет данных для экспорта.\nСначала выберите изделие.")
            return

        # Если окно ешё не открыто
        if self.document_window is None:
            # Создаём окно
            self.document_window = create_document_window(product_name=self.model.current_product,
                                                          norms_calculations_value=self.model.norms_calculations_value,
                                                          materials=self.model.current_product_materials,
                                                          current_product_path=self.model.current_product_path)
            self.document_window.show() # Показываем окно
            # Если окно было закрыто, вызываем отключение ссылки
            self.document_window.destroyed.connect(self.on_document_window_destroyed)
            self.view.update_create_document_button_state(enabled=False) # Отключаем кнопку "Создать документ"

    def on_document_window_destroyed(self):
        """Функция обрабатывает закрытие окна документов."""
        self.document_window = None # Обнуляем ссылку
        self.view.update_create_document_button_state(enabled=True) # Включаем кнопку "Создать документ"
        
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

    def on_export_button_clicked(self):
        """Функция обрабатывает нажатие кнопки экспорта."""
        self.view.set_window_enabled_state(enabled=False) # Отключаем окно
        self.model.export_data() # Вызываем экспорт данных

    def on_search_in_materials_checkbox_state_changed(self, state):
        """Функция обрабатывает изменение состояния чекбокса поиска по материалам."""
        self.model.search_in_materials = state
        
        if state:
            product_name = self.view.get_search_field_text()
            self.model.current_product = product_name
            self.view.clear_search_field()
            self.view.set_search_in_materials_checkbox_text(text=product_name)
            
            material_names = [item['Номенклатура'] for item in self.model.current_product_materials]
            string_list_model = QStringListModel(material_names)
            self.proxy_model.setSourceModel(string_list_model)
        else:
            self.view.set_search_field_text(self.model.current_product)
            self.model.current_product = ""
            self.view.set_search_in_materials_checkbox_text(text="")
            
            suggestions_lst = [" ".join(name) for name in self.model.products_names]
            string_list_model = QStringListModel(suggestions_lst)
            self.proxy_model.setSourceModel(string_list_model)
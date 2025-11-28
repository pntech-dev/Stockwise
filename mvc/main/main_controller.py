import sys

from PyQt5.QtCore import QStringListModel

from classes import Notification
from mvc.document import create_document_window
from utils.proxy_models import CustomFilterProxyModel



class MainController:
    """The main controller for the application.

    This class connects the main model and view, handling user interactions
    and business logic. It manages the application's state, responds to UI
    events, and coordinates data flow between the model and the view.
    """

    def __init__(self, model, view) -> None:
        """Initializes the MainController.

        Args:
            model: The main data model (`MainModel`).
            view: The main UI view (`MainView`).
        """
        self.model = model
        self.view = view
        self.is_highlighting: bool = False

        self.document_window = None  # Reference to the document window

        self._check_program_version()
        self._check_available_products_folder()

        # Set up the QCompleter for the search field
        self.model.update_products_names()
        suggestions_lst = [" ".join(name) for name in self.model.products_names]

        string_list_model = QStringListModel(suggestions_lst)

        self.proxy_model = CustomFilterProxyModel()
        self.proxy_model.setSourceModel(string_list_model)

        self.completer = self.view.set_search_completer(self.proxy_model)
        self.completer.highlighted.connect(self.on_completer_highlighted)

        self.view.get_search_line_edit().textChanged.connect(
            self.on_text_changed_for_filter
        )

        # Connect view signals to controller slots
        self.view.search_field_changed(self.on_search_field_changed)
        self.view.clear_button_clicked(self.on_clear_button_clicked)
        self.view.create_document_button_clicked(self.on_create_document_button_clicked)
        self.view.norms_calculations_changed(self.on_norms_calculations_changed)
        self.view.export_button_clicked(self.on_export_button_clicked)
        self.view.search_in_materials_checkbox_state_changed(
            self.on_search_in_materials_checkbox_state_changed
        )

        # Connect model signals to controller slots
        self.model.show_notification.connect(self.show_notification)

    def show_notification(self, msg_type: str, text: str) -> None:
        """Shows a notification message.

        This is a slot connected to the model's `show_notification` signal.

        Args:
            msg_type: The type of message ("info", "warning", "error").
            text: The notification text to display.
        """
        Notification().show_notification_message(msg_type=msg_type, text=text)
        self.view.set_window_enabled_state(enabled=True)  # Re-enable the window

    def on_search_field_changed(self) -> None:
        """Handles text changes in the search field to update the table."""
        search_text = self.view.get_search_field_text()
        self.view.update_clear_button_state(enabled=bool(search_text))

        data_for_view = []

        # If not searching within materials, search for products
        if not self.model.search_in_materials:
            product_name = search_text
            product_tuple = next(
                (
                    name
                    for name in self.model.products_names
                    if " ".join(name) == product_name
                ),
                None,
            )

            if product_tuple:
                self.model.current_product = product_name
                semi_finished_products = self.model.get_semi_finished_products(
                    product_tuple
                )
                if semi_finished_products:
                    product_materials = self.model.get_product_materials(
                        semi_finished_products
                    )
                    self.model.current_product_materials = product_materials
                    data_for_view = product_materials
        # Otherwise, filter the materials of the current product
        else:
            search_words = search_text.strip().lower().split()
            data_for_view = [
                item
                for item in self.model.current_product_materials
                if all(word in item["Номенклатура"].lower() for word in search_words)
            ]
            self.model.search_in_materials_data = data_for_view

        self.view.update_table_widget_data(data=data_for_view)
        buttons_enabled = bool(data_for_view)
        self.view.update_create_document_button_state(enabled=buttons_enabled)
        self.view.update_export_button_state(enabled=buttons_enabled)

    def on_clear_button_clicked(self) -> None:
        """Handles the click event of the clear search field button."""
        self.view.clear_search_field()
        self.view.update_clear_button_state(enabled=False)

    def on_create_document_button_clicked(self) -> None:
        """Handles the create document button click.

        Opens the document creation window if data is available.
        """
        if not self.model.current_product_materials:
            self.show_notification(
                "error", "Нет данных для экспорта.\nСначала выберите продукт."
            )
            return

        if self.document_window is None:
            self.document_window = create_document_window(
                product_name=self.model.current_product,
                norms_calculations_value=self.model.norms_calculations_value,
                materials=self.model.current_product_materials,
                current_product_path=self.model.current_product_path,
            )
            self.document_window.show()
            self.document_window.destroyed.connect(self.on_document_window_destroyed)
            self.view.update_create_document_button_state(enabled=False)

    def on_document_window_destroyed(self) -> None:
        """Handles the destruction of the document window.

        Resets the window reference and re-enables the create document button.
        """
        self.document_window = None
        self.view.update_create_document_button_state(enabled=True)

    def on_norms_calculations_changed(self, value: int) -> None:
        """Handles changes to the norms calculation value.

        Args:
            value: The new value for the calculation norm.
        """
        self.model.norms_calculations_value = value
        self.on_search_field_changed()  # Update the data in the table

    def on_completer_highlighted(self, text: str) -> None:
        """Sets a flag when a completer item is highlighted.

        This helps prevent the filter from re-running when the user is just
        navigating the suggestion list.

        Args:
            text: The highlighted text (not used).
        """
        self.is_highlighting = True

    def on_text_changed_for_filter(self, text: str) -> None:
        """Updates the completer filter, ignoring changes from highlighting.

        Args:
            text: The current text in the line edit.
        """
        if self.is_highlighting:
            self.is_highlighting = False
            return
        self.proxy_model.setFilterRegExp(text)

    def on_export_button_clicked(self) -> None:
        """Handles the click event of the export button."""
        self.view.set_window_enabled_state(enabled=False)
        self.model.export_data()

    def on_search_in_materials_checkbox_state_changed(self, state: int) -> None:
        """Handles the state change of the 'search in materials' checkbox.

        Switches the search mode and updates the completer model accordingly.

        Args:
            state: The new checked state from the checkbox signal
                   (0 for Unchecked, 2 for Checked).
        """
        is_checked = bool(state)
        self.model.search_in_materials = is_checked

        if is_checked:  # Switched to searching in materials
            product_name = self.view.get_search_field_text()
            self.model.current_product = product_name
            self.view.clear_search_field()
            self.view.set_search_in_materials_checkbox_text(text=product_name)

            material_names = [
                item["Номенклатура"] for item in self.model.current_product_materials
            ]
            string_list_model = QStringListModel(material_names)
            self.proxy_model.setSourceModel(string_list_model)
        else:  # Switched back to searching for products
            self.view.set_search_field_text(self.model.current_product)
            self.model.current_product = ""
            self.view.set_search_in_materials_checkbox_text(text="")

            suggestions_lst = [" ".join(name) for name in self.model.products_names]
            string_list_model = QStringListModel(suggestions_lst)
            self.proxy_model.setSourceModel(string_list_model)

    def _check_program_version(self) -> None:
        """Checks the program version and prompts for updates if necessary."""
        is_version_ok = self.model.check_program_version()

        # An error occurred during the version check
        if is_version_ok is None:
            action = Notification().show_action_message(
                msg_type="error",
                title="Ошибка проверки версии",
                text="При проверке версии произошла ошибка!\n" 
                "Это может быть связано с путем к файлу конфигурации.\n" 
                "Хотите открыть файл конфигурации?",
                buttons=["Да, открыть", "Нет, выйти"],
            )
            if action:
                self.model.open_config_file()
            sys.exit()

        # A new version is available
        elif not is_version_ok:
            action = Notification().show_action_message(
                msg_type="warning",
                title="Доступно обновление",
                text="Обнаружена новая версия программы.\n" 
                "Хотите обновить?",
                buttons=["Обновить", "Закрыть"],
            )
            if action:
                self.model.update_program()
            sys.exit()  # Close the main application

    def _check_available_products_folder(self) -> None:
        """Checks for the availability of the products folder on the server."""
        if not self.model.is_products_folder_available:
            sys.exit()
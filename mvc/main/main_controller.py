import sys

from PyQt5.QtCore import Qt, QStringListModel

from classes import Notification
from mvc.document import create_document_window
from utils.proxy_models import CustomFilterProxyModel


NOM_KEY = "Номенклатура"


class MainController:
    """The main controller for the application."""

    def __init__(self, model, view) -> None:
        self.model = model
        self.view = view
        self.is_highlighting: bool = False
        self.current_table_data: list[dict] = []

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
        self.view.header_checkbox_state_changed(self.on_header_checkbox_state_changed)

        # Connect model signals to controller slots
        self.model.show_notification.connect(self.show_notification)

        # Disable header checkbox until data appears
        self.view.set_header_checkbox_enabled(False)

    def show_notification(self, msg_type: str, text: str) -> None:
        """Shows a notification message and re-enables the window.

        Args:
            msg_type: Notification category such as ``"info"`` or ``"error"``.
            text: Message body to display.
        """
        Notification().show_notification_message(msg_type=msg_type, text=text)
        self.view.set_window_enabled_state(enabled=True)  # Re-enable the window

    def on_search_field_changed(self) -> None:
        """Refreshes table data when the search field text changes."""
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
                is_new_product = product_name != self.model.current_product
                self.model.current_product = product_name
                semi_finished_products = self.model.get_semi_finished_products(
                    product_tuple
                )
                if semi_finished_products:
                    product_materials = self.model.get_product_materials(
                        semi_finished_products
                    )
                    self.model.current_product_materials = product_materials
                    self.model.sync_material_selection(
                        product_materials, reset=is_new_product
                    )
                    data_for_view = product_materials
                else:
                    self.model.sync_material_selection([], reset=is_new_product)
                    self.model.current_product_materials = []
            else:
                self.model.current_product_materials = []
                self.model.sync_material_selection([], reset=True)
        # Otherwise, filter the materials of the current product
        else:
            search_words = search_text.strip().lower().split()
            data_for_view = [
                item
                for item in self.model.current_product_materials
                if all(word in item[NOM_KEY].lower() for word in search_words)
            ]
            self.model.search_in_materials_data = data_for_view

        self._update_table(data_for_view)

    def on_clear_button_clicked(self) -> None:
        """Clears the search field and disables the clear button."""
        self.view.clear_search_field()
        self.view.update_clear_button_state(enabled=False)

    def on_create_document_button_clicked(self) -> None:
        """Opens the document window for the selected materials."""
        if not self.model.current_product_materials:
            self.show_notification(
                "error",
                "Нет данных для документа. Выберите товар и повторите попытку.",
            )
            return

        selected_materials = [
            item
            for item in self.model.current_product_materials
            if self.model.material_selection.get(item[NOM_KEY], True)
        ]
        if not selected_materials:
            self.show_notification(
                "error",
                "Нет выбранных материалов для документа. Отметьте хотя бы один чекбокс.",
            )
            return

        if self.document_window is None:
            self.document_window = create_document_window(
                product_name=self.model.current_product,
                norms_calculations_value=self.model.norms_calculations_value,
                materials=selected_materials,
                current_product_path=self.model.current_product_path,
            )
            self.document_window.show()
            self.document_window.destroyed.connect(self.on_document_window_destroyed)
            self.view.update_create_document_button_state(enabled=False)

    def on_document_window_destroyed(self) -> None:
        """Resets state when the document window is closed."""
        self.document_window = None
        self.view.update_create_document_button_state(enabled=True)

    def on_norms_calculations_changed(self, value: int) -> None:
        """Applies a new norms multiplier and refreshes the table.

        Args:
            value: Multiplier used to scale material quantities.
        """
        self.model.norms_calculations_value = value
        self.on_search_field_changed()  # Update the data in the table

    def on_completer_highlighted(self, text: str) -> None:
        """Sets a flag when a completer item is highlighted.

        Args:
            text: Highlighted completer text (unused).
        """
        self.is_highlighting = True

    def on_text_changed_for_filter(self, text: str) -> None:
        """Updates the completer filter, ignoring changes from highlighting.

        Args:
            text: Current text in the search field.
        """
        if self.is_highlighting:
            self.is_highlighting = False
            return
        self.proxy_model.setFilterRegExp(text)

    def on_export_button_clicked(self) -> None:
        """Exports the current selection to Excel."""
        self.view.set_window_enabled_state(enabled=False)
        self.model.export_data()

    def on_search_in_materials_checkbox_state_changed(self, state: bool) -> None:
        """Toggles between product search and in-material search modes.

        Args:
            state: Checkbox state; True enables searching within materials.
        """
        is_checked = state
        self.model.search_in_materials = is_checked

        if is_checked:  # Switched to searching in materials
            product_name = self.view.get_search_field_text()
            self.model.current_product = product_name
            self.view.clear_search_field()
            self.view.set_search_in_materials_checkbox_text(text=product_name)

            material_names = [item[NOM_KEY] for item in self.model.current_product_materials]
            string_list_model = QStringListModel(material_names)
            self.proxy_model.setSourceModel(string_list_model)
        else:  # Switched back to searching for products
            self.view.set_search_field_text(self.model.current_product)
            self.model.current_product = ""
            self.view.set_search_in_materials_checkbox_text(text="")

            suggestions_lst = [" ".join(name) for name in self.model.products_names]
            string_list_model = QStringListModel(suggestions_lst)
            self.proxy_model.setSourceModel(string_list_model)

        self._refresh_table()

    def on_row_checkbox_state_changed(self, material_name: str, is_checked: bool) -> None:
        """Updates selection map when a row checkbox is toggled.

        Args:
            material_name: Name of the material tied to the row.
            is_checked: Whether the row is now selected.
        """
        self.model.set_material_selected(material_name, is_checked)
        self._update_buttons_and_header()

    def on_header_checkbox_state_changed(self, state: int) -> None:
        """Selects or deselects all materials when the header checkbox changes.

        Args:
            state: Qt check state from the header checkbox.
        """
        if not self.model.current_product_materials:
            self.view.set_header_checkbox_state(Qt.Unchecked)
            return

        select_all = state != Qt.Unchecked
        self.model.set_all_materials_selected(select_all)
        self._refresh_table()

    def _update_table(self, data: list[dict]) -> None:
        """Refreshes table data, buttons, and header checkbox state.

        Args:
            data: Rows to display in the materials table.
        """
        self.current_table_data = data
        any_selected = any(
            self.model.material_selection.get(item[NOM_KEY], True) for item in data
        )
        self.view.update_table_widget_data(
            data=data,
            selection=self.model.material_selection,
            on_row_checkbox_changed=self.on_row_checkbox_state_changed,
        )
        self._update_buttons_and_header()

    def _refresh_table(self) -> None:
        """Repaints table rows to reflect updated checkbox states."""
        self._update_table(self.current_table_data)

    def _update_buttons_and_header(self) -> None:
        """Updates action buttons and header checkbox state based on selection."""
        has_rows = bool(self.current_table_data)
        any_selected = has_rows and any(
            self.model.material_selection.get(item[NOM_KEY], True)
            for item in self.current_table_data
        )
        self.view.update_create_document_button_state(enabled=any_selected)
        self.view.update_export_button_state(enabled=any_selected)
        self.view.set_header_checkbox_enabled(has_rows)
        self._update_header_checkbox_state()

    def _update_header_checkbox_state(self) -> None:
        """Sets header checkbox state based on visible rows."""
        if not self.current_table_data:
            self.view.set_header_checkbox_state(Qt.Unchecked)
            return

        checked_count = 0
        for item in self.current_table_data:
            if self.model.material_selection.get(item[NOM_KEY], True):
                checked_count += 1

        if checked_count == len(self.current_table_data):
            state = Qt.Checked
        elif checked_count == 0:
            state = Qt.Unchecked
        else:
            state = Qt.PartiallyChecked

        self.view.set_header_checkbox_state(state)

    def _check_program_version(self) -> None:
        """Checks the program version and prompts for updates if necessary."""
        is_version_ok = self.model.check_program_version()

        # An error occurred during the version check
        if is_version_ok is None:
            action = Notification().show_action_message(
                msg_type="error",
                title="Ошибка проверки версии",
                text=(
                    "Не удалось проверить наличие обновлений.\n"
                    "Открыть config.yaml для проверки пути?"
                ),
                buttons=["Открыть config", "Закрыть"],
            )
            if action:
                self.model.open_config_file()
            sys.exit()

        # A new version is available
        elif not is_version_ok:
            action = Notification().show_action_message(
                msg_type="warning",
                title="Доступна новая версия",
                text="Найдена новая версия программы. Обновить сейчас?",
                buttons=["Обновить", "Выход"],
            )
            if action:
                self.model.update_program()
            sys.exit()  # Close the main application

    def _check_available_products_folder(self) -> None:
        """Checks for the availability of the products folder on the server."""
        if not self.model.is_products_folder_available:
            sys.exit()

from PyQt5.QtCore import QStringListModel
from PyQt5.QtWidgets import QFileDialog, QLineEdit

from classes import Notification
from utils.proxy_models import CustomFilterProxyModel



class DocumentController:
    """The controller for the document generation window.

    This class connects the document model and view, handling user input,
    populating the view with initial data, setting up completers, and
    initiating the document export process.
    """

    def __init__(self, model, view) -> None:
        """Initializes the DocumentController.

        Args:
            model: The data model for the document window.
            view: The view for the document window.
        """
        self.model = model
        self.view = view
        self.is_highlighting: bool = False
        self.proxy_models: dict[QLineEdit, CustomFilterProxyModel] = {}

        # Set initial state of the view
        self._set_lineedits()
        self._set_completers()
        self._set_dateedit()
        self.model.current_date = self.view.get_date().toString("dd.MM.yyyy")

        # Connect view signals to controller slots
        self.view.choose_save_file_path_button_clicked(
            self.on_choose_save_file_path_button_clicked
        )
        self.view.document_type_radiobutton_clicked(
            self.on_document_type_radiobutton_clicked
        )
        self.view.export_button_clicked(self.on_export_button_clicked)
        self.view.outgoing_number_lineedit_text_changed(
            self.on_outgoing_number_lineedit_text_changed
        )
        self.view.date_dateedit_changed(self.on_date_dateedit_changed)
        self.view.from_position_lineedit_text_changed(
            self.on_from_position_lineedit_text_changed
        )
        self.view.from_fio_lineedit_text_changed(self.on_from_fio_lineedit_text_changed)
        self.view.whom_position_lineedit_text_changed(
            self.on_whom_position_lineedit_text_changed
        )
        self.view.whom_fio_lineedit_text_changed(self.on_whom_fio_lineedit_text_changed)

        # Connect model signals to controller slots
        self.model.show_notification.connect(self.on_show_notification)
        self.model.progress_changed.connect(self.on_progress_bar_changed)
        self.model.export_fineshed.connect(self.on_export_finished)

    def on_completer_highlighted(self, text: str) -> None:
        """Sets a flag to indicate a completer item is highlighted."""
        self.is_highlighting = True

    def on_text_changed_for_filter(self, text: str, line_edit: QLineEdit) -> None:
        """Filters completer suggestions based on user input.

        This ignores text changes caused by highlighting a suggestion.

        Args:
            text: The current text in the line edit.
            line_edit: The specific QLineEdit that emitted the signal.
        """
        if self.is_highlighting:
            self.is_highlighting = False
            return

        if line_edit in self.proxy_models:
            self.proxy_models[line_edit].setFilterRegExp(text)

    def on_export_button_clicked(self) -> None:
        """Handles the export button click event."""
        document_type = self.view.get_selected_document_type()
        if not document_type:
            self.on_show_notification("warning", "Пожалуйста, выберите тип документа.")
            return

        self.model.export_in_thread(
            document_type=document_type,
            save_folder_path=self.view.get_save_folder_path(),
        )

        self.view.set_export_button_state(enabled=False)

    def on_show_notification(self, msg_type: str, text: str) -> None:
        """Displays a notification and resets the progress bar."""
        Notification().show_notification_message(msg_type=msg_type, text=text)
        self.view.set_progress_bar_value(value=0)
        self.view.set_progress_bar_labels_text(text="Process...", value=0)

    def on_choose_save_file_path_button_clicked(self) -> None:
        """Opens a dialog to choose a folder and updates the view."""
        folder_path = QFileDialog.getExistingDirectory()
        if folder_path:
            self.view.set_save_folder_path(path=folder_path)

    def on_document_type_radiobutton_clicked(self) -> None:
        """Enables the export button when a document type is selected."""
        self.view.set_export_button_state(True)

    def on_outgoing_number_lineedit_text_changed(self) -> None:
        """Updates the outgoing number in the model."""
        self.model.outgoing_number = self.view.get_outgoing_number()

    def on_date_dateedit_changed(self) -> None:
        """Updates the current date in the model."""
        self.model.current_date = self.view.get_date().toString("dd.MM.yyyy")

    def on_from_position_lineedit_text_changed(self) -> None:
        """Updates the 'sender' position in the model."""
        self.model.from_position = self.view.get_from_position()

    def on_from_fio_lineedit_text_changed(self) -> None:
        """Updates the 'sender' full name in the model."""
        self.model.from_fio = self.view.get_from_fio()

    def on_whom_position_lineedit_text_changed(self) -> None:
        """Updates the 'recipient' position in the model."""
        self.model.whom_position = self.view.get_whom_position()

    def on_whom_fio_lineedit_text_changed(self) -> None:
        """Updates the 'recipient' full name in the model."""
        self.model.whom_fio = self.view.get_whom_fio()

    def on_progress_bar_changed(self, text: str, value: int) -> None:
        """Updates the progress bar's value and text label."""
        self.view.set_progress_bar_value(value)
        self.view.set_progress_bar_labels_text(text=text, value=value)

    def on_export_finished(self) -> None:
        """Enabled export button state"""
        self.view.set_export_button_state(enabled=True)

    def _set_lineedits(self) -> None:
        """Populates various line edits with initial data from the model."""
        self.view.set_product_nomenclature_lineedit(self.model.product_name)
        self.view.set_quantity_spinbox(self.model.quantity)

    def _set_completers(self) -> None:
        """Sets up completers for fields that have suggestions."""
        completer_fields = {
            self.view.ui.from_position_lineEdit: self.model.signature_from_position,
            self.view.ui.from_fio_lineEdit: self.model.signature_from_human,
            self.view.ui.whom_position_lineEdit: self.model.signature_whom_position,
            self.view.ui.whom_fio_lineEdit: self.model.signature_whom_human,
        }

        for line_edit, suggestions in completer_fields.items():
            string_list_model = QStringListModel(suggestions)
            proxy_model = CustomFilterProxyModel()
            proxy_model.setSourceModel(string_list_model)
            self.proxy_models[line_edit] = proxy_model

            completer = self.view.set_completer(line_edit, proxy_model)
            completer.highlighted.connect(self.on_completer_highlighted)

            # Use a lambda to pass the specific line_edit to the slot
            line_edit.textChanged.connect(
                lambda text, le=line_edit: self.on_text_changed_for_filter(text, le)
            )

    def _set_dateedit(self) -> None:
        """Sets the current date in the date edit widget."""
        self.view.set_current_date(date=self.model.get_current_date())
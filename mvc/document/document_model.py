import os
import threading
from copy import copy
from pathlib import Path
from typing import Dict, List, Optional

import yaml
from jinja2 import Template
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.workbook import Workbook
from PyQt5.QtCore import QDate, QObject, pyqtSignal

NOM_KEY = "Номенклатура"
QTY_KEY = "Количество"
UNIT_KEY = "Ед. изм."
RMP_KEY = "РМП"


class DocumentModel(QObject):
    """Manages data and business logic for the document generation window.

    This class loads necessary configuration, filters materials based on
    document type, and handles the creation of styled Excel documents in a
    separate thread to keep the UI responsive.

    Attributes:
        show_notification: Signal emitting messages for the user.
        progress_changed: Signal to update the progress bar during export.
        export_fineshed: Export completion signal.
    """

    show_notification = pyqtSignal(str, str)
    progress_changed = pyqtSignal(str, int)
    export_fineshed = pyqtSignal()

    def __init__(
        self,
        product_name: str,
        norms_calculations_value: int,
        materials: List[Dict],
        current_product_path: str,
    ) -> None:
        """Initializes document data and settings.

        Args:
            product_name: Name of the product the documents describe.
            norms_calculations_value: Quantity multiplier applied to exported data.
            materials: Full materials list for the current product.
            current_product_path: Filesystem path to the current product folder.
        """
        super().__init__()

        self.product_name: str = product_name
        self.materials: List[Dict] = materials
        self.current_product_path: str = current_product_path

        self.signature_from_human: List[str] = []
        self.signature_from_position: List[str] = []
        self.signature_whom_human: List[str] = []
        self.signature_whom_position: List[str] = []
        self.document_blacklist: List[str] = []
        self.document_whitelist: List[str] = []
        self.bid_blacklist: List[str] = []
        self.bid_whitelist: List[str] = []
        self.templates_folder_path: str = ""
        self._load_config()

        self.outgoing_number: str = ""
        self.current_date: str = ""
        self.quantity: int = norms_calculations_value
        self.whom_position: str = ""
        self.whom_fio: str = ""
        self.from_position: str = ""
        self.from_fio: str = ""

        self.progress_bar_export_excel_step_size: int = 7

    def get_current_date(self) -> QDate:
        """Returns the current date."""
        return QDate.currentDate()

    def get_desktop_path(self) -> Optional[Path]:
        """Resolves a Desktop path, handling OneDrive RU/EN variants.

        Returns:
            Desktop path if it can be resolved, otherwise ``None``.
        """
        try:
            home = Path.home()
            onedrive_ru = home / "OneDrive" / "Рабочий стол"
            if onedrive_ru.exists():
                return onedrive_ru
            onedrive_en = home / "OneDrive" / "Desktop"
            if onedrive_en.exists():
                return onedrive_en
            return home / "Desktop"
        except Exception as e:
            self.show_notification.emit("error", f"Не удалось определить путь рабочего стола: {e}")
            return None

    def export_in_thread(self, document_type: str, save_folder_path: str) -> Optional[threading.Thread]:
        """Starts export in a background thread to keep the UI responsive.

        Args:
            document_type: Document variant to render (for example ``"document"`` or ``"bid"``).
            save_folder_path: Directory to place the output file; Desktop is used if empty.

        Returns:
            Thread object if started successfully, otherwise ``None``.
        """
        path = save_folder_path
        if not path:
            desktop = self.get_desktop_path()
            if not desktop:
                self.show_notification.emit("error", "Не удалось определить путь сохранения.")
                return None
            path = str(desktop)

        try:
            thread = threading.Thread(target=self._export_to_excel, args=(document_type, path))
            thread.daemon = True
            thread.start()
            return thread
        except Exception as e:
            self.show_notification.emit("error", f"Не удалось запустить экспорт: {e}")
            return None

    def _load_config(self) -> None:
        """Loads template paths and signature settings from config.yaml."""
        try:
            with open("config.yaml", "r", encoding="utf-8") as file:
                config = yaml.safe_load(file) or {}
        except FileNotFoundError:
            self.show_notification.emit("error", "Файл config.yaml не найден.")
            return

        configured_templates = config.get("templates_folder_path")
        if configured_templates and os.path.isdir(configured_templates):
            self.templates_folder_path = configured_templates
        else:
            local_templates = os.path.join(os.getcwd(), "templates")
            if os.path.isdir(local_templates):
                self.templates_folder_path = local_templates
            else:
                self.templates_folder_path = ""
                self.show_notification.emit(
                    "error",
                    "Путь к шаблонам не найден. Укажите templates_folder_path в config.yaml.",
                )

        self.signature_from_human = config.get("signature_from_human", [])
        self.signature_from_position = config.get("signature_from_position", [])
        self.signature_whom_human = config.get("signature_whom_human", [])
        self.signature_whom_position = config.get("signature_whom_position", [])
        self.bid_blacklist = config.get("bid_blacklist", [])
        self.document_blacklist = config.get("document_blacklist", [])
        self.bid_whitelist = config.get("bid_whitelist", [])
        self.document_whitelist = config.get("document_whitelist", [])

    def _get_document_materials_list(self) -> List[Dict]:
        """Builds a filtered materials list for documents using whitelist/blacklist.

        Returns:
            Materials allowed for the main document.
        """
        filtered: List[Dict] = []
        try:
            for item in self.materials:
                nomenclature_lower = item[NOM_KEY].lower()

                if any(word.lower() in nomenclature_lower for word in self.document_whitelist if word):
                    filtered.append(item)
                    continue

                if item[UNIT_KEY] != "шт" and not item[RMP_KEY]:
                    is_in_blacklist = any(word.lower() in nomenclature_lower for word in self.document_blacklist if word)
                    if not is_in_blacklist:
                        filtered.append(item)
            return filtered
        except Exception as e:
            self.show_notification.emit("error", f"Не удалось получить список материалов: {e}")
            return []

    def _get_bid_materials_list(self) -> List[Dict]:
        """Builds a filtered materials list for bids using whitelist/blacklist.

        Returns:
            Materials allowed for the bid document.
        """
        filtered: List[Dict] = []
        try:
            for item in self.materials:
                nomenclature_lower = item[NOM_KEY].lower()

                if any(word.lower() in nomenclature_lower for word in self.bid_whitelist if word):
                    filtered.append(item)
                    continue

                if item[UNIT_KEY] == "шт":
                    is_in_blacklist = any(word.lower() in nomenclature_lower for word in self.bid_blacklist if word)
                    if not is_in_blacklist:
                        filtered.append(item)
            return filtered
        except Exception as e:
            self.show_notification.emit("error", f"Не удалось получить список материалов для заявки: {e}")
            return []

    def _export_materials_list(
        self,
        workbook: Workbook,
        materials_list: List[Dict],
        progress_bar_value: int,
        progress_bar_process_text: str,
    ) -> int:
        """Copies the table template sheet and fills it with material rows.

        Args:
            workbook: Workbook that will receive the populated sheet.
            materials_list: Materials to insert into the table.
            progress_bar_value: Current progress bar value to update incrementally.
            progress_bar_process_text: Text displayed alongside the progress bar.

        Returns:
            Updated progress bar value after processing.
        """
        if not self.templates_folder_path:
            self.show_notification.emit(
                "error",
                "Путь к шаблонам не задан. Укажите templates_folder_path в config.yaml.",
            )
            return progress_bar_value

        try:
            table_template_path = os.path.join(self.templates_folder_path, "table.xlsx")
            if not os.path.exists(table_template_path):
                self.show_notification.emit("error", f"Файл шаблона не найден: {table_template_path}")
                return progress_bar_value

            template_wb = load_workbook(table_template_path)
            template_sheet = template_wb.active
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            new_sheet = workbook.create_sheet(title="Материалы")
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            for row in template_sheet.iter_rows():
                for cell in row:
                    new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = cell.number_format
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)

            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            for col_letter, col_dim in template_sheet.column_dimensions.items():
                new_sheet.column_dimensions[col_letter].width = col_dim.width
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            for row_idx, row_dim in template_sheet.row_dimensions.items():
                new_sheet.row_dimensions[row_idx].height = row_dim.height
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            start_row = 2
            for i, item in enumerate(materials_list):
                row = start_row + i
                new_sheet.cell(row=row, column=1, value=item[NOM_KEY])
                new_sheet.cell(row=row, column=2, value=item[UNIT_KEY])
                new_sheet.cell(row=row, column=3, value=item[QTY_KEY])
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            font = Font(name="Times New Roman", size=14)
            for row in new_sheet.iter_rows(min_row=start_row, max_row=new_sheet.max_row, max_col=3):
                for cell in row:
                    cell.font = font
                    if cell.column == 1:
                        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

            thick_side = Side(border_style="thick", color="000000")
            thin_side = Side(border_style="thin", color="000000")
            for row_idx in range(start_row, new_sheet.max_row + 1):
                for col_idx in range(1, 4):
                    cell = new_sheet.cell(row=row_idx, column=col_idx)
                    left = thick_side if col_idx == 1 else thin_side
                    right = thick_side if col_idx == 3 else thin_side
                    cell.border = Border(left=left, right=right, top=thin_side, bottom=thin_side)
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            new_sheet.page_setup.fitToWidth = 1
            new_sheet.page_setup.fitToHeight = 0
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_bar_process_text, progress_bar_value)

            return progress_bar_value
        except Exception as e:
            self.show_notification.emit("error", f"Ошибка при формировании листа материалов:\n{e}")
            return progress_bar_value

    def _export_to_excel(self, document_type: str, save_folder_path: str) -> None:
        """Renders the selected document type to Excel and writes it to disk.

        Args:
            document_type: Export mode, either ``"document"`` or ``"bid"``.
            save_folder_path: Target directory for the generated file.
        """
        if not self.templates_folder_path:
            self.show_notification.emit(
                "error",
                "Путь к шаблонам не задан. Укажите templates_folder_path в config.yaml.",
            )
            self.export_fineshed.emit()
            return

        context = {
            "outgoing_number": self.outgoing_number,
            "current_date": self.current_date,
            "whom_position": self.whom_position,
            "whom_fio": self.whom_fio,
            "product_name": self.product_name,
            "product_quantity": self.quantity,
            "from_position": self.from_position,
            "from_fio": self.from_fio,
        }
        progress_bar_value = 0

        try:
            if document_type == "document":
                template_path = os.path.join(self.templates_folder_path, "document.xlsx")
                safe_product = self._sanitize_filename(self.product_name) or ""
                save_filename = f"Докладная записка {safe_product}.xlsx"
                sheet_title = "Докладная записка"
                materials_list = self._get_document_materials_list()
            elif document_type == "bid":
                template_path = os.path.join(self.templates_folder_path, "bid.xlsx")
                safe_product = self._sanitize_filename(self.product_name) or ""
                save_filename = f"Заявка {safe_product}.xlsx"
                sheet_title = "Заявка"
                materials_list = self._get_bid_materials_list()
            else:
                self.export_fineshed.emit()
                return

            if not os.path.exists(template_path):
                self.show_notification.emit("error", f"Файл шаблона не найден: {template_path}")
                self.export_fineshed.emit()
                return

            save_path = Path(save_folder_path) / save_filename
            progress_text = f"Экспорт в {save_filename}..."

            wb = load_workbook(template_path)
            ws = wb.active
            ws.title = sheet_title
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_text, progress_bar_value)

            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and "{{" in cell.value:
                        template = Template(cell.value)
                        cell.value = template.render(context)
            progress_bar_value += self.progress_bar_export_excel_step_size
            self.progress_changed.emit(progress_text, progress_bar_value)

            progress_bar_value = self._export_materials_list(
                workbook=wb,
                materials_list=materials_list,
                progress_bar_value=progress_bar_value,
                progress_bar_process_text=progress_text,
            )

            wb.save(save_path)
            self.progress_changed.emit("Экспорт завершён", 100)
            self.show_notification.emit("info", f"Экспорт в {save_filename} успешно выполнен")
            self.export_fineshed.emit()
            
        except Exception as e:
            self.progress_changed.emit("Экспорт не выполнен", 100)
            self.show_notification.emit("error", f"Не удалось выполнить экспорт документа: {e}")
            self.export_fineshed.emit()

    def _sanitize_filename(self, name: str) -> str:
        """Strips invalid characters from a filename-friendly string.

        Args:
            name: Raw filename component.

        Returns:
            Safe filename fragment; defaults to ``"file"`` when empty.
        """
        invalid_chars = set('\\/:*?"<>|')
        cleaned = "".join(ch for ch in name if ch not in invalid_chars).strip()
        return cleaned or "file"

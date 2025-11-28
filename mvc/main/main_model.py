import os
import yaml
import subprocess
import pandas as pd

from pathlib import Path
from openpyxl import load_workbook
from PyQt5.QtCore import QObject, pyqtSignal
from openpyxl.styles import Alignment, Border, Font, Side

NOM_KEY = "Номенклатура"
QTY_KEY = "Количество"
UNIT_KEY = "Ед. изм."
RMP_KEY = "РМП"


class MainModel(QObject):
    """Manages the application's data and business logic."""

    show_notification = pyqtSignal(str, str)

    def __init__(self) -> None:
        """Initializes the MainModel."""
        super().__init__()

        # Program configuration
        self.program_version_number: str = ""
        self.program_server_path: str = ""
        self.path_to_products_folder: str = ""
        self.is_products_folder_available: bool = False
        self._load_config()

        # Product data
        self.products_names: list[list[str]] = []
        self.update_products_names()

        # Current state
        self.current_product: str = ""
        self.current_product_path: str = ""
        self.current_product_materials: list[dict] = []
        self.material_selection: dict[str, bool] = {}
        self.norms_calculations_value: int = 1
        self.search_in_materials: bool = False
        self.search_in_materials_data: list[dict] = []

    def get_semi_finished_products(self, product_name: tuple) -> list[str]:
        """Gets a list of semi-finished product files for a given product."""

        def add_file(file_path: str, file_list: list[str]) -> None:
            """Adds a file to the list if it is a valid Excel file."""
            if os.path.isfile(file_path) and file_path.lower().endswith(
                (".xlsx", ".xls")
            ):
                file_list.append(file_path)

        semi_finished_products = []
        try:
            product_path = os.path.join(self.path_to_products_folder, *product_name)

            if os.path.exists(product_path) and os.path.isdir(product_path):
                self.current_product_path = product_path
                for item_name in os.listdir(product_path):
                    item_path = os.path.join(product_path, item_name)
                    # Handle the 'RMP' (raw material product) special folder
                    if item_name.lower() == "��??��??��?��?" and os.path.isdir(item_path):
                        for rmp_file_name in os.listdir(item_path):
                            rmp_file_path = os.path.join(item_path, rmp_file_name)
                            add_file(rmp_file_path, semi_finished_products)
                    else:
                        add_file(item_path, semi_finished_products)
            return semi_finished_products
        except Exception as e:
            self.show_notification.emit(
                "error", f"Failed to collect semi-finished products: {e}"
            )
            return []

    def get_product_materials(
        self, semi_finished_products: list[str]
    ) -> list[dict]:
        """Aggregates all materials from a list of semi-finished product files."""
        product_materials_dict = {}
        semi_product_names = [
            os.path.splitext(os.path.basename(p))[0].lower().strip()
            for p in semi_finished_products
        ]

        for file_path in semi_finished_products:
            if not os.path.exists(file_path):
                self.show_notification.emit(
                    "error", f"File not found: {file_path}"
                )
                continue

            try:
                is_rmp = (
                    os.path.basename(os.path.dirname(file_path)).lower() == "��??��??��?��?"
                )
                df = pd.read_excel(file_path, sheet_name="TDSheet")
                df = self._normalize_material_columns(df)
                df = df.dropna(subset=[NOM_KEY])
                df = df.fillna("")

                for item in df.to_dict("records"):
                    nomenclature = item[NOM_KEY]
                    # Skip materials that are themselves semi-finished products
                    if nomenclature.lower().strip() in semi_product_names:
                        continue

                    # Calculate quantity based on norm
                    quantity = item[QTY_KEY] / 1000
                    if self.norms_calculations_value != 1:
                        quantity *= self.norms_calculations_value
                    item[QTY_KEY] = quantity
                    item[RMP_KEY] = is_rmp

                    # Aggregate materials
                    if nomenclature in product_materials_dict:
                        product_materials_dict[nomenclature][QTY_KEY] += quantity
                    else:
                        product_materials_dict[nomenclature] = item

            except Exception as e:
                self.show_notification.emit(
                    "error", f"Failed to read file {os.path.basename(file_path)}: {e}"
                )

        return list(product_materials_dict.values())

    def sync_material_selection(
        self, materials: list[dict], reset: bool = False
    ) -> None:
        """Keeps the materials selection map in sync with the current dataset."""
        new_selection = {}
        for item in materials:
            name = item[NOM_KEY]
            if reset or name not in self.material_selection:
                new_selection[name] = True
            else:
                new_selection[name] = self.material_selection[name]
        self.material_selection = new_selection

    def set_material_selected(self, name: str, is_selected: bool) -> None:
        """Updates selection state for a single material."""
        if name in self.material_selection:
            self.material_selection[name] = is_selected

    def set_all_materials_selected(self, is_selected: bool) -> None:
        """Sets selection state for all materials."""
        for key in list(self.material_selection.keys()):
            self.material_selection[key] = is_selected

    def get_desktop_path(self) -> Path | None:
        """Finds the path to the user's desktop folder."""
        try:
            # Check for Russian OneDrive path
            onedrive_ru = Path.home() / "OneDrive" / "�����+��?�?�� ��?�?'��?�?>"
            if onedrive_ru.exists():
                return onedrive_ru

            # Check for English OneDrive path
            onedrive_en = Path.home() / "OneDrive" / "Desktop"
            if onedrive_en.exists():
                return onedrive_en

            # Fallback to standard Desktop path
            return Path.home() / "Desktop"
        except Exception as e:
            self.show_notification.emit("error", f"Failed to resolve Desktop path: {e}")
            return None

    def check_program_version(self) -> bool | None:
        """Checks if the local program version is up-to-date with the server."""
        if not self.program_server_path or not os.path.exists(self.program_server_path):
            return None

        for file_name in os.listdir(self.program_server_path):
            if file_name.startswith("config") and file_name.endswith(".yaml"):
                try:
                    with open(
                        os.path.join(self.program_server_path, file_name),
                        "r",
                        encoding="utf-8",
                    ) as f:
                        config_data = yaml.safe_load(f)
                    server_version = config_data.get("program_version_number")
                    if server_version is None:
                        return None
                    return server_version <= self.program_version_number
                except (IOError, yaml.YAMLError, KeyError):
                    # Ignore corrupted config files or files that don't match schema
                    continue
        return False  # Return False if no valid config found on server

    def update_products_names(self) -> None:
        """Recursively finds all products in the products folder."""
        if not self.is_products_folder_available:
            self.products_names = []
            self.show_notification.emit(
                "error",
                "Products folder is not available.",
            )
            return

        products = []
        try:
            for root, dirs, _ in os.walk(self.path_to_products_folder):
                # Exclude the 'rmp' directory from the walk
                if "��??��??��?��?" in [d.lower() for d in dirs]:
                    dirs.remove("��??��??��?��?" if "��??��??��?��?" in dirs else "��??��?")

                # A directory is considered a product if it has no subdirectories
                if not dirs:
                    relative_path = os.path.relpath(root, self.path_to_products_folder)
                    if relative_path != ".":
                        products.append(relative_path.split(os.sep))
            self.products_names = products
        except Exception as e:
            self.show_notification.emit(
                "error", f"Failed to scan products: {e}"
            )

    def update_program(self) -> None:
        """Launches the updater executable with administrator privileges."""
        updater_path = os.path.join(os.getcwd(), "updater.exe")
        if not os.path.exists(updater_path):
            self.show_notification.emit(
                "error", f"Updater executable not found: {updater_path}"
            )
            return

        try:
            # Use PowerShell to run the updater with admin rights
            command = f'Start-Process "{updater_path}" -ArgumentList "{self.program_server_path}" -Verb RunAs'
            subprocess.run(["powershell", "-Command", command], check=True, shell=True)
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            self.show_notification.emit(
                "error", f"Failed to launch updater: {e}"
            )

    def export_data(self) -> None:
        """Exports the current materials list to a styled Excel file."""
        try:
            workbook = load_workbook("templates/table.xlsx")
            sheet = workbook.active
            sheet.title = self._sanitize_sheet_title("Материалы")

            selected_data = [
                item
                for item in self.current_product_materials
                if self.material_selection.get(item[NOM_KEY], True)
            ]

            if not selected_data:
                self.show_notification.emit("warning", "Нет данных для экспорта.")
                return

            # --- Fill data ---
            start_row = 2
            for i, item in enumerate(selected_data):
                row_idx = start_row + i
                sheet.cell(row=row_idx, column=1, value=item[NOM_KEY])
                sheet.cell(row=row_idx, column=2, value=item[UNIT_KEY])
                sheet.cell(row=row_idx, column=3, value=item[QTY_KEY])

            # --- Apply styles ---
            font = Font(name="Times New Roman", size=14)
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            for row in sheet.iter_rows(
                min_row=start_row, max_row=sheet.max_row, max_col=3
            ):
                for cell in row:
                    cell.font = font
                    cell.border = thin_border
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

                # Column-specific alignments
                sheet.cell(row=cell.row, column=1).alignment = Alignment(
                    horizontal="left", vertical="top", wrap_text=True
                )
                sheet.cell(row=cell.row, column=2).alignment = Alignment(
                    horizontal="center", vertical="top", wrap_text=True
                )
                sheet.cell(row=cell.row, column=3).alignment = Alignment(
                    horizontal="center", vertical="top", wrap_text=True
                )

            # --- Page setup ---
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.fitToHeight = 0

            # --- Save file ---
            desktop_path = self.get_desktop_path()
            if not desktop_path:
                self.show_notification.emit(
                    "error", "Не удалось определить путь сохранения."
                )
                return

            file_name = f"Материалы изделия {self.current_product}.xlsx"
            file_path = desktop_path / file_name
            workbook.save(file_path)

            self.show_notification.emit(
                "info", f"Экспорт выполнен: {file_name}"
            )

        except Exception as e:
            self.show_notification.emit(
                "error", f"Не удалось выполнить экспорт: {e}"
            )
        finally:
            self.show_notification.emit("", "")  # Resets status bar and enables window

    def _sanitize_sheet_title(self, title: str) -> str:
        """Removes invalid characters from Excel sheet titles."""
        invalid_chars = set(r'[]:*?/\\')
        sanitized = "".join(ch for ch in title if ch not in invalid_chars)
        return sanitized or "Sheet1"

    def _load_config(self) -> None:
        """Loads configuration from the config.yaml file."""
        try:
            with open("config.yaml", "r", encoding="utf-8") as file:
                config = yaml.safe_load(file)
        except FileNotFoundError:
            self.show_notification.emit("error", "Файл config.yaml не найден.")
            return

        if not isinstance(config, dict) or "path_to_products_folder" not in config:
            self.show_notification.emit(
                "error",
                "Файл config.yaml не содержит обязательных полей.",
            )
            return

        # Store configuration data
        self.program_version_number = config.get("program_version_number", "")
        self.program_server_path = config.get("server_program_path", "")

        path = config["path_to_products_folder"]
        if os.path.exists(path) and os.path.isdir(path):
            self.path_to_products_folder = path
            self.is_products_folder_available = True
        else:
            self.show_notification.emit(
                "error",
                "Путь к папке с продуктами недоступен.",
            )
            self.is_products_folder_available = False
    def _normalize_material_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Renames material columns to standard keys based on fuzzy matching."""
        columns_lower = {col: str(col).lower() for col in df.columns}

        def find_column(substrings: list[str]) -> str | None:
            for col, low in columns_lower.items():
                if all(sub in low for sub in substrings):
                    return col
            return None

        mapping = {}
        candidates = {
            NOM_KEY: ["номенк"],
            QTY_KEY: ["кол"],
            UNIT_KEY: ["ед", "изм"],
        }

        for target, substrings in candidates.items():
            col = find_column(substrings)
            if col:
                mapping[col] = target

        df = df.rename(columns=mapping)

        missing = [key for key in (NOM_KEY, QTY_KEY, UNIT_KEY) if key not in df.columns]
        if missing:
            raise ValueError(f"Usecols do not match columns, missing: {missing}, got: {list(df.columns)}")

        return df

    def _normalize_material_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Renames material columns to standard keys based on fuzzy matching (clean override)."""
        columns_lower = {col: str(col).lower() for col in df.columns}

        def find_column(substrings: list[str]) -> str | None:
            for col, low in columns_lower.items():
                if all(sub in low for sub in substrings):
                    return col
            return None

        mapping = {}
        candidates = {
            NOM_KEY: ["номенк"],
            QTY_KEY: ["кол"],
            UNIT_KEY: ["ед", "изм"],
        }

        for target, substrings in candidates.items():
            col = find_column(substrings)
            if col:
                mapping[col] = target

        df = df.rename(columns=mapping)

        missing = [key for key in (NOM_KEY, QTY_KEY, UNIT_KEY) if key not in df.columns]
        if missing:
            raise ValueError(
                f"Usecols do not match columns, missing: {missing}, got: {list(df.columns)}"
            )

        return df

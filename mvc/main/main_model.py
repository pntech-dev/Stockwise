import os
import yaml
import subprocess
import pandas as pd

from pathlib import Path
from openpyxl import load_workbook
from PyQt5.QtCore import QObject, pyqtSignal
from openpyxl.styles import Alignment, Border, Font, Side


class MainModel(QObject):
    """Manages the application's data and business logic.

    This class handles loading configuration, finding and processing product data,
    checking for updates, and exporting data to Excel. It communicates with
    the controller/view using signals.

    Attributes:
        show_notification: A signal that emits messages to be displayed to the user.
    """

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
        self.norms_calculations_value: int = 1
        self.search_in_materials: bool = False
        self.search_in_materials_data: list[dict] = []

    def get_semi_finished_products(self, product_name: tuple) -> list[str]:
        """Gets a list of semi-finished product files for a given product.

        Args:
            product_name: A tuple representing the path components of the product.

        Returns:
            A list of full paths to the semi-finished product Excel files.
        """
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
                    if item_name.lower() == "рмп" and os.path.isdir(item_path):
                        for rmp_file_name in os.listdir(item_path):
                            rmp_file_path = os.path.join(item_path, rmp_file_name)
                            add_file(rmp_file_path, semi_finished_products)
                    else:
                        add_file(item_path, semi_finished_products)
            return semi_finished_products
        except Exception as e:
            self.show_notification.emit(
                "error", f"Ошибка при получении списка полуфабрикатов: {e}"
            )
            return []

    def get_product_materials(
        self, semi_finished_products: list[str]
    ) -> list[dict]:
        """Aggregates all materials from a list of semi-finished product files.

        Reads each Excel file, calculates material quantities based on the norm,
        and aggregates identical materials by summing their quantities.

        Args:
            semi_finished_products: A list of paths to semi-finished product files.

        Returns:
            A list of dictionaries, where each dictionary represents a unique material.
        """
        product_materials_dict = {}
        semi_product_names = [
            os.path.splitext(os.path.basename(p))[0].lower().strip()
            for p in semi_finished_products
        ]

        for file_path in semi_finished_products:
            if not os.path.exists(file_path):
                self.show_notification.emit("error", f"Файл не найден: {file_path}")
                continue

            try:
                is_rmp = (
                    os.path.basename(os.path.dirname(file_path)).lower() == "рмп"
                )
                df = pd.read_excel(
                    file_path,
                    sheet_name="TDSheet",
                    usecols=["Номенклатура", "Количество", "Ед. изм."],
                )
                df = df.dropna(subset=["Номенклатура"])
                df = df.fillna("")

                for item in df.to_dict("records"):
                    nomenclature = item["Номенклатура"]
                    # Skip materials that are themselves semi-finished products
                    if nomenclature.lower().strip() in semi_product_names:
                        continue

                    # Calculate quantity based on norm
                    quantity = item["Количество"] / 1000
                    if self.norms_calculations_value != 1:
                        quantity *= self.norms_calculations_value
                    item["Количество"] = quantity
                    item["РМП"] = is_rmp

                    # Aggregate materials
                    if nomenclature in product_materials_dict:
                        product_materials_dict[nomenclature]["Количество"] += quantity
                    else:
                        product_materials_dict[nomenclature] = item

            except Exception as e:
                self.show_notification.emit(
                    "error", f"Ошибка при чтении файла {os.path.basename(file_path)}: {e}"
                )

        return list(product_materials_dict.values())

    def get_desktop_path(self) -> Path | None:
        """Finds the path to the user's desktop folder.

        Checks for common OneDrive and standard desktop paths.

        Returns:
            A Path object to the desktop, or None if an error occurs.
        """
        try:
            # Check for Russian OneDrive path
            onedrive_ru = Path.home() / "OneDrive" / "Рабочий стол"
            if onedrive_ru.exists():
                return onedrive_ru

            # Check for English OneDrive path
            onedrive_en = Path.home() / "OneDrive" / "Desktop"
            if onedrive_en.exists():
                return onedrive_en

            # Fallback to standard Desktop path
            return Path.home() / "Desktop"
        except Exception as e:
            self.show_notification.emit("error", f"Ошибка при получении пути к рабочему столу: {e}")
            return None

    def check_program_version(self) -> bool | None:
        """Checks if the local program version is up-to-date with the server.

        Returns:
            True if the version is current or newer.
            False if an update is needed.
            None if an error occurred (e.g., server path not found).
        """
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
            self.show_notification.emit("error", "Папка с продукцией недоступна на сервере.")
            return

        products = []
        try:
            for root, dirs, _ in os.walk(self.path_to_products_folder):
                # Exclude the 'rmp' directory from the walk
                if "рмп" in [d.lower() for d in dirs]:
                    dirs.remove("рмп" if "рмп" in dirs else "РМП")

                # A directory is considered a product if it has no subdirectories
                if not dirs:
                    relative_path = os.path.relpath(root, self.path_to_products_folder)
                    if relative_path != ".":
                        products.append(relative_path.split(os.sep))
            self.products_names = products
        except Exception as e:
            self.show_notification.emit(
                "error", f"Ошибка при обновлении списка наименований продукции: {e}"
            )

    def update_program(self) -> None:
        """Launches the updater executable with administrator privileges."""
        updater_path = os.path.join(os.getcwd(), "updater.exe")
        if not os.path.exists(updater_path):
            self.show_notification.emit("error", f"Программа обновления не найдена: {updater_path}")
            return

        try:
            # Use PowerShell to run the updater with admin rights
            command = f'Start-Process "{updater_path}" -ArgumentList "{self.program_server_path}" -Verb RunAs'
            subprocess.run(["powershell", "-Command", command], check=True, shell=True)
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            self.show_notification.emit("error", f"Не удалось запустить программу обновления: {e}")

    def export_data(self) -> None:
        """Exports the current materials list to a styled Excel file."""
        try:
            workbook = load_workbook("templates/table.xlsx")
            sheet = workbook.active
            sheet.title = "Список материалов"

            data = (
                self.search_in_materials_data
                if self.search_in_materials
                else self.current_product_materials
            )
            
            if not data:
                self.show_notification.emit("warning", "Нет данных для экспорта.")
                return

            # --- Fill data ---
            start_row = 2
            for i, item in enumerate(data):
                row_idx = start_row + i
                sheet.cell(row=row_idx, column=1, value=item["Номенклатура"])
                sheet.cell(row=row_idx, column=2, value=item["Ед. изм."])
                sheet.cell(row=row_idx, column=3, value=item["Количество"])

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
                    "error", "Не удалось определить путь к рабочему столу для сохранения файла."
                )
                return

            file_name = f"Список материалов для {self.current_product}.xlsx"
            file_path = desktop_path / file_name
            workbook.save(file_path)

            self.show_notification.emit("info", f"Файл сохранен на рабочем столе: {file_name}")

        except Exception as e:
            self.show_notification.emit(
                "error", f"Произошла ошибка при создании списка материалов: {e}"
            )
        finally:
            self.show_notification.emit("", "") # Resets status bar and enables window

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
                "В файле config.yaml не указан путь к папке с продукцией.",
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
                "Указанный путь к папке с продукцией не существует или не является каталогом.",
            )
            self.is_products_folder_available = False
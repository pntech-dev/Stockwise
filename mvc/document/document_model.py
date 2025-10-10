import os
import yaml
import threading

from PyQt5.QtCore import QDate

# Excel экспорт
from copy import copy
from jinja2 import Template
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class DocumentModel:
    def __init__(self, product_name, norms_calculations_value, materials, current_product_path):
        # Данные из основного окна приложения
        self.product_name = product_name
        self.materials = materials
        self.current_product_path = current_product_path

        # Данные из файла конфигурации
        self.signature_from_human = [] # Подпись от кого
        self.signature_from_position = [] # Должность от кого
        self.signature_whom_human = [] # Подпсиь кому
        self.signature_whom_position = [] # Должность кому

        # Данные для подстановки в документ
        self.outgoing_number = "" # Номер исходящего документа
        self.current_date = "" # Текущая дата
        self.product_name = self.get_current_product_name() # Определяем имя текущего изделия
        self.quantity = norms_calculations_value
        self.whom_position = "" # Должность кому
        self.whom_fio = "" # ФИО кому
        self.from_position = "" # Должность от кого
        self.from_fio = "" # ФИО от кого

        # Black и White листы
        self.bid_blacklist = []
        self.document_blacklist = []

        self.__load_config() # Загружаем конфигурацию при инициализации

    def __load_config(self):
        """Функция загружает конфигурацию из файла config.yaml."""
        try:
            with open("config.yaml", "r", encoding="utf-8") as file:
                config = yaml.safe_load(file)
        except FileNotFoundError:
            print("Ошибка, Файл config.yaml не найден.")
            # self.show_notification.emit("error", "Файл config.yaml не найден.")
            return
        
        if "path_to_products_folder" not in config:
            print("Ошибка, В файле config.yaml не указан путь к папке изделий.")
            # self.show_notification.emit("error", "В файле config.yaml не указан путь к папке изделий.")
            return

        # Записываем данные из файла кофигурации
        self.signature_from_human = config["signature_from_human"]
        self.signature_from_position = config["signature_from_position"]
        self.signature_whom_human = config["signature_whom_human"]
        self.signature_whom_position = config["signature_whom_position"]

        path = config["path_to_products_folder"]
        if os.path.exists(path) and os.path.isdir(path):
            self.path_to_products_folder = path
            self.is_products_folder_available = True
        else:
            print("Ошибка, Указанный путь к папке изделий не существует или не является папкой.")
            # self.show_notification.emit("error", "Указанный путь к папке изделий не существует или не является папкой.")
            return
        
        # Получаем блэклист заявки
        bid_blacklist = config.get("bid_blacklist")
        if bid_blacklist:
            self.bid_blacklist = bid_blacklist
        else:
            print("Ошибка при чтении блэклиста заявок")
            return
        
        # Получаем блэклист докладной
        document_blacklist = config.get("document_blacklist")
        if document_blacklist:
            self.document_blacklist = document_blacklist
        else:
            print("Ошибка при чтении блэклиста докладной записки")
            return

    def __get_document_materials_list(self):
        """Функция возвращает список материалов для докладной записки."""
        self.current_materials = []

        for item in self.materials:
            if item['Ед. изм.'] != "шт" and not item["РМП"]:
                # Проверяем находиться ли материал в блэклисте
                in_blacklist = False
                if any(word.lower() in item['Номенклатура'].lower() for word in self.document_blacklist):
                    in_blacklist = True

                # Если материал не в блэклисте, добавляем его
                if not in_blacklist:
                    self.current_materials.append(item)

        return self.current_materials
    
    def __get_bid_materials_list(self):
        """Функция возвращает список материалов для заявки."""
        self.current_materials = []

        for item in self.materials:
            if item['Ед. изм.'] == "шт":
                # Проверяем находиться ли материал в блэклисте
                in_blacklist = False
                if any(word.lower() in item['Номенклатура'].lower() for word in self.bid_blacklist):
                    in_blacklist = True

                # Если материал не в блэклисте, добавляем его
                if not in_blacklist:
                    self.current_materials.append(item)

        return self.current_materials
    
    def __export_materials_list(self, workbook, materials_list):
        """Функция экспортирует список материалов на новую страницу в Excel."""
        wb_template = load_workbook("templates/table.xlsx")
        template_sheet = wb_template.active

        new_sheet = workbook.create_sheet(title="Перечень материалов")

        # Копируем значения и стили безопасно
        for row in template_sheet.iter_rows():
            for cell in row:
                new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)

                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)

                if cell.hyperlink:
                    new_cell.hyperlink = copy(cell.hyperlink)
                if cell.comment:
                    new_cell.comment = copy(cell.comment)

        # Копируем ширину колонок
        for col_letter, col_dim in template_sheet.column_dimensions.items():
            new_sheet.column_dimensions[col_letter].width = col_dim.width

        # Копируем высоту строк
        for row_idx, row_dim in template_sheet.row_dimensions.items():
            new_sheet.row_dimensions[row_idx].height = row_dim.height
    
        # Заполняем материалы
        start_row = 2
        for i, item in enumerate(materials_list):
            row = start_row + i
            new_sheet.cell(row=row, column=1, value=item['Номенклатура'])
            new_sheet.cell(row=row, column=2, value=item['Ед. изм.'])
            new_sheet.cell(row=row, column=3, value=item['Количество'])

        # Общие настройки шрифта и перенос текста
        font = Font(name="Times New Roman", size=14)

        for row in new_sheet.iter_rows(min_row=start_row, max_row=new_sheet.max_row, max_col=3):
            for cell in row:
                cell.font = font
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        # Отдельно задаём выравнивание текста для каждой колонки
        for row in range(start_row, new_sheet.max_row + 1):
            # 1-я колонка — слева сверху
            new_sheet.cell(row=row, column=1).alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            # 2-я и 3-я — по центру сверху
            for col in [2, 3]:
                new_sheet.cell(row=row, column=col).alignment = Alignment(
                    horizontal="center", vertical="top", wrap_text=True
                )

        # Настраиваем рамки таблицы 
        thick = Side(border_style="thick", color="000000")  # Жирная линия
        thin = Side(border_style="thin", color="000000")    # Тонкая линия

        for row in range(start_row, new_sheet.max_row + 1):
            for col in [1, 2, 3]:
                cell = new_sheet.cell(row=row, column=col)

                # Задаём рамки по умолчанию
                left = thick if col == 1 else thin
                right = thick if col == 3 else thin
                top = thin
                bottom = thin

                # Применяем границы
                cell.border = Border(left=left, right=right, top=top, bottom=bottom)

        # Настраиваем автоматическое выравнивание
        new_sheet.page_setup.fitToWidth = 1 # По высоте строки
        new_sheet.page_setup.fitToHeight = 0 # По ширине столбца

        return new_sheet
        
    def __export_to_excel(self, document_type, save_folder_path):
        """Функция обрабатывает экспорт в Excel."""
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

        # Если документ - Докладная записка (Цех)
        if document_type == "document":
            # Страница докладной записки
            wb = load_workbook("templates/document.xlsx")
            ws = wb.active

            ws.title = "Докладная"

            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and "{{" in cell.value:
                        template = Template(cell.value)
                        cell.value = template.render(context)

            # Страница перечня материалов
            materials_list = self.__get_document_materials_list()
            self.__export_materials_list(workbook=wb, materials_list=materials_list)

            # Сохраняем документ
            wb.save(os.path.join(save_folder_path, f"Докладная {self.product_name}.xlsx"))

        # Если документ - Заявка (ПДС)
        elif document_type == "bid":
            wb = load_workbook("templates/bid.xlsx")
            ws = wb.active
            ws.title = "Заявка"

            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and "{{" in cell.value:
                        template = Template(cell.value)
                        cell.value = template.render(context)

            # Страница перечня материалов
            materials_list = self.__get_bid_materials_list()
            self.__export_materials_list(workbook=wb, materials_list=materials_list)

            wb.save(os.path.join(save_folder_path, f"Заявка {self.product_name}.xlsx"))

    def get_current_date(self):
        """Функция возвращает текущую дату."""
        return QDate.currentDate()
    
    def get_current_product_name(self):
        """Функция возвращает название текущего изделия."""
        # Проверяем что путь к папке изделия существует и является папкой
        if os.path.exists(self.current_product_path) and os.path.isdir(self.current_product_path):
            files = os.listdir(self.current_product_path) # Получаем список всех файлов в папке
            for file in files:
                # Если файл и имя файла заканчивается на .xlsx или .xls
                if os.path.isfile(os.path.join(self.current_product_path, file)) and file.lower().endswith(('.xlsx', '.xls')):
                    if any(product_name_word.lower() in file.lower().strip() for product_name_word in self.product_name.strip().split()):
                        return os.path.splitext(file)[0]
                    
        return self.product_name
    
    def export_in_thread(self, document_type, save_folder_path):
        """Функция запускает процесс экспорта документа в отдельном потоке."""
        if len(save_folder_path) == 0:
            try:
                save_folder_path = os.path.join(os.path.expanduser("~"), "Desktop")
            except Exception as e:
                print(f"Произошла ошибка при получении пути к папке Desktop: {e}")
                return

        try:
            thread = threading.Thread(target=self.__export_to_excel, args=(document_type, save_folder_path,))
            thread.daemon = True
            thread.start()
            return thread
        
        except Exception as e:
            print(f"Произошла ошибка во время запуска потока экспорта документа: {e}")
            return
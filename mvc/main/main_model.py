import os
import yaml
import subprocess
import pandas as pd

from PyQt5.QtCore import QObject, pyqtSignal

from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

class MainModel(QObject):
    show_notification = pyqtSignal(str, str) # Сигнал показа уведомления

    def __init__(self):
        super().__init__()

        self.program_version_number = None # Версия программы
        self.program_server_path = None # Путь к программе на сервере
        self.path_to_products_folder = None
        self.is_products_folder_available = False
        self.__load_config()  # Загружаем конфигурацию при инициализации

        self.products_names = []  # Список названий изделий
        self.update_products_names()  # Обновляем список названий изделий

        self.current_product = ""  # Текущее выбранное изделие
        self.current_product_path = ""  # Путь к текущему продукту
        self.current_product_materials = []  # Список материалов текущего изделия

        self.norms_calculations_value = 1 # Значение на которое рассчитываются нормы изделия

    def __load_config(self):
        """Функция загружает конфигурацию из файла config.yaml."""
        try:
            with open("config.yaml", "r", encoding="utf-8") as file:
                config = yaml.safe_load(file)
        except FileNotFoundError:
            self.show_notification.emit("error", "Файл config.yaml не найден.")
            return
        
        if "path_to_products_folder" not in config:
            self.show_notification.emit("error", "В файле config.yaml не указан путь к папке изделий.")
            return

        # Записываем данные из файла кофигурации
        self.program_version_number = config["program_version_number"]
        self.program_server_path = config["server_program_path"]

        path = config["path_to_products_folder"]
        if os.path.exists(path) and os.path.isdir(path):
            self.path_to_products_folder = path
            self.is_products_folder_available = True
        else:
            self.show_notification.emit("error", "Указанный путь к папке изделий не существует или не является папкой.")
            return

    def get_semi_finished_products(self, product_name):
        """Функция возвращает список полуфабрикатов входящих в продукт."""
        def add_file(file_path, lst):
            """Функция добавляет файл в список полуфабрикатов."""
            if os.path.isfile(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
                lst.append(file_path)

        try:
            semi_finished_products = [] # Создаем пустой список полуфабрикатов
            # Формируем путь к папке изделия
            self.product_path = os.path.join(self.path_to_products_folder, *product_name)

            # Если путь к папке изделия существует
            if os.path.exists(self.product_path) and os.path.isdir(self.product_path):
                self.current_product_path = self.product_path

                # Получаем список файлов и папок в папке изделия
                for item_name in os.listdir(self.product_path):
                    item_path = os.path.join(self.product_path, item_name) # Формируем путь к элементу

                    # Проверка на папку "рмп"
                    if item_name.lower() == "рмп" and os.path.isdir(item_path):
                        for rmp_file_name in os.listdir(item_path):
                            rmp_file_path = os.path.join(item_path, rmp_file_name)

                            add_file(rmp_file_path, semi_finished_products)

                    else:
                        add_file(item_path, semi_finished_products)

            return semi_finished_products
        
        except:
            self.show_notification.emit("error", "Ошибка при получении списка полуфабрикатов.")
            return []

    def get_product_materials(self, semi_finished_products):
        """Функция возвращает список материалов входящих в продукт."""
        product_materials_dict = {} # Используем словарь для быстрой проверки существования

        for semi_finished_product in semi_finished_products:
            if os.path.exists(semi_finished_product): # Если полуфабрикат существует
                try:
                    is_rmp_folder = os.path.basename(os.path.dirname(semi_finished_product)).lower() == "рмп"

                    # Определяем колонки для чтения
                    columns_to_read = ['Номенклатура', 'Количество', 'Ед. изм.']
                    # Читаем только нужные колонки из Excel файла
                    df = pd.read_excel(semi_finished_product, sheet_name='TDSheet', usecols=columns_to_read)
                    df = df.fillna('') # Заменяем все NaN значения на пустые строки

                    # Преобразуем DataFrame в список словарей
                    data_list = df.to_dict('records')

                    for new_item in data_list:
                        new_item['РМП'] = is_rmp_folder # Отмечаем какие материалы из папки РМП

                        # Рассчитываем норму на еденицу по умолчанию
                        new_item['Количество'] = new_item['Количество'] / 1000

                        # Если заданное количество изменилось, рассчитываем норму материалов
                        if self.norms_calculations_value is not None and self.norms_calculations_value != 1:
                            new_item['Количество'] = new_item['Количество'] * self.norms_calculations_value

                        nomenclature = new_item['Номенклатура']
                        product_name = os.path.splitext(os.path.basename(semi_finished_product))[0]
                        
                        # Проверяем находиться ли материал в спсике полуфабрикатов
                        is_product = False
                        semi_products_names = [os.path.splitext(os.path.basename(product))[0] for product in semi_finished_products]
                        for product in semi_products_names:
                            if product.lower().strip() in nomenclature.lower().strip():
                                is_product = True
                                break
                        
                        if is_product:
                            continue
                        else:
                            # Проверяем, есть ли материал в словаре
                            if nomenclature in product_materials_dict:
                                existing_item = product_materials_dict[nomenclature]
                                # Увеличиваем количество
                                existing_item['Количество'] += new_item['Количество']
                            else:
                                product_materials_dict[nomenclature] = new_item

                except Exception as e:
                    self.show_notification.emit("error", f"Ошибка при чтении файла {semi_finished_product}: {e}")
            else:
                self.show_notification.emit("error", f"Файл {semi_finished_product} не найден.")

        return list(product_materials_dict.values())
    
    def get_desktop_path(self):
        """Функция возвращает путь к папке Desktop"""
        try:
            # Проверяем сначала путь к рабочему столу в OneDrive с русским названием
            onedrive_desktop_path = Path.home() / "OneDrive" / "Рабочий стол"
            if onedrive_desktop_path.exists():
                return onedrive_desktop_path

            # Проверяем снова путь к рабочему столу в OneDrive
            onedrive_desktop_path = Path.home() / "OneDrive" / "Desktop"
            if onedrive_desktop_path.exists():
                return onedrive_desktop_path

            # Если не найден, используем стандартный путь
            desktop_path = Path.home() / "Desktop"
            return desktop_path

        except Exception as e:
            self.show_notification.emit("error", f"Произошла ошибка при получении пути к папке Desktop.\nОшибка: {e}")
            return

    def check_program_version(self):
        """Функция проверяет версию программы"""
        if not self.program_server_path or not os.path.exists(self.program_server_path):
            return None
        
        program_server_files = os.listdir(self.program_server_path)
        for file_name in program_server_files:
            # Ищем файл, который начинается с "config" и заканчивается ".yaml"
            if file_name.startswith("config") and file_name.endswith(".yaml"):
                try:
                    with open(os.path.join(self.program_server_path, file_name), "r", encoding="utf-8") as f:
                        config_data = yaml.safe_load(f)

                    program_server_version = config_data["program_version_number"] # Получаем версию программы на сервере
                    
                    if not program_server_version:
                        return None

                    if program_server_version <= self.program_version_number:
                        return True
                except IndexError:
                    # Если формат имени файла не соответствует ожидаемому, пропускаем его
                    continue
        else:
            return False

    def update_products_names(self):
        """Функция обновляет список названий изделий, рекурсивно обходя папку."""
        products_names = []  # Создаем пустой список названий изделий

        if not self.is_products_folder_available:  # Если папка изделий недоступна
            self.products_names = products_names  # Устанавливаем пустой список названий изделий
            self.show_notification.emit("error", "Папка изделий недоступна на сервере.")
        try:
            for root, dirs, files in os.walk(self.path_to_products_folder):  # Рекурсивно обходим папку изделий
                # Определяем, какие из подпапок являются папками продуктов (т.е. не "рмп")
                product_subdirs = [d for d in dirs if d.lower() != 'рмп']

                # Исключаем папку "рмп" из дальнейшего обхода os.walk
                # Модификация dirs влияет на последующие итерации os.walk
                if 'рмп' in [d.lower() for d in dirs]:
                    dirs.remove('рмп' if 'рмп' in dirs else 'РМП')

                # Если в текущей папке нет вложенных папок-продуктов, то она сама является продуктом
                if not product_subdirs:
                    relative_path = os.path.relpath(
                        root, self.path_to_products_folder)

                    if relative_path != '.':
                        products_names.append(relative_path.split(os.sep))

            self.products_names = products_names
        except Exception as e:
            self.show_notification.emit("error", f"Ошибка при обновлении списка названий изделий: {e}")

    def update_program(self):
        """Функция вызывает обновление программы"""
        updater_path = os.path.join(os.getcwd(), "updater.exe") # Создаём путь к программе обновления
        
        # Проверяем, что updater.exe существует
        if not os.path.exists(updater_path):
            self.show_notification.emit("error", f"Не найден файл обновления: {updater_path}")
            return

        # Запускаем updater.exe с запросом прав администратора
        # Используем 'runas' для повышения прав
        subprocess.run(['powershell', '-Command', f'Start-Process "{updater_path}" -ArgumentList "{self.program_server_path}" -Verb RunAs'], shell=True)

    def export_data(self):
        """Функция экспортирует данные в Excel"""
        try:
            # Загружаем шаблон
            workbook = load_workbook("templates/table.xlsx")
            new_sheet = workbook.active
            new_sheet.title = "Перечень материалов"

            # Заполняем материалы
            start_row = 2
            for i, item in enumerate(self.current_product_materials):
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

            # Сохраняем файл на рабочем столе
            desktop_path = self.get_desktop_path()
            if not desktop_path:
                self.show_notification.emit("error", f"Произошла ошибка во время получения пути рабочего стола.\nОшибка: {e}")
                return
            
            file_name = f"Перечень материалов для {''.join(self.current_product)}.xlsx"
            file_path = os.path.join(desktop_path, file_name)
            
            workbook.save(file_path)
            
            self.show_notification.emit("info", f"Файл сохранен на рабочем столе: {file_name}")

        except Exception as e:
            self.show_notification.emit("error", f"Произошла ошибка во время создания перечня материалов\nОшибка: {e}")
            return
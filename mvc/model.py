import os
import yaml
import pandas as pd


class Model:
    def __init__(self):
        self.path_to_products_folder = None
        self.is_products_folder_available = False
        self.__load_config()  # Загружаем конфигурацию при инициализации

        self.products_names = []  # Список названий изделий
        self.update_product_names()  # Обновляем список названий изделий

        self.current_product = ""  # Текущее выбранное изделие
        self.current_product_materials = []  # Список материалов текущего изделия

    def __load_config(self):
        """Функция загружает конфигурацию из файла config.yaml."""
        try:
            with open("config.yaml", "r", encoding="utf-8") as file:
                config = yaml.safe_load(file)
        except FileNotFoundError:
            print("Файл config.yaml не найден.")
            return

        if "path_to_products_folder" not in config:
            print("В файле config.yaml не указан путь к папке изделий.")
            return

        path = config["path_to_products_folder"]
        if os.path.exists(path) and os.path.isdir(path):
            self.path_to_products_folder = path
            self.is_products_folder_available = True
        else:
            print(f"Путь '{path}' не существует или не является папкой.")

    def __export_to_excel(self, save_path, data):
        """Функция экспортирует данные в Excel файл."""
        df = pd.DataFrame(data)
        df.to_excel(save_path, index=False)

    def __export_to_word(self, save_path, data):
        """Функция экспортирует данные в Word файл."""
        pass

    def __export_to_pdf(self, save_path, data):
        """Функция экспортирует данные в PDF файл."""
        pass

    def get_semi_finished_products(self, product_name):
        """Функция возвращает список полуфабрикатов входящих в продукт."""
        semi_finished_products = []  # Создаем пустой список полуфабрикатов
        # Формируем путь к папке изделия
        self.product_path = os.path.join(
            self.path_to_products_folder, *product_name)

        # Если путь к папке изделия существует
        if os.path.exists(self.product_path) and os.path.isdir(self.product_path):
            # Получаем список файлов и папок в папке изделия
            for item_name in os.listdir(self.product_path):
                item_path = os.path.join(
                    self.product_path, item_name)  # Формируем путь к элементу

                # Если это папка "рмп", обрабатываем ее содержимое
                if item_name.lower() == "рмп" and os.path.isdir(item_path):
                    # Получаем список файлов в папке "рмп"
                    for rmp_file_name in os.listdir(item_path):
                        rmp_file_path = os.path.join(item_path, rmp_file_name)
                        # Проверяем, что это файл Excel
                        if os.path.isfile(rmp_file_path) and rmp_file_path.lower().endswith(('.xlsx', '.xls')):
                            semi_finished_products.append(rmp_file_path)
                # Иначе, если это файл Excel
                elif os.path.isfile(item_path) and item_path.lower().endswith(('.xlsx', '.xls')):
                    semi_finished_products.append(item_path)

        return semi_finished_products

    def get_product_materials(self, semi_finished_products):
        """Функция возвращает список материалов входящих в продукт. (Оптимизировано)"""
        product_materials_dict = {}  # Используем словарь для быстрой проверки существования

        for semi_finished_product in semi_finished_products:
            if os.path.exists(semi_finished_product):  # Если полуфабрикат существует
                try:
                    # Определяем колонки для чтения
                    columns_to_read = ['Номенклатура', 'Количество', 'Ед. изм.']
                    # Читаем только нужные колонки из Excel файла
                    df = pd.read_excel(semi_finished_product, sheet_name='TDSheet', usecols=columns_to_read)
                    df = df.fillna('')  # Заменяем все NaN значения на пустые строки

                    # Преобразуем DataFrame в список словарей
                    data_list = df.to_dict('records')

                    for new_item in data_list:
                        nomenclature = new_item['Номенклатура']
                        product_name = os.path.splitext(
                            os.path.basename(semi_finished_product))[0]

                        # Если материал уже есть в словаре
                        if nomenclature in product_materials_dict:
                            existing_item = product_materials_dict[nomenclature]
                            # Увеличиваем количество
                            existing_item['Количество'] += new_item['Количество']

                            # Если название изделия не найдено в списке изделий
                            if product_name not in existing_item['Изделие']:
                                # Добавляем название изделия в список изделий
                                existing_item['Изделие'].append(product_name)
                        else:
                            # Если материал не найден, добавляем его в словарь
                            new_item['Изделие'] = [product_name]
                            product_materials_dict[nomenclature] = new_item
                except Exception as e:
                    print(f"Ошибка при чтении файла {semi_finished_product}: {e}")
            else:
                print(f"Файл {semi_finished_product} не найден.")

        return list(product_materials_dict.values())
    
    def get_desktop_path(self):
        """Функция возвращает путь к папке Desktop."""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        return desktop_path

    def update_product_names(self):
        """Функция обновляет список названий изделий, рекурсивно обходя папку."""
        products_names = []  # Создаем пустой список названий изделий

        if not self.is_products_folder_available:  # Если папка изделий недоступна
            self.products_names = products_names  # Устанавливаем пустой список названий изделий
            return

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

    def export_data(self, save_path, export_format, data):
        """Функция экспортирует данные в указанный формат."""
        if export_format is None:
            print("Не выбран формат экспорта.")
            return
        
        if not data:
            print("Нет данных для экспорта.")
            return
        
        save_file_path = os.path.join(save_path, f"{self.current_product}.{export_format}")

        if export_format == "xlsx":
            self.__export_to_excel(save_path=save_file_path, data=data)
        elif export_format == "docx":
            self.__export_to_word(save_path=save_file_path, data=data)
        elif export_format == "pdf":
            self.__export_to_pdf(save_path=save_file_path, data=data)
        else:
            print("Неподдерживаемый формат экспорта.")
            return
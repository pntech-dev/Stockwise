import os
import yaml

from PyQt5.QtCore import QDate


class DocumentModel:
    def __init__(self, product_name, norms_calculations_value, materials, current_product_path):
        # Данные из основного окна приложения
        self.product_name = product_name
        self.quantity = norms_calculations_value
        self.materials = materials
        self.current_product_path = current_product_path

        # Данные из файла конфигурации
        self.signature_from_human = [] # Подпись от кого
        self.signature_from_position = [] # Должность от кого
        self.signature_whom_human = [] # Подпсиь кому
        self.signature_whom_position = [] # Должность кому
        self.__load_config() # Загружаем конфигурацию при инициализации

        # Определяем имя текущего изделия
        self.current_product_name = self.get_current_product_name()

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
    
    def export_in_thread(self, document_type, save_folder_path, export_format="excel"):
        """Функция запускает процесс экспорта документа в отдельном потоке."""
        if len(save_folder_path) == 0:
            try:
                save_folder_path = os.path.join(os.path.expanduser("~"), "Desktop")
            except Exception as e:
                print(f"Произошла ошибка при получении пути к папке Desktop: {e}")
                return

        if document_type[0]:
            print("Экспорт докладной записки")
            print(f"Формат экспорта: {export_format}")
            print(f"Путь к папке сохранения: {save_folder_path}")
        elif document_type[1]:
            print("Экспорт заявки")
            print(f"Формат экспорта: {export_format}")
            print(f"Путь к папке сохранения: {save_folder_path}")
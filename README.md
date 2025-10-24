# Stockwise (English Version)

A simple and powerful desktop application for calculating and consolidating bill of materials, and generating production documents.

## Description

**Stockwise** is a desktop application designed for engineers, technologists, and planning specialists. It automates the process of gathering and calculating material requirements for manufacturing various products, and generates formatted documents like memos and requests.

The application scans a structured network directory where each product is represented by a folder hierarchy, and its bill ofmaterials is described in Excel files. Stockwise reads these files, sums up identical material items, allows for dynamic recalculation of quantities for a given production volume, and exports the final report into a convenient format.

The new version introduces the ability to generate two types of documents: "Memos" and "Requests", with customizable fields for signatures and separate blacklists and whitelists for materials.

## Key Features

-   **Hierarchical Search**: Smart product search with support for partial matches and multiple keywords.
-   **Automatic Consolidation**: Gathers data from multiple Excel files, including nested specification folders.
-   **Material Aggregation**: Automatically combines and sums quantities for identical nomenclature items.
-   **Dynamic Norm Calculation**: Ability to calculate required material quantities for any production volume.
-   **Document Generation**: Create "Memos" and "Requests" with pre-filled data.
-   **Export to 3 Formats**: Exports the final specification to **Excel (.xlsx)**, **Word (.docx)**, and **PDF**, with pre-configured formatting for A4 printing.
-   **Auto-Update**: The application automatically checks for a new version on the server and prompts for an update.
-   **Configurable Blacklists/Whitelists**: Exclude or include specific materials from the final calculation for different document types.
-   **Configurable Signatures**: Customize the "from" and "to" fields in the generated documents.
-   **Modern UI**: A clean and responsive user interface built with PyQt5.

## How It Works

For the application to function correctly, the data must be stored in a specific structure within the folder specified in `config.yaml`.

1.  **Root Folder**: The main folder containing all product data.
2.  **Product Folder**: Each product is a folder hierarchy. The product name is derived from the path to the final folder.
    -   *Example*: The path `\SRV\PNDatabases\Stockwise\БВВ\01` will be interpreted as the product `БВВ 01`.
3.  **Specification Files**: Inside the final product folder there should be Excel files (`.xlsx`, `.xls`) containing the list of materials.
4.  **Excel File Structure**: The application expects a sheet named `TDSheet` to contain at least three columns: `Номенклатура` (Nomenclature), `Количество` (Quantity), and `Ед. изм.` (Unit).
5.  **Configuration**: The main settings are in `config.yaml`:
    - `path_to_products_folder`: Path to the root folder with product data.
    - `server_program_path`: Path to the server folder for checking updates.
    - `signature_from_human`, `signature_from_position`, `signature_whom_human`, `signature_whom_position`: Lists of names and positions for the "from" and "to" fields in the documents.
    - `document_blacklist`, `bid_blacklist`: Blacklists for materials for "Memos" and "Requests".
    - `document_whitelist`, `bid_whitelist`: Whitelists for materials for "Memos" and "Requests".

## Installation and Launch

To run the project from the source code, follow these steps:

1.  **Clone the repository:**
    ```bash
    git clone <repository URL>
    cd Stockwise
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure the application:**
    -   Open the `config.yaml` file.
    -   Set the correct path to the network folder with products in the `path_to_products_folder` parameter.
    -   Specify the path to the folder with the current version of the program on the server in `server_program_path`.
    -   Customize signatures and blacklists/whitelists as needed.

5.  **Run the application:**
    ```bash
    python app.py
    ```

## Technology Stack

-   **Language**: Python 3.10
-   **GUI**: PyQt5
-   **Data Processing**: Pandas
-   **Exporting**: Excel: `openpyxl`
-   **Architecture**: MVC (Model-View-Controller)

---

# Stockwise

Простое и мощное настольное приложение для расчёта и консолидации спецификаций материалов, а также создания производственных документов.

## Описание

**Stockwise** — это десктопное приложение, разработанное для инженеров, технологов и специалистов по планированию. Оно автоматизирует процесс сбора и расчёта требований к материалам для производства различных изделий, а также формирует отформатированные документы, такие как докладные записки и заявки.

Приложение сканирует структурированную сетевую директорию, где каждое изделие представлено в виде иерархии папок, а его состав (спецификация) описан в Excel-файлах. Stockwise считывает эти файлы, суммирует одинаковые позиции материалов, позволяет динамически пересчитывать их количество на заданный объём выпуска и экспортирует итоговую ведомость в удобный формат.

Новая версия представляет возможность создавать два типа документов: "Докладные записки" и "Заявки", с настраиваемыми полями для подписей и раздельными черными и белыми списками для материалов.

## Ключевые возможности

-   **Иерархический поиск**: Умный поиск по изделиям с поддержкой частичных совпадений и нескольких ключевых слов.
-   **Автоматическая консолидация**: Сбор данных из множества Excel-файлов, включая вложенные папки спецификаций.
-   **Суммирование материалов**: Автоматическое объединение и суммирование количества для одинаковых позиций номенклатуры.
-   **Динамический пересчёт норм**: Возможность рассчитать требуемое количество материалов для любого объёма производства.
-   **Создание документов**: Формирование "Докладных записок" и "Заявок" с предзаполненными данными.
-   **Экспорт в 3 формата**: Выгрузка итоговой спецификации в **Excel (.xlsx)**, **Word (.docx)** и **PDF** с настроенным форматированием для печати на листах A4.
-   **Автообновление**: Приложение автоматически проверяет наличие новой версии на сервере и предлагает выполнит обновление.
-   **Настраиваемые блэклисты/вайтлисты**: Исключайте или включайте определённые материалы из итогового расчёта для разных типов документов.
-   **Настраиваемые подписи**: Настраивайте поля "от кого" и "кому" в создаваемых документах.
-   **Современный интерфейс**: Понятный и отзывчивый интерфейс, созданный с помощью PyQt5.

## Принцип работы

Для корректной работы приложения необходимо соблюдать определённую структуру хранения данных в папке, указанной в `config.yaml`.

1.  **Корневая папка**: Папка, в которой хранятся все изделия.
2.  **Папка изделия**: Каждое изделие — это иерархия вложенных папок. Имя изделия формируется из пути к конечной папке.
    -   *Пример*: Путь `\SRV\PNDatabases\Stockwise\БВВ\01` будет интерпретирован как изделие `БВВ 01`.
3.  **Файлы спецификаций**: Внутри конечной папки изделия должны находиться Excel-файлы (`.xlsx`, `.xls`), содержащие перечень материалов.
4.  **Структура Excel-файла**: Приложение ожидает, что лист с названием `TDSheet` будет содержать как минимум три колонки: `Номенклатура`, `Количество`, `Ед. изм.`.
5.  **Конфигурация**: Основные настройки находятся в `config.yaml`:
    - `path_to_products_folder`: Путь к корневой папке с данными изделий.
    - `server_program_path`: Путь к папке на сервере для проверки обновлений.
    - `signature_from_human`, `signature_from_position`, `signature_whom_human`, `signature_whom_position`: Списки имен и должностей для полей "от кого" и "кому" в документах.
    - `document_blacklist`, `bid_blacklist`: Черные списки материалов для "Докладных записок" и "Заявок".
    - `document_whitelist`, `bid_whitelist`: Белые списки материалов для "Докладных записок" и "Заявок".

## Установка и запуск

Для запуска проекта из исходного кода выполните следующие шаги:

1.  **Клонируйте репозиторий:**
    ```bash
    git clone <URL репозитория>
    cd Stockwise
    ```

2.  **Создайте и активируйте виртуальное окружение:**
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Установите зависимости:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Настройте конфигурацию:**
    -   Откройте файл `config.yaml`.
    -   Укажите корректный путь к сетевой папке с изделиями в параметре `path_to_products_folder`.
    -   Укажите путь к папке с актуальной версий программы на сервере в `server_program_path`.
    -   При необходимости настройте подписи и черные/белые списки.

5.  **Запустите приложение:**
    ```bash
    python app.py
    ```

## Технологический стек

-   **Язык**: Python 3.10
-   **GUI**: PyQt5
-   **Обработка данных**: Pandas
-   **Экспорт**: Excel: `openpyxl`
-   **Архитектура**: MVC (Model-View-Controller)
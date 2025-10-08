# Stockwise (English Version)

A simple and powerful desktop application for calculating and consolidating bill of materials.

## Changes in the new version

- **The logic of forming a list of materials has been changed.** Now the names of semi-finished products are not included in the final specification.

## Description

**Stockwise** is a desktop application designed for engineers, technologists, and planning specialists. It automates the process of gathering and calculating material requirements for manufacturing various products.

The application scans a structured network directory where each product is represented by a folder hierarchy, and its bill of materials is described in Excel files. Stockwise reads these files, sums up identical material items, allows for dynamic recalculation of quantities for a given production volume, and exports the final report into a convenient format.

## Key Features

-   **Hierarchical Search**: Smart product search with support for partial matches and multiple keywords.
-   **Automatic Consolidation**: Gathers data from multiple Excel files, including nested specification folders (e.g., "рмп").
-   **Material Aggregation**: Automatically combines and sums quantities for identical nomenclature items.
-   **Dynamic Norm Calculation**: Ability to calculate required material quantities for any production volume (defaults to 1000 units).
-   **Export to 3 Formats**: Exports the final specification to **Excel (.xlsx)**, **Word (.docx)**, and **PDF**, with pre-configured formatting for A4 printing.
-   **Auto-Update**: The application automatically checks for a new version on the server and prompts for an update.
-   **Material Blacklist**: Exclude specific materials from the final calculation using a configurable blacklist.
-   **Modern UI**: A clean and responsive user interface built with PyQt5.

## How It Works

For the application to function correctly, the data must be stored in a specific structure within the folder specified in `config.yaml`.

1.  **Root Folder**: The main folder containing all product data.
2.  **Product Folder**: Each product is a folder hierarchy. The product name is derived from the path to the final folder.
    -   *Example*: The path `\SRV\PNDatabases\Stockwise\БВВ\01` will be interpreted as the product `БВВ 01`.
3.  **Specification Files**: Inside the final product folder (or in a subfolder named `рмп`), there should be Excel files (`.xlsx`, `.xls`) containing the list of materials.
4.  **Excel File Structure**: The application expects a sheet named `TDSheet` to contain at least three columns: `Номенклатура` (Nomenclature), `Количество` (Quantity), and `Ед. изм.` (Unit).
5.  **Configuration**: The main settings are in `config.yaml`:
    - `path_to_products_folder`: Path to the root folder with product data.
    - `server_program_path`: Path to the server folder for checking updates.
    - `blacklist`: (Optional) A list of words to exclude materials from the calculation.

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
    -   Optionally, add a `blacklist` with a list of materials to be excluded.

5.  **Run the application:**
    ```bash
    python app.py
    ```

## Technology Stack

-   **Language**: Python 3.10
-   **GUI**: PyQt5
-   **Data Processing**: Pandas
-   **Exporting**:
    -   Excel: `openpyxl`
    -   Word: `python-docx`
    -   PDF: `reportlab`
-   **Architecture**: MVC (Model-View-Controller)

---

# Stockwise

Простое настольное приложение для расчёта и консолидации спецификаций материалов.

## Изменения в новой версии

- **Изменена логика формирования списка материалов.** Теперь названия полуфабрикатов не попадают в итоговую спецификацию.

## Описание

**Stockwise** — это десктопное приложение, разработанное для инженеров, технологов и специалистов по планированию. Оно автоматизирует процесс сбора и расчёта требований к материалам для производства различных изделий.

Приложение сканирует структурированную сетевую директорию, где каждое изделие представлено в виде иерархии папок, а его состав (спецификация) описан в Excel-файлах. Stockwise считывает эти файлы, суммирует одинаковые позиции материалов, позволяет динамически пересчитывать их количество на заданный объём выпуска и экспортирует итоговую ведомость в удобный формат.

## Ключевые возможности

-   **Иерархический поиск**: Умный поиск по изделиям с поддержкой частичных совпадений и нескольких ключевых слов.
-   **Автоматическая консолидация**: Сбор данных из множества Excel-файлов, включая вложенные папки спецификаций (например, "рмп").
-   **Суммирование материалов**: Автоматическое объединение и суммирование количества для одинаковых позиций номенклатуры.
-   **Динамический пересчёт норм**: Возможность рассчитать требуемое количество материалов для любого объёма производства (по умолчанию — на 1000 единиц).
-   **Экспорт в 3 формата**: Выгрузка итоговой спецификации в **Excel (.xlsx)**, **Word (.docx)** и **PDF** с настроенным форматированием для печати на листах A4.
-   **Автообновление**: Приложение автоматически проверяет наличие новой версии на сервере и предлагает выполнит обновление.
-   **Чёрный список материалов**: Исключайте определённые материалы из итогового расчёта с помощью настраиваемого чёрного списка.
-   **Современный интерфейс**: Понятный и отзывчивый интерфейс, созданный с помощью PyQt5.

## Принцип работы

Для корректной работы приложения необходимо соблюдать определённую структуру хранения данных в папке, указанной в `config.yaml`.

1.  **Корневая папка**: Папка, в которой хранятся все изделия.
2.  **Папка изделия**: Каждое изделие — это иерархия вложенных папок. Имя изделия формируется из пути к конечной папке.
    -   *Пример*: Путь `\SRV\PNDatabases\Stockwise\БВВ\01` будет интерпретирован как изделие `БВВ 01`.
3.  **Файлы спецификаций**: Внутри конечной папки изделия (или в подпапке `рмп`) должны находиться Excel-файлы (`.xlsx`, `.xls`), содержащие перечень материалов.
4.  **Структура Excel-файла**: Приложение ожидает, что лист с названием `TDSheet` будет содержать как минимум три колонки: `Номенклатура`, `Количество`, `Ед. изм.`.
5.  **Конфигурация**: Основные настройки находятся в `config.yaml`:
    - `path_to_products_folder`: Путь к корневой папке с данными изделий.
    - `server_program_path`: Путь к папке на сервере для проверки обновлений.
    - `blacklist`: (Опционально) Список слов для исключения материалов из расчёта.

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
    -   При необходимости добавьте `blacklist` со списком материалов для исключения.

5.  **Запустите приложение:**
    ```bash
    python app.py
    ```

## Технологический стек

-   **Язык**: Python 3.10
-   **GUI**: PyQt5
-   **Обработка данных**: Pandas
-   **Экспорт**:
    -   Excel: `openpyxl`
    -   Word: `python-docx`
    -   PDF: `reportlab`
-   **Архитектура**: MVC (Model-View-Controller)
from ui.documentUI import Ui_MainWindow as DocumentUi

from .document_model import DocumentModel
from .document_view import DocumentView
from .document_controller import DocumentController

from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import Qt

def create_document_window(product_name, norms_calculations_value, materials, current_product_path):
    """Создает, настраивает и возвращает экземпляр окна "Документ" с его MVC стеком."""
    window = QMainWindow()
    
    ui = DocumentUi()
    ui.setupUi(window)

    model = DocumentModel(product_name=product_name, 
                          norms_calculations_value=norms_calculations_value,
                          materials=materials,
                          current_product_path=current_product_path)
    view = DocumentView(ui=ui)
    # Создаем контроллер и сохраняем его как атрибут окна, чтобы он не был удален
    window.controller = DocumentController(model=model, view=view)

    window.setAttribute(Qt.WA_DeleteOnClose)
    
    return window
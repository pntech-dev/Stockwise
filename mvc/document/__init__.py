"""Factory for creating the document generation window.

This module provides a function to instantiate and assemble the MVC components
(Model, View, Controller) for the document window.
"""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow

from ui.documentUI import Ui_MainWindow as DocumentUi

from .document_controller import DocumentController
from .document_model import DocumentModel
from .document_view import DocumentView


def create_document_window(
    product_name: str,
    norms_calculations_value: int,
    materials: list[dict],
    current_product_path: str,
) -> QMainWindow:
    """Creates, configures, and returns an instance of the document window.

    This function sets up the entire MVC stack for the document window,
    initializes it with the necessary data, and returns the configured window
    ready to be shown.

    Args:
        product_name: The name of the current product.
        norms_calculations_value: The value for which norms are calculated.
        materials: The list of materials for the product.
        current_product_path: The file system path to the current product.

    Returns:
        A fully configured QMainWindow instance for the document view.
    """
    window = QMainWindow()

    ui = DocumentUi()
    ui.setupUi(window)

    model = DocumentModel(
        product_name=product_name,
        norms_calculations_value=norms_calculations_value,
        materials=materials,
        current_product_path=current_product_path,
    )
    view = DocumentView(ui=ui)
    
    # Create the controller and store it as an attribute of the window
    # to prevent it from being garbage-collected.
    window.controller = DocumentController(model=model, view=view)

    window.setAttribute(Qt.WA_DeleteOnClose)

    return window
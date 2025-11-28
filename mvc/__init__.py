"""Initializes the MVC package.

This file makes the Model, View, and Controller classes from the sub-packages
(main, document) directly accessible under the 'mvc' namespace.
"""

# Main window components
from .main.main_controller import MainController
from .main.main_model import MainModel
from .main.main_view import MainView

# Document window components
from .document.document_controller import DocumentController
from .document.document_model import DocumentModel
from .document.document_view import DocumentView

"""Shared UI styles for reusable widgets."""

# Menu used for search filters popover
FILTER_MENU_STYLE = """
QMenu {
    background-color: #FFFFFF;
    border: 1px solid #CAD1E0;
    border-radius: 6px;
    padding: 6px;
}
QMenu::item {
    padding: 0px;
    margin: 0px;
    background-color: transparent;
}
QMenu::item:selected {
    background-color: transparent;
}
"""

# Popup list used by the search completer
FILTER_POPUP_LIST_STYLE = """
QListView {
    background-color: #FFFFFF;
    border: 1px solid #CAD1E0;
    border-radius: 6px;
    padding: 4px;
    color: #1F2937;
}
QListView::item {
    padding: 6px 8px;
}
QListView::item:selected {
    background-color: #2563EB;
    color: #FFFFFF;
}
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 2px 0 2px 0;
}
QScrollBar::handle:vertical {
    background: #d3d9e8;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #c5ccdd;
}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 0 2px 0 2px;
}
QScrollBar::handle:horizontal {
    background: #d3d9e8;
    min-width: 20px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal:hover {
    background: #c5ccdd;
}
QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal {
    width: 0;
}
"""

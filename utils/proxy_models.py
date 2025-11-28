from PyQt5.QtCore import QModelIndex, QSortFilterProxyModel, Qt



class CustomFilterProxyModel(QSortFilterProxyModel):
    """A custom proxy model for advanced filtering of suggestions.

    This model overrides the default filtering behavior to implement a word-based
    search. The input filter text is split into words, and a row is accepted
    only if all words are present in the item text.
    """

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        """Determines if a row should be included in the model.

        Args:
            source_row: The row number in the source model.
            source_parent: The parent index in the source model.

        Returns:
            True if the row is accepted, False otherwise.
        """
        index = self.sourceModel().index(source_row, 0, source_parent)
        item_text = self.sourceModel().data(index, Qt.DisplayRole)

        if item_text is None:
            return False

        filter_string = self.filterRegExp().pattern()
        search_words = filter_string.lower().split()

        if not search_words:
            return True

        item_text_lower = item_text.lower()

        return all(word in item_text_lower for word in search_words)
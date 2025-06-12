from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt

from model.ui.BD_Base_Frame import BD_Base_Frame


class BD_Table_Frame(BD_Base_Frame):
    def __init__(self, app, qt_table: QTableWidget, include_check_box=False):
        #TODO: appending new row index should be the same
        super().__init__(app)
        self.table = qt_table
        self.include_check_box = include_check_box

        self.table.keyPressEvent = self.keyPressEvent
        if self.include_check_box:
            assert self.table.columnCount() == 1
            #TODO: need to uncheck the button in default in the front end
            self.uncheck_all_button()
    def remove_row(self, row):
        self.table.removeRow(row)

    def remove_all_rows(self):
        for _ in range(self.table.rowCount()):
            self.remove_row(0)

    def update_item_value(self, row, column, value):
        item = self.table.item(row, column)
        item.setText(value)


    def keyPressEvent(self, event):
        row = self.table.currentRow()
        key = event.key()

        if key in (Qt.Key_Return, Qt.Key_Enter):
            if self.is_row_filled(row):
                self.insert_empty_row(row + 1)
                self.table.setCurrentCell(row + 1, 0)

        elif key == Qt.Key_Insert:
            self.insert_empty_row(row + 1)
            self.table.setCurrentCell(row + 1, 0)

        elif key in (Qt.Key_Backspace, Qt.Key_Delete):
            if self.is_row_empty(row) and self.table.rowCount() > 1:
                self.remove_row(row)
                self.table.setCurrentCell(max(0, row - 1), 0)

        else:
            QTableWidget.keyPressEvent(self.table, event)

    def insert_empty_row(self, index):
        self.table.insertRow(index)
        if self.include_check_box:
            checkbox_item = QTableWidgetItem("")
            checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            checkbox_item.setCheckState(Qt.Checked)
            self.table.setItem(index, 0, checkbox_item)
        else:
            for col in range(self.table.columnCount()):
                self.table.setItem(index, col, QTableWidgetItem(""))

    def insert_row(self, index, values):
        self.table.insertRow(index)
        if self.include_check_box:
            checkbox_item = QTableWidgetItem(values)
            checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            checkbox_item.setCheckState(Qt.Checked)
            self.table.setItem(index, 0, checkbox_item)
        else:
            for col in range(self.table.columnCount()):
                self.table.setItem(index, col, QTableWidgetItem(values[col]))

    def is_row_filled(self, row):
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if item is None or not item.text().strip():
                return False
        return True

    def is_row_empty(self, row):
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if item and item.text().strip():
                return False
        return True
    def uncheck_all_button(self):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            item.setCheckState(Qt.Unchecked)

    def get_row_values(self, row):
        result = []
        for column in range(self.table.columnCount()):
            item = self.table.item(row, column)
            result.append(item.text())
        return result

    def get_current_content(self):
        result = []
        if self.include_check_box:
            for row in range(self.table.rowCount()):
                item = self.table.item(row, 0)
                if item.checkState() == Qt.Checked:
                    result.append(item.text())
            return result
        else:
            for row in range(self.table.rowCount()):
                items = []
                for column in range(self.table.columnCount()):
                    item = self.table.item(row, column)
                    items.append(item.text())
                result.append(items)
            return result
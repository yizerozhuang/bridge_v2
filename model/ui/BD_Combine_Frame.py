from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QLineEdit, QTableWidget, QTableWidgetItem, QPushButton

from utility.pdf_utility import open_in_bluebeam, open_folder

import os
from datetime import date
from pathlib import Path

class BD_Combine_Frame(BD_Base_Frame):
    def __init__(self, app, qt_line_edit: QLineEdit, qt_table:QTableWidget,
                 qt_push_button_up: QPushButton,
                 qt_push_button_down: QPushButton,
                 qt_push_button_remove: QPushButton,
                 qt_push_button_remove_all: QPushButton):
        super().__init__(app)
        self.line_edit = qt_line_edit
        self.table = qt_table
        self.push_button_up = qt_push_button_up
        self.push_button_down = qt_push_button_down
        self.push_button_remove = qt_push_button_remove
        self.push_button_remove_all = qt_push_button_remove

        self.table.dragEnterEvent = self.drag_enter_event
        self.table.dragMoveEvent = self.drag_enter_event
        self.table.dropEvent = self.drop_event
        self.table.itemDoubleClicked.connect(self.on_double_click)

        self.push_button_up.clicked.connect(self.move_selected_row_up)
        self.push_button_down.clicked.connect(self.move_selected_row_down)
        self.push_button_remove.clicked.connect(self.remove_selected_row)
        self.push_button_remove_all.clicked.connect(self.remove_all_rows)


    def drag_enter_event(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    def get_timestamp(self):
        return date.today().strftime("%Y%m%d")

    def drop_event(self, event):
        first_file = Path(event.mimeData().urls()[0].toLocalFile())
        output_path = os.path.join(first_file.parent, f"{self.get_timestamp()}-Combined.pdf")
        self.line_edit.setText(output_path)
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path:
                row_position = self.table.rowCount()
                self.table.insertRow(row_position)
                self.table.setItem(row_position, 0, QTableWidgetItem(file_path))

        event.acceptProposedAction()

    def remove_selected_row(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)

    def remove_all_rows(self):
        self.line_edit.setText("")
        self.table.setRowCount(0)

    def move_selected_row_up(self):
        current_row = self.table.currentRow()
        if current_row > 0:
            self._swap_rows(current_row, current_row - 1)
            self.table.selectRow(current_row - 1)

    def move_selected_row_down(self):
        current_row = self.table.currentRow()
        if 0 <= current_row < self.table.rowCount() - 1:
            self._swap_rows(current_row, current_row + 1)
            self.table.selectRow(current_row + 1)

    def _swap_rows(self, row1, row2):
        for col in range(self.table.columnCount()):
            item1 = self.table.item(row1, col)
            item2 = self.table.item(row2, col)
            text1 = item1.text() if item1 else ""
            text2 = item2.text() if item2 else ""
            self.table.setItem(row1, col, QTableWidgetItem(text2))
            self.table.setItem(row2, col, QTableWidgetItem(text1))
    def get_current_item_path(self):
        if self.table.currentItem() is None:
            return None
        return self.table.currentItem().text()

    def on_double_click(self):
        file_path = self.get_current_item_path()
        if file_path.endswith(".pdf"):
            open_in_bluebeam(file_path)
        elif os.path.isdir(file_path):
            open_folder(file_path)

    def get_output_path(self):
        return self.line_edit.text()
    def get_table_items(self):
        paths = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item:
                paths.append(item.text())
        return paths
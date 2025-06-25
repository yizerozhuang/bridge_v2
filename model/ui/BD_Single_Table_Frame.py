from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog
from PyQt5.QtCore import Qt
from utility.pdf_utility import open_in_bluebeam, open_folder
import os

class BD_Single_Table_Frame(BD_Base_Frame):
    def __init__(self, app, qt_text_edit: QLineEdit, qt_push_button: QPushButton, qt_table: QTableWidget):
        super().__init__(app)

        self.line_edit = qt_text_edit
        self.push_button = qt_push_button
        self.table = qt_table

        self.table.setColumnWidth(0, 1000)
        self.table.itemDoubleClicked.connect(self.on_double_click)
        self.line_edit.textChanged.connect(self.on_text_changed)
        self.push_button.clicked.connect(self.click)

    def set_current_folder(self, dir):
        self.line_edit.setText(dir)

    def click(self):
        input_file_dir = QFileDialog.getExistingDirectory(None, "Select Directory", self.app.current_folder_address)
        self.line_edit.setText("")
        self.line_edit.setText(input_file_dir)

    def on_text_changed(self):
        self.table.setRowCount(0)
        if self.line_edit.text()=="":
            return
        file_list = os.listdir(self.line_edit.text())
        for i, file in enumerate(file_list):
            self.table.insertRow(i)
            table_item = QTableWidgetItem(file)
            table_item.setFlags(table_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(i, 0, table_item)

    def get_current_item_path(self):
        if self.table.currentItem() is None:
            return None
        return os.path.join(self.line_edit.text(), self.table.currentItem().text())
    def on_double_click(self):
        file_path = self.get_current_item_path()
        if file_path.endswith(".pdf"):
            open_in_bluebeam(file_path)
        elif os.path.isdir(file_path):
            open_folder(file_path)

    def load(self):
        self.set_current_folder(self.app.current_folder_address)

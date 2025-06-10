from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QLineEdit, QTextEdit, QPushButton, QFileDialog

from utility.global_variables import current_folder_address
class BD_Single_File_Frame(BD_Base_Frame):
    def __init__(self, app, qt_line_edit: QLineEdit, qt_text_edit: QTextEdit, qt_push_button: QPushButton):
        super().__init__(app)

        self.line_edit = qt_line_edit
        self.text_edit = qt_text_edit
        self.push_button = qt_push_button
        self.push_button.clicked.connect(self.click)
        # table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.text_edit.setAcceptDrops(True)
        self.text_edit.dragMoveEvent = self.dragEnterEvent
        self.text_edit.dragEnterEvent = self.dragEnterEvent
        self.text_edit.dropEvent = self.drop_event_single
        # table.setStyleSheet(table_style)

    def click(self):
        input_file_dir = QFileDialog.getOpenFileName(None, "Choose File", current_folder_address)[0]
        self.text_edit.setText(input_file_dir)

    def drop_event_single(self, event):
        file_name = event.mimeData().text().rstrip('\n').split('\n')[0].lstrip("file:///")
        self.text_edit.setText(file_name)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def get_input_file(self):
        return self.text_edit.toPlainText()
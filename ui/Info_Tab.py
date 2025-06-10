from model.ui.BD_Base_Frame import BD_Base_Frame

from utility.pdf_utility import open_folder
from functools import partial

class Info_Tab(BD_Base_Frame):
    def __init__(self, app):
        super().__init__(app)

        self.project_info_tab()
        self.open_function_tab()

    def project_info_tab(self):
        self.info_line_edit_quotation_number = self.ui.info_line_edit_quotation_number
        self.info_line_edit_project_number = self.ui.info_line_edit_project_number
        self.info_line_edit_project_name = self.ui.info_line_edit_project_name

    def open_function_tab(self):
        self.info_push_button_open_folder = self.ui.info_push_button_open_folder
        self.info_push_button_open_backup = self.ui.info_push_button_open_backup
        self.info_push_button_open_asana = self.ui.info_push_button_open_asana

        self.info_push_button_open_folder.clicked.connect(self.open_folder)

    def open_folder(self):
        open_folder(self.app.current_folder_address)

    def load(self):
        self.info_line_edit_quotation_number.setText(self.app.quotation_number)
        self.info_line_edit_project_number.setText(self.app.project_number)
        self.info_line_edit_project_name.setText(self.app.project_name)

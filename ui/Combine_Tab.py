from model.ui.BD_Base_Frame import BD_Base_Frame
from model.ui.BD_Combine_Frame import BD_Combine_Frame

from utility.pdf_tool import PDFTools
from utility.pdf_utility import open_pdf_in_bluebeam, flatten_pdf, move_file_to_ss

from pathlib import Path
import os

class Combine_Tab(BD_Base_Frame):
    def __init__(self, app):
        super().__init__(app)
        self.table = BD_Combine_Frame(self.app, self.ui.combine_line_edit, self.ui.combine_table,
                                      self.ui.combine_push_button_up,
                                      self.ui.combine_push_button_down,
                                      self.ui.combine_push_button_remove,
                                      self.ui.combine_push_button_remove_all
                                      )
        self.combine_push_button = self.ui.combine_push_button_combine
        self.combine_push_button.clicked.connect(self.combine)

    @property
    def output_path(self):
        return self.table.get_output_path()
    @property
    def table_items(self):
        return self.table.get_table_items()

    def combine(self):
        assert self.output_path.endswith(".pdf"), "The output should be in pdf format"
        for item in self.table_items:
            assert item.endswith(".pdf"), f"{item} is not a pdf"
        ss_folder = os.path.join(Path(self.output_path).parent, "SS")
        move_file_to_ss(self.output_path, ss_folder)
        PDFTools.combine_pdf(self.table_items, self.output_path)
        flatten_pdf(self.output_path, self.output_path)
        open_pdf_in_bluebeam(self.output_path)
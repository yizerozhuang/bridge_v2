from model.ui.BD_Base_Frame import BD_Base_Frame

from model.ui.BD_Single_Table_Frame import BD_Single_Table_Frame
from model.ui.BD_Radio_Button_Frame import BD_Radio_Button_Frame
from model.ui.BD_Table_Frame import BD_Table_Frame
from model.ui.BD_Clipboard_Frame import BD_Clipboard_Frame

from utility.pdf_utility import get_pdf_page_numbers

import os

class Drawing_Tab(BD_Base_Frame):
    template_types = ["restaurant", "others"]
    service_types = ["mechanical", "electrical", "hydraulic", "fire_service"]

    template_dir = r"T:\00-Template-Do Not Modify\09-AutoCAD and Bluebeam\01-Mech template"

    def __init__(self, app):
        super().__init__(app)

        self.input_file_table_windows = BD_Single_Table_Frame(
            self.app,
            self.ui.drawing_line_edit_input_folder,
            self.ui.drawing_push_button_open_folder,
            self.ui.drawing_table_input_files
        )
        self.input_file_table_windows.table.itemDoubleClicked.connect(self.on_table_item_double_click)
        self.drawing_line_edit_drawn_by = self.ui.drawing_line_edit_drawn_by
        self.drawing_line_edit_check_by = self.ui.drawing_line_edit_check_by
        self.template_types_radio_button_windows = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui,f"drawing_radio_button_{type}") for type in Drawing_Tab.template_types],
            Drawing_Tab.template_types
        )
        self.service_types_radio_button_windows = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"drawing_radio_button_{type}") for type in Drawing_Tab.service_types],
            Drawing_Tab.service_types,
            self.ui.drawing_radio_button_custom,
            self.ui.drawing_line_edit_custom
        )

        self.drawing_line_edit_current_page = self.ui.drawing_line_edit_current_page
        self.drawing_line_edit_total_page = self.ui.drawing_line_edit_current_page
        self.drawing_line_edit_scale = self.ui.drawing_line_edit_scale
        self.drawing_line_edit_paper_size = self.ui.drawing_line_edit_paper_size

        self.drawing_floor_table = BD_Table_Frame(self.app, self.ui.drawing_floor_table, True)
        self.drawing_revision_table = BD_Table_Frame(self.app, self.ui.drawing_revision_table)

        self.logo_label = BD_Clipboard_Frame(self.app, self.ui.drawing_label_logo)

        self.drawing_floor_table.table.cellChanged.connect(self.on_cell_changed)

        self.drawing_push_button_setup_drawing = self.ui.drawing_push_button_setup_drawing
        self.drawing_push_button_setup_drawing.clicked.connect(self.set_up_drawing)
    def on_table_item_double_click(self):
        file_path = self.input_file_table_windows.get_current_item_path()
        if file_path.endswith(".pdf"):
            page_numbers = get_pdf_page_numbers(file_path)
            self.drawing_line_edit_total_page.setText(str(page_numbers))

    def on_cell_changed(self):
        current_contents = self.drawing_floor_table.get_current_content()
        self.drawing_line_edit_current_page.setText(str(len(current_contents)))
        #TODO: appending new floor should not effect previous floor
        self.drawing_revision_table.remove_all_rows()
        self.drawing_revision_table.insert_row(0, ("M-000", "COVERSHEET, GENERAL NOTES, LEFEND AND DETAILS", "A"))
        for i, content in enumerate(current_contents):
            self.drawing_revision_table.insert_row(i+1, (f"M-10{i+1}", f"{content} LAYOUT", "A"))

    def set_up_drawing(self):
        sketch_dir = self.input_file_table_windows.get_current_item_path()
        pages = int(self.drawing_line_edit_total_page.text())
        paper_size = self.drawing_line_edit_paper_size.text()
        scale = int(self.drawing_line_edit_scale.text())
        template_type = self.template_types_radio_button_windows.get_selection_text()
        service_type = self.service_types_radio_button_windows.get_selection_value()

        assert paper_size in ['A3','A1','A0'], f"the paper size {paper_size} is not a valid paper size"

        template_1, template_2, template_3 = (os.path.join(Drawing_Tab.template_dir, f"{paper_size}-{template_type}", f"template {i}-{scale}.pdf") for i in range(1, 4))

from model.ui.BD_Base_Frame import BD_Base_Frame

from model.ui.BD_Single_Table_Frame import BD_Single_Table_Frame
from model.ui.BD_Radio_Button_Frame import BD_Radio_Button_Frame
from model.ui.BD_Table_Frame import BD_Table_Frame
from model.ui.BD_Clipboard_Frame import BD_Clipboard_Frame

from utility.pdf_utility import get_pdf_page_numbers, move_file_to_ss, duplicate_page, open_in_bluebeam
from utility.pdf_tool import PDFTools

from model.BD_Process import (messagebox, BD_Setup_Drawing_Process, BD_Move_To_Center_Process, BD_Paste_Logo_Process,
                              BD_Fill_Content_Process, BD_Wait_Process)
from conf.conf import CONFIGURATION as conf

import os
import shutil
from datetime import date


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
        self.drawing_line_edit_total_page = self.ui.drawing_line_edit_total_page
        self.drawing_line_edit_scale = self.ui.drawing_line_edit_scale
        self.drawing_line_edit_paper_size = self.ui.drawing_line_edit_paper_size

        self.drawing_floor_table = BD_Table_Frame(self.app, self.ui.drawing_floor_table, True)
        self.drawing_revision_table = BD_Table_Frame(self.app, self.ui.drawing_revision_table)

        self.logo_label = BD_Clipboard_Frame(self.app, self.ui.drawing_label_logo)

        self.drawing_floor_table.table.cellChanged.connect(self.on_cell_changed)

        self.drawing_push_button_setup_drawing = self.ui.drawing_push_button_setup_drawing
        self.drawing_push_button_setup_drawing.clicked.connect(self.move_to_center)
        # self.drawing_push_button_setup_drawing.clicked.connect(self.paste_logo)
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

    def move_to_center(self):
        #TODO: save some variables to the __init__
        sketch_dir = self.input_file_table_windows.get_current_item_path()
        pages = int(self.drawing_revision_table.get_row_count())
        paper_size = self.drawing_line_edit_paper_size.text()
        scale = int(self.drawing_line_edit_scale.text())
        template_type = self.template_types_radio_button_windows.get_selection_text()
        output_dir = os.path.join(self.app.current_folder_address, f"{self.app.project_name}-Mechanical Drawings.pdf")
        assert int(self.drawing_line_edit_total_page.text()) == int(self.drawing_line_edit_current_page.text())
        assert paper_size in ['A3','A1','A0'], f"the paper size {paper_size} is not a valid paper size"
        for i in range(1, PDFTools.page_count(sketch_dir) + 1):
            markups = PDFTools.return_markup_by_page(sketch_dir, i)
            markups = PDFTools.filter_markup_by(markups, {"subject": "Rectangle", "color": "#7C0000"})
            assert len(markups) > 0, f"Page {i} don't have any rectangular"
            assert len(markups) == 1, f"Page {i} have more than one rectangular"
        # assert scale.is_integer(), f"the scale {scale} is wrong"

        template = os.path.join(Drawing_Tab.template_dir, f"{paper_size}-{template_type}", f"template-{scale}.pdf")

        move_file_to_ss(output_dir, os.path.join(self.app.current_folder_address, "SS"))
        shutil.copy(template, output_dir)
        duplicate_page(template, output_dir, pages - 1) # exclude the cover page

        # process = BD_Move_To_Center_Process("Moving sketch to paper center.", self.ui, sketch_dir, paper_size)
        # process.error_occurred.connect(self.handle_thread_error)
        # process.process_finished.connect(self.set_up_drawing)
        # process.start_process()
        self.set_up_drawing()
    
    def set_up_drawing(self):
        sketch_dir = self.input_file_table_windows.get_current_item_path()
        output_dir = os.path.join(self.app.current_folder_address, f"{self.app.project_name}-Mechanical Drawings.pdf")
        process = BD_Setup_Drawing_Process("Changing sketch to drawing.", self.ui, sketch_dir, output_dir)
        if not process.is_available():
            if messagebox('Waiting confirm', 'Someone is doing a task, do you want to wait?', self.ui):
                wait_process = BD_Wait_Process(
                    "Waiting for someone else to finish the task, will start the task once it's available", self.ui)
                wait_process.start_process()
            else:
                return
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.fill_content)
        process.start_process()
    def fill_content(self):
        drawn_by = self.drawing_line_edit_drawn_by.text()
        check_by = self.drawing_line_edit_check_by.text()
        pages = int(self.drawing_revision_table.get_row_count())
        scale = self.drawing_line_edit_scale.text()
        paper_size = self.drawing_line_edit_paper_size.text()
        template_type = self.template_types_radio_button_windows.get_selection_text()
        project_number = self.app.quotation_number if None else self.app.project_number
        project_name = self.app.project_name.upper()
        output_dir = os.path.join(self.app.current_folder_address, f"{self.app.project_name}-Mechanical Drawings.pdf")

        content_replace_dict_list = []

        for i in range(pages):
            drawing_number, drawing_name, drawing_revision = self.drawing_revision_table.get_row_values(i)
            content_replace_dict_list.append(
                {
                    "revision": drawing_revision,
                    "revision_description": "ISSUED FOR APPROVAL",
                    "drawn_by": drawn_by,
                    "check_by":check_by,
                    "date": str(date.today().strftime("%Y%m%d")),
                    "address":project_name,
                    "title":drawing_name,
                    "scale": 'N.T.S@' + paper_size if i == 0 else scale + '@' + paper_size,
                    "project_number": project_number,
                    "drawing_number": drawing_number,
                    "issue": "FOR APPROVAL"
                }
            )

        cover_page = os.path.join(Drawing_Tab.template_dir, f"{paper_size}-{template_type}", f"cover page-{scale}.pdf")

        process = BD_Fill_Content_Process('Setting frame information.', self.ui, cover_page, output_dir, content_replace_dict_list)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.paste_logo)
        process.start_process()

    def paste_logo(self):
        if not self.logo_label.is_image_included():
            self.set_up_success()
            return
        output_dir = os.path.join(self.app.current_folder_address, f"{self.app.project_name}-Mechanical Drawings.pdf")
        paper_size = self.drawing_line_edit_paper_size.text()
        logo_dir = os.path.join(conf["c_temp_dir"], "logo.png")

        self.logo_label.save_image(logo_dir)
        process = BD_Paste_Logo_Process("Setting client logo, open in Bluebeam when done.", self.ui,output_dir, paper_size, logo_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.set_up_success)
        process.start_process()
    def set_up_success(self):
        output_dir = os.path.join(self.app.current_folder_address, f"{self.app.project_name}-Mechanical Drawings.pdf")
        open_in_bluebeam(output_dir)
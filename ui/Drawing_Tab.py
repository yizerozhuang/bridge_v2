from model.ui.BD_Base_Frame import BD_Base_Frame

from model.ui.BD_Single_Table_Frame import BD_Single_Table_Frame
from model.ui.BD_Radio_Button_Frame import BD_Radio_Button_Frame
from model.ui.BD_Table_Frame import BD_Table_Frame
from model.ui.BD_Clipboard_Frame import BD_Clipboard_Frame

from utility.pdf_utility import get_pdf_page_numbers, move_file_to_ss, duplicate_page, open_in_bluebeam
from utility.pdf_tool import PDFTools

from model.BD_Process import (messagebox, BD_Setup_Drawing_Process, BD_Move_To_Center_Process, BD_Paste_Logo_Process,
                              BD_Fill_Content_Process, BD_Wait_Process, BD_Copy_Markup_Process)
from conf.conf import CONFIGURATION as conf

import os
import shutil
from datetime import date


class Drawing_Tab(BD_Base_Frame):
    template_types = ["restaurant", "others"]
    service_types = ["mechanical", "electrical", "hydraulic", "fire_service"]

    template_folder = r"T:\00-Template-Do Not Modify\09-AutoCAD and Bluebeam\Copilot Template"

    service_suffix = {
        "Mechanical": [""],
        "Electrical": ["POWER", "LIGHTING"],
        "Hydraulic": ["DRAINAGE", "WATER"]
    }

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
        self.service_types_radio_button_windows.button_group.buttonClicked.connect(self.on_cell_changed)

        self.drawing_push_button_setup_drawing = self.ui.drawing_push_button_setup_drawing
        self.drawing_push_button_setup_drawing.clicked.connect(self.move_to_center)
        # self.drawing_push_button_setup_drawing.clicked.connect(self.paste_logo)
    @property
    def service(self):
        return self.service_types_radio_button_windows.get_selection_text()
    @property
    def paper_size(self):
        return self.drawing_line_edit_paper_size.text()
    @property
    def scale(self):
        return self.drawing_line_edit_scale.text()
    @property
    def template_type(self):
        return self.template_types_radio_button_windows.get_selection_text()
    @property
    def output_dir(self):
        return os.path.join(self.app.current_folder_address, f"{self.app.project_name}-{self.service} Drawings.pdf")
    @property
    def template_dir(self):
        return os.path.join(Drawing_Tab.template_folder, self.service, f"{self.paper_size}-{self.template_type}-{self.scale}-{self.sketch_dir_orientation}-template.pdf")
    @property
    def cover_page_dir(self):
        return os.path.join(Drawing_Tab.template_folder, self.service, f"{self.paper_size}-{self.template_type}-{self.scale}-{self.sketch_dir_orientation}-cover page.pdf")
    @property
    def current_contents(self):
        return self.drawing_floor_table.get_current_content()
    @property
    def sketch_dir(self):
        return self.input_file_table_windows.get_current_item_path()
    @property
    def sketch_dir_rotation(self):
        if self.sketch_dir is None:
            return None
        #assume all page has the same orientation
        return PDFTools.get_page_rotate(self.sketch_dir, 1)
    @property
    def sketch_dir_orientation(self):
        if self.sketch_dir is None:
            return None
        #assume all page has the same orientation
        return PDFTools.get_page_rotate_orientation(self.sketch_dir, 1)[1]
    @property
    def sketch_pages(self):
        return self.drawing_line_edit_total_page.text() if not self.drawing_line_edit_total_page.text().isdigit() else int(self.drawing_line_edit_total_page.text())
    @property
    def current_pages(self):
        return self.drawing_line_edit_current_page.text() if not self.drawing_line_edit_current_page.text().isdigit() else int(self.drawing_line_edit_current_page.text())
    # @property
    # def output_content_pages(self):
    #     if self.service=="":
    #         return
    #     return self.sketch_pages * len(Drawing_Tab.service_suffix[self.service])
    @property
    def table_row_count(self):
        return int(self.drawing_revision_table.get_row_count())
    @property
    def drawn_by(self):
        return self.drawing_line_edit_drawn_by.text()
    @property
    def check_by(self):
        return self.drawing_line_edit_check_by.text()
    @property
    def project_number(self):
        return self.app.quotation_number if None else self.app.project_number
    @property
    def project_name(self):
        return self.app.project_name.upper()
    def on_table_item_double_click(self):
        file_path = self.input_file_table_windows.get_current_item_path()
        if file_path.endswith(".pdf"):
            page_numbers = get_pdf_page_numbers(file_path)
            self.drawing_line_edit_total_page.setText(str(page_numbers))

    def on_cell_changed(self):
        #TODO: appending new floor should not effect previous floor
        self.drawing_line_edit_current_page.setText(str(len(self.current_contents)))
        self.drawing_revision_table.remove_all_rows()
        self.drawing_revision_table.insert_row_on_the_last(f"{self.service[0]}-000", "COVERSHEET, GENERAL NOTES, LEGEND AND DETAILS", "A")
        if self.service == "Electrical":
            self.drawing_revision_table.insert_row_on_the_last(f"{self.service[0]}-001", "ELECTRICAL SERVICE SINGLE LINE DIAGRAMS", "A")
        for i, suffix in enumerate(Drawing_Tab.service_suffix[self.service]):
            for j, content in enumerate(self.current_contents):
                self.drawing_revision_table.insert_row_on_the_last(f"{self.service[0]}-{i+1}0{j}", f"{content} {suffix} LAYOUT", "A")


    def move_to_center(self):
        assert self.sketch_pages == self.current_pages
        assert self.paper_size in ['A3','A1','A0'], f"the paper size {self.paper_size} is not a valid paper size"
        for i in range(1, PDFTools.page_count(self.sketch_dir) + 1):
            markups = PDFTools.return_markup_by_page(self.sketch_dir, i)
            markups = PDFTools.filter_markup_by(markups, {"subject": "Rectangle", "color": "#7C0000"})
            assert len(markups) > 0, f"Page {i} don't have any rectangular"
            assert len(markups) == 1, f"Page {i} have more than one rectangular"

        move_file_to_ss(self.output_dir, os.path.join(self.app.current_folder_address, "SS"))
        shutil.copy(self.template_dir, self.output_dir)
        duplicate_page(self.template_dir, self.output_dir, self.sketch_pages)

        # process = BD_Move_To_Center_Process("Moving sketch to paper center.", self.ui, sketch_dir, paper_size)
        # process.error_occurred.connect(self.handle_thread_error)
        # process.process_finished.connect(self.set_up_drawing)
        # process.start_process()
        self.set_up_drawing()
    
    def set_up_drawing(self):
        process = BD_Setup_Drawing_Process("Changing sketch to drawing.", self.ui, self.sketch_dir, self.output_dir)
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
        content_replace_dict_list = []

        for i in range(self.table_row_count):
            drawing_number, drawing_name, drawing_revision = self.drawing_revision_table.get_row_values(i)
            content_replace_dict_list.append(
                {
                    "revision": drawing_revision,
                    "revision_description": "ISSUED FOR APPROVAL",
                    "drawn_by": self.drawn_by,
                    "check_by": self.check_by,
                    "date": str(date.today().strftime("%Y%m%d")),
                    "address": self.project_name,
                    "title":drawing_name,
                    "scale": 'N.T.S@' + self.paper_size if i == 0 else f"1:{self.scale}" + '@' + self.paper_size,
                    "project_number": self.project_number,
                    "drawing_number": drawing_number,
                    "issue": "FOR APPROVAL"
                }
            )

        if len(Drawing_Tab.service_suffix[self.service]) == 2:
            PDFTools.duplicate_page(self.output_dir, self.output_dir, 1)

        process = BD_Fill_Content_Process('Setting title block information.', self.ui, self.cover_page_dir, self.output_dir, content_replace_dict_list, self.service, self.sketch_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.paste_logo)
        process.start_process()

    def paste_logo(self):
        if not self.logo_label.is_image_included():
            self.set_up_success()
            return
        logo_dir = os.path.join(conf["c_temp_dir"], "logo.png")

        self.logo_label.save_image(logo_dir)
        process = BD_Paste_Logo_Process("Setting client logo, open in Bluebeam when done.", self.ui, self.output_dir, self.paper_size, self.sketch_dir_orientation, logo_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.set_up_success)
        process.start_process()

    # def copy_markup(self):
    #
    #     process = BD_Copy_Markup_Process("Copying markup, open in Bluebeam when done.", self.ui,self.output_dir, self.sketch_dir)
    #     process.error_occurred.connect(self.handle_thread_error)
    #     process.process_finished.connect(self.set_up_success)
    #     process.start_process()

    def set_up_success(self):
        open_in_bluebeam(self.output_dir)
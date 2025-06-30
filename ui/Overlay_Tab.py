from model.ui.BD_Base_Frame import BD_Base_Frame
from model.ui.BD_Single_File_Frame import BD_Single_File_Frame
from model.ui.BD_Radio_Button_Frame import BD_Radio_Button_Frame
from model.BD_Process import BD_Rescale_Process, BD_Align_Process, BD_Overlay_Process

from utility.pdf_utility import get_paper_sizes,open_folder, convert_mm_to_pixel, get_pdf_page_sizes, open_in_bluebeam
from utility.sql_function import get_value_from_table, format_output, order_dict

from exception.BD_Message_Box import BD_Info_Message

import os
import shutil
from datetime import date


class Overlay_Tab(BD_Base_Frame):
    original_scales = [50, 75, 100, 150, 200]
    output_scales = [50, 100]
    paper_types = ["A4", "A3", "A2", "A1", "A0", "A0_24", "A0_26"]
    paper_sizes = order_dict(format_output(get_value_from_table("paper_size")), paper_types)
    original_sizes = get_paper_sizes(paper_sizes, paper_types)
    output_sizes = get_paper_sizes(paper_sizes, paper_types)
    service_types = ["architecture", "structure", "stormwater", "hydraulic", "electrical", "fire"]
    def __init__(self, app):
        super().__init__(app)
        self.rescale_tab()
        self.align_tab()
        self.tool_box = self.ui.overlay_tool_box
    @property
    def overlay_folder(self):
        return self.setup_line_edit.text()

    @property
    def input_file_dir(self):
        return self.rescale_single_file_window.get_input_file()

    @property
    def architecture_drawing_dir(self):
        return os.path.join(self.overlay_folder, f"Combine {self.service_type}.pdf")
    @property
    def mechanical_drawings_dir(self):
        return self.align_single_file_window.get_input_file()

    @property
    def overlay_color(self):
        return self.align_service_type_radio_button_window.get_selection_background_color()
    @property
    def input_scale(self):
        return self.rescale_original_scale_radio_button_window.get_selection_value()

    @property
    def input_size(self):
        return self.rescale_original_size_radio_button_window.get_selection_text().split("(")[0]

    @property
    def input_size_x(self):
        if self.rescale_original_size_radio_button_window.get_selection_value()[0] == "":
            return 0
        return convert_mm_to_pixel(float(self.rescale_original_size_radio_button_window.get_selection_value()[0]))

    @property
    def input_size_y(self):
        if self.rescale_original_size_radio_button_window.get_selection_value()[1] == "":
            return 0
        return convert_mm_to_pixel(float(self.rescale_original_size_radio_button_window.get_selection_value()[1]))

    @property
    def output_scale(self):
        return self.rescale_output_scale_radio_button_window.get_selection_value()

    @property
    def output_size(self):
        return self.rescale_output_size_radio_button_window.get_selection_text().split("(")[0]

    @property
    def output_size_x(self):
        if self.rescale_output_size_radio_button_window.get_selection_value()[0] == "":
            return 0
        return convert_mm_to_pixel(float(self.rescale_output_size_radio_button_window.get_selection_value()[0]))

    @property
    def output_size_y(self):
        if self.rescale_output_size_radio_button_window.get_selection_value()[1] == "":
            return 0
        return convert_mm_to_pixel(float(self.rescale_output_size_radio_button_window.get_selection_value()[1]))
    @overlay_folder.setter
    #Todo use more property setter
    def overlay_folder(self, app):
        self.setup_line_edit.setText(os.path.join(app.current_folder_address,"Drafting", f"{date.today().strftime('%Y%m%d')}-Overlay"))

    @property
    def service_type(self):
        return self.align_service_type_radio_button_window.get_selection_text()
    @property
    def overlay_checked(self):
        return self.overlay_check_box.isChecked()
    @property
    def update_background_checked(self):
        return self.update_background_check_box.isChecked()

    def rescale_tab(self):
        self.setup_line_edit = self.ui.overlay_setup_line_edit
        self.setup_push_button = self.ui.overlay_setup_push_button
        self.setup_push_button.clicked.connect(self.setup_folder)

        self.rescale_single_file_window = BD_Single_File_Frame(
            self.app,
            self.ui.overlay_rescale_line_edit,
            self.ui.overlay_rescale_text_edit,
            self.ui.overlay_rescale_push_button
        )
        self.rescale_original_scale_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"overlay_rescale_original_scale_radiobutton_{scale}") for scale in self.original_scales],
            self.original_scales,
            self.ui.overlay_rescale_original_scale_radiobutton_custom,
            self.ui.overlay_rescale_original_scale_line_edit
        )
        self.rescale_original_size_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"overlay_rescale_original_size_radiobutton_{size}") for size in self.paper_types],
            self.original_sizes,
            self.ui.overlay_rescale_original_size_radiobutton_custom,
            (
            self.ui.overlay_rescale_original_size_line_edit_width, self.ui.overlay_rescale_original_size_line_edit_height)
        )
        self.rescale_output_scale_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"overlay_rescale_output_scale_radiobutton_{scale}") for scale in self.output_scales],
            self.output_scales,
            self.ui.overlay_rescale_output_scale_radiobutton_custom,
            self.ui.overlay_rescale_output_scale_line_edit
        )
        self.rescale_output_size_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"overlay_rescale_output_size_radiobutton_{size}") for size in self.paper_types],
            self.output_sizes,
            self.ui.overlay_rescale_output_size_radiobutton_custom,
            (self.ui.overlay_rescale_output_size_line_edit_width, self.ui.overlay_rescale_output_size_line_edit_height)
        )
        self.overlay_rescale_button = self.ui.overlay_rescale_button
        self.overlay_rescale_button.clicked.connect(self.rescale)

    def setup_folder(self):
        os.makedirs(self.overlay_folder, exist_ok=True)
        open_folder(self.overlay_folder)

    def rescale(self):
        process = BD_Rescale_Process("Rescaling, open in Bluebeam when done.", self.ui, self.input_file_dir,
                                     self.input_scale, self.input_size, self.input_size_x, self.input_size_y,
                                     self.output_scale, self.output_size_x, self.output_size_y, self.architecture_drawing_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.rescale_success)
        process.start_process()

    def rescale_success(self):
        open_in_bluebeam(self.architecture_drawing_dir)
        self.tool_box.setCurrentIndex(1)

    def align_tab(self):
        self.align_single_file_window = BD_Single_File_Frame(
            self.app,
            self.ui.overlay_align_line_edit,
            self.ui.overlay_align_text_edit,
            self.ui.overlay_align_push_button_open
        )
        self.align_service_type_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"overlay_align_radiobutton_{service}") for service in self.service_types],
            self.service_types
        )
        self.overlay_check_box = self.ui.overlay_align_check_box_overlay
        self.update_background_check_box = self.ui.overlay_align_check_box_update_background
        self.overlay_align_push_button_process = self.ui.overlay_align_push_button_process
        self.overlay_align_push_button_process.clicked.connect(self.align)
    def align(self):
        process = BD_Align_Process("Aligning, open in Bluebeam when done.", self.ui, self.architecture_drawing_dir, self.mechanical_drawings_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.overlay)
        process.start_process()

    def overlay(self):
        if not self.overlay_checked:
            self.update_background()
            return
        process = BD_Overlay_Process("Overlaying, open in Bluebeam when done.", self.ui, self.architecture_drawing_dir, self.mechanical_drawings_dir, self.overlay_color)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.update_background)
        process.start_process()

    def update_background(self):
        if not self.update_background_checked:
            self.process_success()
            return
        self.process_success()
    def process_success(self):
        self.tool_box.setCurrentIndex(2)
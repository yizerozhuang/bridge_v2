from model.ui.BD_Base_Frame import BD_Base_Frame
from model.ui.BD_Single_File_Frame import BD_Single_File_Frame
from model.ui.BD_Radio_Button_Frame import BD_Radio_Button_Frame
from model.ui.BD_Table_Frame import BD_Table_Frame
from model.ui.BD_Single_Table_Frame import BD_Single_Table_Frame
from model.ui.BD_Color_Frame import BD_Color_Frame
from model.BD_Process import (messagebox, BD_Rescale_Process, BD_Align_Process,
                              BD_Color_Process, BD_Grayscale_and_Tags_Process,
                              BD_Copy_Markup_Process, BD_Wait_Process)

from utility.pdf_utility import (convert_mm_to_pixel, get_timestamp, get_pdf_page_numbers, get_paper_sizes,
                                 generate_possible_colors,
                                 get_pdf_page_sizes, open_in_bluebeam)
from utility.sql_function import get_value_from_table, format_output, order_dict

from exception.BD_Message_Box import BD_Info_Message

import os
import shutil


class Sketch_Tab(BD_Base_Frame):
    original_scales = [50, 75, 100, 150, 200]
    output_scales = [50, 100]
    paper_types = ["A4", "A3", "A2", "A1", "A0", "A0_24", "A0_26"]
    paper_sizes = order_dict(format_output(get_value_from_table("paper_size")), paper_types)
    original_sizes = get_paper_sizes(paper_sizes, paper_types)
    output_sizes = get_paper_sizes(paper_sizes, paper_types)

    def __init__(self, app):
        super().__init__(app)
        self.rescale_tab()
        self.align_tab()
        self.markup_tab()
        self.checklist_tab()

    @property
    def input_file_dir(self):
        return self.rescale_single_file_window.get_input_file()
    @property
    def current_sketch_dir(self):
        return os.path.join(self.app.current_folder_address, "set_up_sketch.pdf")
    @property
    def page_number(self):
        if os.path.exists(self.current_sketch_dir):
            return get_pdf_page_numbers(self.current_sketch_dir)
        return 0
    @property
    def input_scale(self):
        return self.rescale_original_scale_radio_button_window.get_selection_value()
    @property
    def input_size_x(self):
        if self.rescale_original_size_radio_button_window.get_selection_value()[0] == "":
            return 0
        return convert_mm_to_pixel(self.rescale_original_size_radio_button_window.get_selection_value()[0])
    @property
    def input_size_y(self):
        if self.rescale_original_size_radio_button_window.get_selection_value()[1] == "":
            return 0
        return convert_mm_to_pixel(self.rescale_original_size_radio_button_window.get_selection_value()[1])
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
        return convert_mm_to_pixel(self.rescale_output_size_radio_button_window.get_selection_value()[0])
    @property
    def output_size_y(self):
        if self.rescale_output_size_radio_button_window.get_selection_value()[1] == "":
            return 0
        return convert_mm_to_pixel(self.rescale_output_size_radio_button_window.get_selection_value()[1])
    @property
    def selected_colors(self):
        return generate_possible_colors(self.color_window.get_selected_colors())
    @property
    def luminocity_checked(self):
        return self.color_window.luminocity_is_checked()
    @property
    def luminocity_value(self):
        return self.color_window.get_luminoicity()
    @property
    def selected_floors(self):
        return self.floor_table.get_current_content()
    @property
    def grayscale_checked(self):
        return self.color_window.grayscale_is_checked()
    @property
    def add_tags_checked(self):
        return self.color_window.add_tags_is_checked()
    @property
    def current_table_item(self):
        return self.table_window.get_current_item_path()
    def rescale_tab(self):
        self.rescale_single_file_window = BD_Single_File_Frame(
            self.app,
            self.ui.sketch_rescale_line_edit,
            self.ui.sketch_rescale_text_edit,
            self.ui.sketch_rescale_push_button
        )
        self.rescale_original_scale_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"sketch_rescale_original_scale_radiobutton_{scale}") for scale in self.original_scales],
            self.original_scales,
            self.ui.sketch_rescale_original_scale_radiobutton_custom,
            self.ui.sketch_rescale_original_scale_line_edit
        )
        self.rescale_original_size_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"sketch_rescale_original_size_radiobutton_{size}") for size in self.paper_types],
            self.original_sizes,
            self.ui.sketch_rescale_original_size_radiobutton_custom,
            (self.ui.sketch_rescale_original_size_line_edit_width, self.ui.sketch_rescale_original_size_line_edit_height)
        )
        self.rescale_output_scale_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"sketch_rescale_output_scale_radiobutton_{scale}") for scale in self.output_scales],
            self.output_scales,
            self.ui.sketch_rescale_output_scale_radiobutton_custom,
            self.ui.sketch_rescale_output_scale_line_edit
        )
        self.rescale_output_size_radio_button_window = BD_Radio_Button_Frame(
            self.app,
            [getattr(self.ui, f"sketch_rescale_output_size_radiobutton_{size}") for size in self.paper_types],
            self.output_sizes,
            self.ui.sketch_rescale_output_size_radiobutton_custom,
            (self.ui.sketch_rescale_output_size_line_edit_width, self.ui.sketch_rescale_output_size_line_edit_height)
        )
        self.sketch_rescale_button = self.ui.sketch_rescale_button
        self.sketch_rescale_button.clicked.connect(self.rescale)

    def rescale(self):
        if self.input_scale == self.output_scale and self.input_size_x == self.output_size_x and self.input_size_y == self.output_size_y:
            shutil.copy(self.input_file_dir, self.current_sketch_dir)
            self.rescale_success()
            return
        original_input_size_x, original_input_size_y = get_pdf_page_sizes(self.input_file_dir)
        if round(original_input_size_x, 2) != self.input_size_x or round(original_input_size_y, 2) != self.input_size_y:
            input_size = self.rescale_original_size_radio_button_window.get_selection_text().split("(")[0]
            raise ValueError(f"The input size is not {input_size}")

        process = BD_Rescale_Process("Rescaling, open in Bluebeam when done.", self.ui, self.input_file_dir,
                                     self.input_scale, self.input_size_x, self.input_size_y,
                                     self.output_scale, self.output_size_x, self.output_size_y, self.current_sketch_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.rescale_success)
        process.start_process()

    def rescale_success(self):
        open_in_bluebeam(self.current_sketch_dir)
        self.sketch_align_line_edit_total_page.setText(str(self.page_number))
        self.sketch_align_line_edit_scale.setText(str(self.output_scale))
        self.sketch_align_line_edit_paper_size.setText(self.output_size)
        self.color_window.set_total_pages_number(self.page_number)
        self.table_window.set_current_folder(self.app.current_folder_address)
        self.ui.toolBox.setCurrentIndex(1)

    def align_tab(self):
        self.sketch_align_line_edit_current_page = self.ui.sketch_align_line_edit_current_page
        self.sketch_align_line_edit_total_page = self.ui.sketch_align_line_edit_total_page
        self.sketch_align_line_edit_scale = self.ui.sketch_align_line_edit_scale
        self.sketch_align_line_edit_paper_size = self.ui.sketch_align_line_edit_paper_size

        self.floor_table = BD_Table_Frame(self.app, self.ui.sketch_floor_table, True)
        self.floor_table.table.cellChanged.connect(self.on_cell_changed)

        self.color_window = BD_Color_Frame(
            self.app,
            self.ui.sketch_align_combo_box_page,
            self.ui.sketch_align_frame_source_color_from,
            self.ui.sketch_align_frame_source_color_to,
            self.ui.sketch_align_label_before,
            self.ui.sketch_align_label_after,
            self.ui.sketch_align_check_box_luminocity,
            self.ui.sketch_align_line_edit_luminocity,
            self.ui.sketch_align_check_box_grayscale,
            self.ui.sketch_align_check_box_add_tags
        )
        self.sketch_align_push_button_process = self.ui.sketch_align_push_button_process
        self.sketch_align_push_button_process.clicked.connect(self.align)


    def align(self):
        process = BD_Align_Process("Aligning Sketch, open in Bluebeam when done.", self.ui, self.current_sketch_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.color_modify)
        process.start_process()

    def color_modify(self):
        process = BD_Color_Process("Modifying color and changing luminosity, open in Bluebeam when done.", self.ui, self.current_sketch_dir,
                                   self.page_number, self.selected_colors, self.luminocity_checked, self.luminocity_value)
        if not process.is_available():
            if messagebox('Waiting confirm', 'Someone is doing a task, do you want to wait?', self.ui):
                wait_process = BD_Wait_Process(
                    "Waiting for someone else to finish the task, will start the task once it's avaliable", self.ui)
                wait_process.start_process()
            else:
                return
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.add_tags)
        process.start_process()

    def add_tags(self):
        process = BD_Grayscale_and_Tags_Process("Grayscale and adding tags on Sketch, open in Bluebeam when done.", self.ui,
                                                self.current_sketch_dir, self.selected_floors, self.grayscale_checked,
                                                self.add_tags_checked, self.output_scale, self.output_size)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.process_success)
        process.start_process()

    def process_success(self):
        open_in_bluebeam(self.current_sketch_dir)
        self.ui.toolBox.setCurrentIndex(2)

    def markup_tab(self):
        self.table_window = BD_Single_Table_Frame(
            self.app,
            self.ui.sketch_markup_line_edit,
            self.ui.sketch_markup_push_button,
            self.ui.sketch_markup_table
        )
        self.sketch_markup_push_button_copy_markup = self.ui.sketch_markup_push_button_copy_markup
        self.sketch_markup_push_button_copy_markup.clicked.connect(self.copy_markup)


    def copy_markup(self):
        assert get_pdf_page_numbers(self.current_sketch_dir) == get_pdf_page_numbers(self.current_table_item), "The input page number is not the same as the output page number"
        # if not process.is_available():
        #     if messagebox('Waiting confirm', 'Someone is doing a task, do you want to wait?', self.ui):
        #         wait_process = BD_Wait_Process(
        #             "Waiting for someone else to finish the task, will start the task once it's avaliable", self.ui)
        #         wait_process.start_process()
        #     else:
        #         return
        process = BD_Copy_Markup_Process("Copying markup, open in Bluebeam when done.", self.ui,
                                         self.current_sketch_dir, self.current_table_item)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.copy_success)
        process.start_process()

    def copy_success(self):
        open_in_bluebeam(self.current_sketch_dir)
        self.ui.toolBox.setCurrentIndex(3)

    def checklist_tab(self):
        self.sketch_checklist_push_button = self.ui.sketch_checklist_push_button
        self.sketch_checklist_push_button.clicked.connect(self.complete)

    def complete(self):
        output_dir = os.path.join(self.app.current_folder_address, f"{self.app.project_name}-Mechanical Sketch.pdf")
        if os.path.exists(output_dir):
            shutil.move(output_dir, os.path.join(self.app.current_folder_address, "SS",
                                                 f"{get_timestamp()}-{self.app.project_name}-Mechanical Sketch.pdf"))
        os.rename(self.current_sketch_dir, output_dir)
        open_in_bluebeam(output_dir)
        BD_Info_Message("The Sketch procedure is completed")

    def on_cell_changed(self):
        self.sketch_align_line_edit_current_page.setText(str(len(self.floor_table.get_current_content())))
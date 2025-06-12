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

        self.current_sketch_dir = ""

        self.rescale_tab()
        self.align_tab()
        self.markup_tab()
        self.checklist_tab()


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
        input_file_dir = self.rescale_single_file_window.get_input_file()
        input_scale = self.rescale_original_scale_radio_button_window.get_selection_value()
        input_size_x, input_size_y = self.rescale_original_size_radio_button_window.get_selection_value()
        input_size_x = convert_mm_to_pixel(input_size_x)
        input_size_y = convert_mm_to_pixel(input_size_y)
        output_scale = self.rescale_output_scale_radio_button_window.get_selection_value()
        output_size_x, output_size_y = self.rescale_output_size_radio_button_window.get_selection_value()
        output_size_x = convert_mm_to_pixel(output_size_x)
        output_size_y = convert_mm_to_pixel(output_size_y)

        self.current_sketch_dir = os.path.join(self.app.current_folder_address, f"{get_timestamp()}-set_up_sketch.pdf")

        if input_scale == output_scale and input_size_x == output_size_x and input_size_y == output_size_y:
            shutil.copy(input_file_dir, self.current_sketch_dir)
            self.rescale_success()
            return
        original_input_size_x, original_input_size_y = get_pdf_page_sizes(input_file_dir)
        if round(original_input_size_x, 2) != input_size_x or round(original_input_size_y, 2) != input_size_y:
            input_size = self.rescale_original_size_radio_button_window.get_selection_text().split("(")[0]
            raise ValueError(f"The input size is not {input_size}")

        process = BD_Rescale_Process("Rescaling, open in Bluebeam when done.", self.ui, input_file_dir,
                                     input_scale, input_size_x, input_size_y,
                                     output_scale, output_size_x, output_size_y, self.current_sketch_dir)
        process.error_occurred.connect(self.handle_thread_error)
        process.process_finished.connect(self.rescale_success)
        process.start_process()

    def rescale_success(self):
        page_number = get_pdf_page_numbers(self.current_sketch_dir)
        open_in_bluebeam(self.current_sketch_dir)
        self.sketch_align_line_edit_total_page.setText(str(page_number))
        self.sketch_align_line_edit_scale.setText(str(self.rescale_output_scale_radio_button_window.get_selection_value()))
        output_size = self.rescale_output_size_radio_button_window.get_selection_text().split("(")[0]
        self.sketch_align_line_edit_paper_size.setText(output_size)
        self.color_window.set_total_pages_number(page_number)
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
        selected_colors = generate_possible_colors(self.color_window.get_selected_colors())
        luminocity_checked = self.color_window.luminocity_is_checked()
        luminocity_value = self.color_window.get_luminoicity()
        page_number = get_pdf_page_numbers(self.current_sketch_dir)

        process = BD_Color_Process("Modifying color and changing luminosity, open in Bluebeam when done.", self.ui, self.current_sketch_dir,
                                   page_number, selected_colors, luminocity_checked, luminocity_value)
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
        selected_floors = self.floor_table.get_current_content()
        grayscale_checked = self.color_window.grayscale_is_checked()
        add_tags_checked = self.color_window.add_tags_is_checked()
        output_scale = self.rescale_output_scale_radio_button_window.get_selection_value()
        output_paper_type = self.rescale_output_size_radio_button_window.get_selection_text().split("(")[0]
        process = BD_Grayscale_and_Tags_Process("Grayscale and adding tags on Sketch, open in Bluebeam when done.", self.ui,
                                                self.current_sketch_dir, selected_floors, grayscale_checked,
                                                add_tags_checked, output_scale, output_paper_type)
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
        current_item = self.table_window.get_current_item_path()
        process = BD_Copy_Markup_Process("Copying markup, open in Bluebeam when done.", self.ui, self.current_sketch_dir, current_item)
        assert get_pdf_page_numbers(self.current_sketch_dir) == get_pdf_page_numbers(
            current_item), "The input page number is not the same as the output page number"
        if not process.is_available():
            if messagebox('Waiting confirm', 'Someone is doing a task, do you want to wait?', self.ui):
                wait_process = BD_Wait_Process(
                    "Waiting for someone else to finish the task, will start the task once it's avaliable", self.ui)
                wait_process.start_process()
            else:
                return
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
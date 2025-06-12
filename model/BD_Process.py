import sys

from PyQt5.QtCore import QThread, pyqtSignal
from utility.pdf_utility import (resize_pdf_multiprocessing,align_multiprocessing, grays_scale_pdf_multiprocessing,
                                 add_tags_multiprocessing, flatten_pdf, get_pdf_page_numbers, pdf_move_center, get_number_of_page, combine_pdf,
                                 insert_logo_into_pdf)
from utility.pdf_tool import PDFTools
from PyQt5.QtWidgets import QDialog, QProgressBar, QLabel, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
import time
from datetime import datetime
import multiprocessing
import os
from conf.conf import CONFIGURATION as conf
import traceback


def messagebox(name, mess, main_window):
    message_box = QMessageBox()
    message_box.setWindowTitle(name)
    message_box.setText(mess)
    message_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    message_box.setDefaultButton(QMessageBox.No)
    screen = main_window.screen()
    screen_geometry = screen.availableGeometry()
    message_geometry = message_box.frameGeometry()
    message_geometry.moveCenter(screen_geometry.center())
    message_box.move(message_geometry.topLeft())
    set = message_box.exec_()
    return set == QMessageBox.Yes

class BD_Worker(QThread):
    progress = pyqtSignal(int)
    def __init__(self, sum_time):
        super().__init__()
        self.sum_time=sum_time
    def run(self):
        self.running = True
        for i in range(100):
            if not self.running:
                break
            time.sleep(self.sum_time/100)
            self.progress.emit(i)
    def stop(self):
        self.running = False

class ProgressDialog(QDialog):
    def __init__(self, info, ui):
        super().__init__()
        self.setWindowTitle("Progress")
        self.layout = QVBoxLayout()
        self.label = QLabel(info)
        self.layout.addWidget(self.label)
        self.setWindowFlags(self.label.windowFlags() & ~Qt.WindowCloseButtonHint)
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("QProgressBar::chunk { background-color: #00FF00; }")
        self.layout.addWidget(self.progress_bar)
        self.setLayout(self.layout)
        screen = ui.screen()
        screen_geometry = screen.availableGeometry()
        message_geometry = self.frameGeometry()
        message_geometry.moveCenter(screen_geometry.center())
        self.move(message_geometry.topLeft())

    def update_progress(self, value):
        self.progress_bar.setValue(value)


class BD_Process(QThread):
    #find the type of traceback
    error_occurred = pyqtSignal(tuple)
    process_finished = pyqtSignal()
    def __init__(self, info, ui):
        self.ui=ui
        super().__init__()
        self.progress_dialog = ProgressDialog(info, ui)
        self.processing_time = 300

    def run(self):
        try:
            self.worker = BD_Worker(self.processing_time)
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            self.function()
            self.worker.stop()
            self.process_finished.emit()
        except Exception:
            self.worker.stop()
            self.error_occurred.emit(sys.exc_info())

    def function(self):
        pass

    # def stop(self):
    #     self.quit()

    def start_process(self):
        self.start()
        self.progress_dialog.exec_()
        self.quit()


    @staticmethod
    def is_available():
        for dir in os.listdir(conf["trans_dir"]):
            if dir.isdigit():
                return False
        return True



class BD_Rescale_Process(BD_Process):
    def __init__(self, info, ui, input_file_dir, input_scale, input_size_x, input_size_y,
                 output_scale, output_size_x, output_size_y, output_file_dir):
        super().__init__(info, ui)
        self.input_file_dir = input_file_dir
        self.input_scale = input_scale
        self.input_size_x = input_size_x
        self.input_size_y = input_size_y
        self.output_scale = output_scale
        self.output_size_x = output_size_x
        self.output_size_y = output_size_y
        self.output_file_dir = output_file_dir

    def function(self):
        # temp_file_path = os.path.join(conf["c_temp_dir"], f"{get_timestamp()}-temp.pdf")
        pool = multiprocessing.Pool(20)
        resize_pdf_multiprocessing(pool, self.input_file_dir, self.input_scale, self.input_size_x, self.input_size_y,
                               self.output_scale, self.output_size_x, self.output_size_y, self.output_file_dir)
        pool.close()
        time.sleep(2)
        flatten_pdf(self.output_file_dir, self.output_file_dir)


class BD_Align_Process(BD_Process):
    def __init__(self, info, ui, current_sketch_dir):
        super().__init__(info, ui)
        self.current_sketch_dir = current_sketch_dir

    def function(self):
        self.worker.progress.connect(self.progress_dialog.update_progress)
        self.worker.finished.connect(self.progress_dialog.accept)
        self.worker.start()
        pool = multiprocessing.Pool(20)
        align_multiprocessing(pool, self.current_sketch_dir)
        pool.close()
        # except AttributeError:
        #     messagebox("error", "you should put down the polyline", self.ui)
        # finally:
        #     pool.close()



class BD_Color_Process(BD_Process):
    def __init__(self, info, ui, current_sketch_dir, page_number, selected_colors, luminocity_checked, luminocity_value):
        super().__init__(info, ui)
        self.current_sketch_dir = current_sketch_dir
        self.page_number = page_number
        self.selected_colors = selected_colors
        self.luminocity_checked = luminocity_checked
        self.luminocity_value = luminocity_value

    def function(self):
        self.worker.progress.connect(self.progress_dialog.update_progress)
        self.worker.finished.connect(self.progress_dialog.accept)
        self.worker.start()

        file_path = os.path.join(conf["trans_dir"], datetime.now().strftime("%Y%m%d%H%M%S"), 'file_names_lumicolor.txt')
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        with open(file_path, 'w') as file:
            file_names = {
                'lumionoff': self.luminocity_checked,
                'changecoloronoff': len(self.selected_colors)==0,
                'file': self.current_sketch_dir,
                'pagenum': self.page_number,
                'color_change': self.selected_colors,
                'lumi_set': self.luminocity_value
            }
            file.write(str(file_names))
        while True:
            time.sleep(2)
            if self.is_available():
                break

class BD_Grayscale_and_Tags_Process(BD_Process):

    markup_dict = {'A4': "Textbox-Size24", 'A3': "Textbox-Size_A3", 'A2': "Textbox-Size_A2",
                   'A1': "Textbox-Size_A1", 'A0': "Textbox-Size_A0", 'A0-24': "Textbox-Size_A0_24",
                   'A0-26': "Textbox-Size_A0_26", 'custom': "Textbox-Size_A0_26"}
    markup_width_dict = {'A4': 14, 'A3': 24, 'A2': 24, 'A1': 40, 'A0': 50, 'A0-24': 72, 'A0-26': 100, 'custom': 100}

    def __init__(self, info, ui, current_sketch_dir, selected_floors, grayscale_checked, add_tags_checked,
                 output_scale, output_paper_type):
        super().__init__(info, ui)
        self.current_sketch_dir = current_sketch_dir
        self.selected_floors = selected_floors
        self.grayscale_checked = grayscale_checked
        self.add_tags_checked = add_tags_checked
        self.output_scale = output_scale
        self.output_paper_type = output_paper_type
    def function(self):
        self.worker.progress.connect(self.progress_dialog.update_progress)
        self.worker.finished.connect(self.progress_dialog.accept)
        self.worker.start()
        if self.grayscale_checked:
            pool = multiprocessing.Pool(20)
            grays_scale_pdf_multiprocessing(pool, self.current_sketch_dir)
            pool.close()

        if self.add_tags_checked:
            tag = f"1:{self.output_scale}@{self.output_paper_type}"
            tag_x = 0
            tag_y = 0.1
            markup_type = self.markup_dict[self.output_paper_type]
            markup_width_coe = self.markup_width_dict[self.output_paper_type]
            pool = multiprocessing.Pool(20)
            add_tags_multiprocessing(pool, self.current_sketch_dir, markup_type, tag_x, tag_y,
                                     tag, self.selected_floors, markup_width_coe)
            pool.close()

class BD_Copy_Markup_Process(BD_Process):
    def __init__(self, info, ui, current_sketch_dir, existing_sketch_dir):
        super().__init__(info, ui)
        self.current_sketch_dir = current_sketch_dir
        self.existing_sketch_dir = existing_sketch_dir
    def function(self):
        self.worker.progress.connect(self.progress_dialog.update_progress)
        self.worker.finished.connect(self.progress_dialog.accept)
        self.worker.start()
        page_number = get_pdf_page_numbers(self.current_sketch_dir)
        file_path = os.path.join(conf["trans_dir"], datetime.now().strftime("%Y%m%d%H%M%S"), 'file_names_copymarkup.txt')
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # offsets = []
        # for i in range(1, page_number + 1):
        #     existing_markups = PDFTools.return_markup_by_page(self.existing_sketch_dir, i)
        #     new_markups = PDFTools.return_markup_by_page(self.current_sketch_dir, i)
        #     existing_markups = PDFTools.filter_markup_by(existing_markups, {"type": "PolyLine", "color": "#7C0000"})
        #     new_markups = PDFTools.filter_markup_by(new_markups, {"type": "PolyLine", "color": "#7C0000"})
        #     assert len(existing_markups) > 0, f"Page {i} of existing sketch dont have base point"
        #     assert len(new_markups) > 0, f"Page {i} of new sketch dont have base point"
        #     assert len(existing_markups) == 1, f"Page {i} of existing sketch have more than one base point"
        #     assert len(new_markups) == 1, f"Page {i} of existing sketch have more than one base point"
        #
        #     markup_1 = list(existing_markups.values())[0]
        #     markup_2 = list(new_markups.values())[0]
        #     offset = (float(markup_2['x']) - float(markup_1['x']), float(markup_2['y']) - float(markup_1['y']))



        with open(file_path, 'w') as file:
            file_names = {
                "File1": self.existing_sketch_dir,
                "File2": self.current_sketch_dir,
                'copyornot':[[True]*page_number, [True]*page_number]
            }
            file.write(str(file_names))
        # Todo: design the waiting function
        while True:
            time.sleep(2)
            if self.is_available():
                break

class BD_Move_To_Center_Process(BD_Process):
    def __init__(self, info, ui, sketch_dir, paper_size):
        super().__init__(info, ui)
        self.sketch_dir = sketch_dir
        self.paper_size = paper_size
    def function(self):
        pass
        # pdf_move_center(self.sketch_dir, self.paper_size)


class BD_Setup_Drawing_Process(BD_Process):
    def __init__(self, info, ui, sketch_dir, output_dir):
        super().__init__(info, ui)
        self.sketch_dir = sketch_dir
        self.output_dir = output_dir

    def function(self):
        file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
        file_dir_trans = os.path.join(conf["trans_dir"], file_time)
        txt_name = os.path.join(file_dir_trans, 'file_names_setupdrawing.txt')
        os.makedirs(os.path.dirname(txt_name), exist_ok=True)
        with open(txt_name, 'w') as file:
            file_names = self.sketch_dir + ';' + self.output_dir + ';'
            file.write(file_names)
        while True:
            time.sleep(2)
            if self.is_available():
                break

class BD_Fill_Content_Process(BD_Process):
    # TODO: need to fix the name error
    def __init__(self, info, ui, temp1, output_file, temp3, content_replace_dict_list, content_replace_dict_firstpage):
        super().__init__(info, ui)
        self.temp1 = temp1
        self.output_file = output_file
        self.temp3 = temp3
        self.content_replace_dict_list = content_replace_dict_list
        self.content_replace_dict_firstpage = content_replace_dict_firstpage

    def function(self):
        #TODO: remove get number of pages
        pages = get_number_of_page(self.output_file)
        for i in range(pages):
            content_replace_dict = self.content_replace_dict_list[i]
            PDFTools.paste_markup_to_file(self.temp1, self.output_file, page_number=1, new_page_number=i + 1,
                                             offset=(0, 0), content_replace_dict=content_replace_dict)
        time.sleep(2)
        combine_pdf([self.temp3, self.output_file], self.output_file)
        PDFTools.paste_markup_to_file(self.temp1, self.output_file, page_number=1, new_page_number=1, offset=(0, 0),
                                         content_replace_dict=self.content_replace_dict_firstpage)



class BD_Paste_Logo_Process(BD_Process):
    def __init__(self, info, ui, output_dir, paper_size, logo_dir):
        super().__init__(info, ui),
        self.output_dir = output_dir
        self.paper_size = paper_size
        self.logo_dir = logo_dir
    def function(self):
        pages = get_number_of_page(self.output_dir)
        for i in range(pages):
            insert_logo_into_pdf(self.output_dir, self.logo_dir, i, self.paper_size)
            time.sleep(2)


class BD_Wait_Process(BD_Process):
    def __init__(self, info, ui):
        super().__init__(info, ui)
        self.processing_time = 900
    def function(self):
        self.worker.progress.connect(self.progress_dialog.update_progress)
        self.worker.finished.connect(self.progress_dialog.accept)
        self.worker.start()
        while True:
            time.sleep(2)
            if self.is_available():
                break
from ui.Sketch_Tab import Sketch_Tab
from ui.Drawing_Tab import Drawing_Tab
from utility.sql_function import get_value_from_table_with_filter, get_value_from_table_with_filter, format_output


from PyQt5 import uic
from PyQt5.QtWidgets import (QApplication, QWidget, QTableWidget, QTableWidgetItem, QHBoxLayout, QHeaderView, QButtonGroup, QCheckBox, QMessageBox, QVBoxLayout, QLabel,
                             QProgressBar, QDialog, QSizePolicy, QFileDialog, QTabWidget, QMenu, QAction, QPushButton, QAbstractItemView, QShortcut, QFrame,
                             QToolButton, QLineEdit, QGridLayout, QSpacerItem, QTextEdit, QMainWindow)
import os
import json
import ast
from conf.config_copilot import CONFIGURATION as conf
from utility.pdf_utility import (init_environment, open_folder, open_link_with_edge, file_exists, combine_pdf, convert_mm_to_pixel, is_float,
                                 flatten_pdf, resize_pdf_multiprocessing, grays_scale_pdf_multiprocessing, align_multiprocessing,
                                 add_tags_multiprocessing, is_pdf_open, create_or_clear_directory, create_directory,
                                 get_number_of_page, find_keyword, extract_first_page, remove_first_page, get_paper_size, svg_set_opacity_multicolor,
                                 markup_wait, lumicolor_pc_wait, setupdrawing_wait, aligncombine_wait, overlay_pc_wait, plot_wait, remove_wait,
                                 pdf_move_center, insert_logo_into_pdf, remove_duplicates_keep_order, version_setter, get_frame_content, getlogo, getpage,
                                 open_in_bluebeam, remove_pdf_page, parse_page_ranges, get_file_bytes, compress_directory_to_zip,
                                 adjust_hex_color, f_waitinline, page_delete, get_temp_name, list_all_files,
                                 char2num)
init_environment(conf)
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtCore import Qt,QThread, pyqtSignal, QSize,QEvent,QDate, QLoggingCategory
from PyQt5.QtGui import QColor,QKeySequence,QPainter,QPixmap,QFont
import shutil
import sys
from pathlib import Path
import threading
from datetime import date,datetime
import traceback
import multiprocessing
from PyQt5.QtNetwork import QNetworkAccessManager
import keyboard
import pyperclip
import time
import mysql.connector
import subprocess
from utility.google_drive import GoogleDrive
from functools import partial
from utility.pdf_tool import PDFTools as PDFTools_v2
PDFTools_v2.SetEnvironment(conf["bluebeam_dir"], conf["bluebeam_engine_dir"], r"C:\Progra~1\Inkscape\bin\inkscape.exe", conf["c_temp_dir"])
from Bridge import Stats_bridge
import re
from utility.asana_function import get_asana_tasks_by_user_by_date,update_asana_estimated_time,update_asana_actual_time
import math
QLoggingCategory.setFilterRules("*.debug=false\n*.warning=false\n*.critical=false")


'''========================================General set========================================='''
'''Message window set'''

def message(text,main_window):
    message_box = QMessageBox()
    message_box.setTextFormat(Qt.RichText)
    message_box.setWindowTitle('Information')
    set_text = '<font color="black" size="4">' + text.replace('\n', '<br>') + '</font>'
    message_box.setText(set_text)
    message_box.setWindowFlags(message_box.windowFlags() | Qt.WindowStaysOnTopHint)

    screen = QApplication.primaryScreen()
    screen_geometry = screen.availableGeometry()
    main_window_geometry = main_window.geometry()
    message_box.move(screen_geometry.left() + main_window_geometry.left()+800,
                     screen_geometry.top() + main_window_geometry.top()+400)

    message_box.exec_()
'''Messagebox window set'''
def get_timestamp():
    return datetime.now().strftime("%Y%m%d%H%M%S")
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
'''Process bar set'''
class ProgressDialog(QDialog):
    def __init__(self,info, parent=None):
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

        if parent:
            screen = parent.screen()
            screen_geometry = screen.availableGeometry()
            message_geometry = self.frameGeometry()
            message_geometry.moveCenter(screen_geometry.center())
            self.move(message_geometry.topLeft())


    def update_progress(self, value):
        self.progress_bar.setValue(value)
'''Process worker set'''
class Worker(QThread):
    progress = pyqtSignal(int)
    def __init__(self,sum_time):
        super().__init__()
        self.sumtime=sum_time
    def run(self):
        self.running = True
        global processing_or_not
        processing_or_not=True
        for i in range(100):
            if not self.running:
                break
            time.sleep(self.sumtime/100)
            self.progress.emit(i)
        while self.running:
            pass
    def stop(self):
        self.running = False
        global processing_or_not
        processing_or_not=False
'''Setup snipping tool'''
class SnippingThread(QThread):
    def __init__(self):
        super().__init__()
    def run(self):
        subprocess.run(['snippingtool'])
'''==================================Process function set=============================================='''
'''Sketch - Rescale - Rescale'''
class Process(QThread):
    def __init__(self,progress_dialog,file_dir,size_list,output_file):
        super().__init__()
        self.ui=ui
        self.progress_dialog = progress_dialog
        self.file_dir=file_dir
        self.size_list=size_list
        self.output_file=output_file
    def run(self):
        try:
            sum_size = os.path.getsize(self.file_dir)
            sum_time=5*sum_size/1000000+10
            self.worker = Worker(sum_time)
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            self.resize()
            self.worker.stop()
        except:
            traceback.print_exc()
    def resize(self):
        try:
            pool = multiprocessing.Pool(20)
            resize_pdf_multiprocessing(pool, self.file_dir, self.size_list[0], self.size_list[1], self.size_list[2],
                        self.size_list[3], self.size_list[4], self.size_list[5], self.output_file)
            pool.close()
            time.sleep(2)
            flatten_pdf(self.output_file,self.output_file)
            open_in_bluebeam(self.output_file)
        except:
            traceback.print_exc()
            return
'''Sketch - Align - Add tag'''
class Addtag(QThread):
    def __init__(self, progress_dialog,file_dir,gray_onoff,tags,tag_size_dict,tag_onoff):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.file_dir=file_dir
        self.gray_onoff=gray_onoff
        self.tag_onoff=tag_onoff
        self.tags=tags
        self.tag_size_dict=tag_size_dict
    def run(self):
        try:
            sum_time=200
            self.worker = Worker(sum_time)
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            if self.gray_onoff==True:
                self.gray_scale()
            if self.tag_onoff==True:
                self.tag()
            self.worker.stop()
            open_in_bluebeam(self.file_dir)
        except:
            traceback.print_exc()
            self.worker.stop()
    def gray_scale(self):
        try:
            pool = multiprocessing.Pool(20)
            grays_scale_pdf_multiprocessing(pool, self.file_dir)
            pool.close()
        except:
            traceback.print_exc()

    def tag(self):
        try:
            if self.tag_size_dict['output_scale']=="custom":
                tag="1:"+str(self.tag_size_dict['cus_list'][0])+"@"
            else:
                tag ="1:"+str(self.tag_size_dict['output_scale'])+"@"
            if self.tag_size_dict['output_size']=="custom":
                tag+=str(self.tag_size_dict['cus_list'][1])+':'+str(self.tag_size_dict['cus_list'][2])
            else:
                tag+=self.tag_size_dict['output_size']
            tag_x = 0
            tag_y = 0.1
            markup_dict={'A4':"Textbox-Size24",'A3':"Textbox-Size_A3", 'A2':"Textbox-Size_A2", 'A1':"Textbox-Size_A1", 'A0':"Textbox-Size_A0", 'A0-24':"Textbox-Size_A0_24",'A0-26':"Textbox-Size_A0_26", 'custom':"Textbox-Size_A0_26"}
            markup_type=markup_dict[self.tag_size_dict['output_size']]
            markup_width_dict={'A4':14,'A3':24, 'A2':24, 'A1':40, 'A0':50, 'A0-24':72,'A0-26':100, 'custom':100}
            markup_width_coe=markup_width_dict[self.tag_size_dict['output_size']]
            pool = multiprocessing.Pool(20)
            add_tags_multiprocessing(pool, self.file_dir, markup_type, tag_x, tag_y, tag, self.tags,markup_width_coe)
            pool.close()
        except:
            traceback.print_exc()
'''Drawing - Setup drawing - Update frame'''
class Changeframe(QThread):
    def __init__(self, progress_dialog,temp1, output_file, temp3, content_replace_dict_list, content_replace_dict_firstpage):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.temp1=temp1
        self.output_file=output_file
        self.temp3 =temp3
        self.content_replace_dict_list =content_replace_dict_list
        self.content_replace_dict_firstpage =content_replace_dict_firstpage
    def run(self):
        sum_time = 60
        self.worker = Worker(sum_time)
        try:
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            pages = get_number_of_page(self.output_file)
            for i in range(pages):
                content_replace_dict=self.content_replace_dict_list[i]
                PDFTools_v2.paste_markup_to_file(self.temp1, self.output_file, page_number=1,new_page_number=i+1,offset=(0, 0),content_replace_dict=content_replace_dict)
            time.sleep(2)
            combine_pdf([self.temp3,self.output_file],self.output_file)
            PDFTools_v2.paste_markup_to_file(self.temp1, self.output_file, page_number=1, new_page_number=1, offset=(0, 0),content_replace_dict=self.content_replace_dict_firstpage)
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
'''Drawing - Setup drawing - Paste logo'''
class Pastelogo(QThread):
    def __init__(self, progress_dialog,logo_dir,file_dir,paper_size):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.logo_dir=logo_dir
        self.file_dir=file_dir
        self.paper_size=paper_size
    def run(self):
        sum_time = 600
        self.worker = Worker(sum_time)
        try:
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            pages = get_number_of_page(self.file_dir)
            for i in range(pages):
                insert_logo_into_pdf(self.file_dir, self.logo_dir, i, self.paper_size)
                time.sleep(2)
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
'''Overlay - Service - Grayscale'''
class Gray(QThread):
    def __init__(self, progress_dialog,file_dir):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.file_dir=file_dir
    def run(self):
        try:
            sum_time=30
            self.worker = Worker(sum_time)
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            self.gray_scale()
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
    def gray_scale(self):
        try:
            pool = multiprocessing.Pool(20)
            grays_scale_pdf_multiprocessing(pool, self.file_dir)
            pool.close()
        except:
            traceback.print_exc()
'''Combine/Overlay - Combine'''
class Combine(QThread):
    def __init__(self, progress_dialog,combinefile_list_ab, combine_pdf_dir):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.combinefile_list_ab = combinefile_list_ab
        self.combine_pdf_dir = combine_pdf_dir
    def run(self):
        try:
            sum_time=10
            self.worker = Worker(sum_time)
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            combine_pdf(self.combinefile_list_ab, self.combine_pdf_dir)
            self.worker.stop()
        except:
            traceback.print_exc()
'''Sketch - Align - Align'''
class Align(QThread):
    def __init__(self, progress_dialog,file_dir):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.file_dir = file_dir
    def run(self):
        sum_time = 120
        self.worker = Worker(sum_time)
        pool = multiprocessing.Pool(20)
        try:
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            align_multiprocessing(pool, self.file_dir)
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
            pool.close()
'''Process to PC'''
class Process2PC(QThread):
    def __init__(self, progress_dialog,process_info):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.process_info = process_info
    def run(self):
        try:
            sumtime_dict={"LumiColor":1200,"Overlay_PC":1200,"Snapsketch":1200,"Movesketch":120,"Copymarkup":1200,"Plot":1200,"Remove":300,"Align_combine":1200}
            category=self.process_info[0]
            file=self.process_info[1]
            sum_time = sumtime_dict[category]
            self.worker = Worker(sum_time)
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            if category=="LumiColor":
                lumicolor_pc_wait(file)
            elif category=="Overlay_PC":
                overlay_pc_wait(file)
            elif category=="Snapsketch":
                setupdrawing_wait(file)
            elif category=="Movesketch":
                pdf_move_center(file, self.process_info[2][0])
            elif category=="Copymarkup":
                markup_wait(file)
            elif category=="Plot":
                plot_wait(file)
            elif category=="Remove":
                remove_wait(file)
            elif category=="Align_combine":
                aligncombine_wait(file)
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
'''Waitinginline process'''
class Waitinline(QThread):
    def __init__(self, progress_dialog,file_time):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.file_time = file_time
    def run(self):
        sum_time = 1200
        self.worker = Worker(sum_time)
        try:
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            for i in range(120):
                time.sleep(10)
                process_now=f_waitinline(self.file_time)
                if process_now=='ok now':
                    break
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
'''Drawing - Update frame'''
class UpdateFrame(QThread):
    def __init__(self, progress_dialog,pdf_file, last_drawing_page, drawingframe_context_list,chk_by,drw_by,issue_type,pro_no,pro_name,new_rev_list_all,scale_a1,scale_b,scale_a2,last_cover_page):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.pdf_file=pdf_file
        self.last_drawing_page =last_drawing_page
        self.drawingframe_context_list =drawingframe_context_list
        self.chk_by =chk_by
        self.drw_by =drw_by
        self.issue_type =issue_type
        self.pro_no=pro_no
        self.pro_name =pro_name
        self.new_rev_list_all =new_rev_list_all
        self.scale_a1 =scale_a1
        self.scale_b =scale_b
        self.scale_a2 =scale_a2
        self.last_cover_page=last_cover_page
    def run(self):
        sum_time = 120
        self.worker = Worker(sum_time)
        try:
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            for i in range(self.last_drawing_page):
                new_drawing_content = self.drawingframe_context_list[i]
                new_drawing_content['CHECKED BY'][1][2] = self.chk_by
                new_drawing_content['DRAWN BY'][1][2] = self.drw_by
                new_drawing_content['TYPE OF ISSUE'][1][2] = self.issue_type
                new_drawing_content['PROJECT No.'][1][2] = self.pro_no
                new_drawing_content['PROJECT'][1][2] = self.pro_name
                for j in range(len(self.new_rev_list_all)):
                    if j >= len(new_drawing_content['VERSION']):
                        new_drawing_content['VERSION'].append(
                            {
                                'REV': ['', [0, 0, self.new_rev_list_all[j][0]]],
                                'REVISION DESCRIPTION': ['', [0, 0, self.new_rev_list_all[j][1]]],
                                'DRW': ['', [0, 0, self.new_rev_list_all[j][2]]],
                                'CHK': ['', [0, 0, self.new_rev_list_all[j][3]]],
                                'DATE Y.M.D': ['', [0, 0, self.new_rev_list_all[j][4]]]
                            })
                    else:
                        new_drawing_content['VERSION'][j]['REV'][1][2] = self.new_rev_list_all[j][0]
                        new_drawing_content['VERSION'][j]['REVISION DESCRIPTION'][1][2] = self.new_rev_list_all[j][1]
                        new_drawing_content['VERSION'][j]['DRW'][1][2] = self.new_rev_list_all[j][2]
                        new_drawing_content['VERSION'][j]['CHK'][1][2] = self.new_rev_list_all[j][3]
                        new_drawing_content['VERSION'][j]['DATE Y.M.D'][1][2] = self.new_rev_list_all[j][4]
                if i < self.last_cover_page:
                    new_drawing_content['SCALE'][1][2] = self.scale_a1 + self.scale_b
                else:
                    new_drawing_content['SCALE'][1][2] = self.scale_a2 + self.scale_b
                PDFTools_v2.set_structured_markups(self.pdf_file, i + 1, new_drawing_content, (0, -11.6))
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()
'''Drawing - Generate link'''
class Link(QThread):
    def __init__(self, progress_dialog, zip_name,folder_for_zip):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.zip_name = zip_name
        self.folder_for_zip=folder_for_zip
    def run(self):
        sum_time = 300
        self.worker = Worker(sum_time)
        try:
            self.worker.progress.connect(self.progress_dialog.update_progress)
            self.worker.finished.connect(self.progress_dialog.accept)
            self.worker.start()
            compress_directory_to_zip(self.zip_name, self.folder_for_zip)
            drive_json=os.path.join(conf["database_dir"],'useful-cathode-435107-r6-c6ea48c620cb.json')
            google_drive = GoogleDrive(drive_json)
            global link_path
            link_path = google_drive.upload_for_someone_read(self.zip_name)
        except:
            traceback.print_exc()
        finally:
            self.worker.stop()

'''Main function'''
class Stats():
    working_dir = "P:\\"
    def __init__(self,ui):
        super().__init__()
        self.ui=ui
        # self.ui.setWindowState(Qt.WindowMaximized)
        # self.ui.showMaximized()
        '''Global parameter'''
        self.inv_state_dict = {}
        self.bill_state_dict = {}
        self.bill_gst_dict = {}
        self.search_history = ['',]
        self.align_files = []
        self.projects_all = {}
        self.project_dict = {}
        self.service_name=''
        self.last_tab=0
        self.checkbox_search_gst_state=True
        # self.project_name = "-test-"
        self.backup_disk = r"B:\02.Copilot\tmp"
        '''Clipboard set'''
        self.clipboard_history=''
        '''Tab widget position set'''
        self.ui.tabWidget.setTabPosition(QTabWidget.South)
        '''Network manager'''
        self.manager = QNetworkAccessManager(self.ui)
        '''Load search bar content'''
        self.generate_search_bar(order=1)
        '''Search page set'''
        self.change_table(self.projects_all)
        '''Lock function'''
        self.lock = threading.Lock()
        '''Clear template folder'''
        directory_path = r'C:\Copilot_template'
        create_or_clear_directory(directory_path)
        '''Drawing add plot folder'''
        self.ui.checkbox_plot_folder.stateChanged.connect(self.f_text_drawing_plot_filefrom)
        '''Auto search'''
        self.function_to_run()
        '''Search table copy function'''
        self.ui.table_search_1.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.ui.table_search_1.setContextMenuPolicy(Qt.CustomContextMenu)
        self.ui.table_search_1.customContextMenuRequested.connect(self.showContextMenu)
        self.copyShortcut = QShortcut(QKeySequence("Ctrl+C"), self.ui)
        self.copyShortcut.activated.connect(self.copySelection)
        '''Timesheet user table select'''
        self.ui.table_timesheet_user.setSelectionBehavior(QTableWidget.SelectRows)

        self.ui.table_timesheet_user.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: lightblue;
                color: black;
            }
        """)

        '''Modify color init set'''
        self.renderer1 = QSvgRenderer()
        self.renderer2= QSvgRenderer()
        '''Modify color toolbox 1'''
        # self.frame_mode = self.ui.findChild(QFrame, "frame_modify_color")
        # self.frame_sourcecolor1 = self.ui.findChild(QFrame, "frame_sourcecolor_from")
        # if self.frame_sourcecolor1.layout() is None:
        #     self.layout1 = QVBoxLayout()
        #     self.frame_sourcecolor1.setLayout(self.layout1)
        # else:
        #     self.layout1 = self.frame_sourcecolor1.layout()
        # '''Modify color toolbox 2'''
        # self.frame_sourcecolor2 = self.ui.findChild(QFrame, "frame_sourcecolor_to")
        # if self.frame_sourcecolor2.layout() is None:
        #     self.layout2 = QVBoxLayout()
        #     self.frame_sourcecolor2.setLayout(self.layout2)
        # else:
        #     self.layout2 = self.frame_sourcecolor2.layout()
        '''Copy markup preview frame'''
        self.frame_copymk_from= self.ui.findChild(QFrame, "frame_copymk_from")
        if self.frame_copymk_from.layout() is None:
            self.layout_copymk_from = QVBoxLayout()
            self.frame_copymk_from.setLayout(self.layout_copymk_from)
        else:
            self.layout_copymk_from = self.frame_copymk_from.layout()
        self.frame_copymk_to= self.ui.findChild(QFrame, "frame_copymk_to")
        if self.frame_copymk_to.layout() is None:
            self.layout_copymk_to = QVBoxLayout()
            self.frame_copymk_to.setLayout(self.layout_copymk_to)
        else:
            self.layout_copymk_to = self.frame_copymk_to.layout()
        '''Overlay update-grayscale connect'''
        self.ui.checkbox_overlay_update.stateChanged.connect(self.f_checkbox_overlay_update)
        self.ui.checkbox_search_gst.stateChanged.connect(self.f_checkbox_search_gst)
        '''Set keyboard check'''
        keyboard.add_hotkey('ctrl+alt+z', self.function_to_run)
        shortcut2 = QShortcut(QKeySequence('Ctrl+Alt+Shift+A'), self.ui)
        shortcut2.activated.connect(self.function_to_run2)
        '''Press button set'''
        self.button_to_function_copilot = {
            self.ui.button_all_1_clearup: self.f_button_all_clearup,
            self.ui.button_all_2_search: self.f_all_search,
            self.ui.button_all_3_openasana: self.f_all_openasana,
            self.ui.button_search_1_refresh: self.f_search_refresh2,
            self.ui.button_logo_open: self.f_open_logo_folder,
            self.ui.button_drawing_changeout: self.f_button_drawing_changeout,
            self.ui.button_combine_6_combine: self.f_combine_combine,
            # self.ui.button_sketch_rescale_6_process: self.f_sketch_rescale_process,


            # self.ui.button_sketch_rescale_6_process: self.test_sketch_rescale,
            # self.ui.button_sketch_rescale_6_change_7: self.test_sketch_align,
            # self.ui.button_sketch_process: self.test_sketch_color_process,
            # self.ui.button_sketch_rescale_6_change_4: self.test_sketch_copy_markup,
            # self.ui.button_sketch_rescale_6_change_5: self.test_sketch_erase,

            self.ui.button_overlay_rescale: self.f_button_overlay_rescale,
            self.ui.button_filing_set: self.f_filing_set,
            self.ui.button_drawing_logo: self.f_getlogo,
            self.ui.button_drawing_gettem: self.f_button_drawing_gettem,
            self.ui.button_changesketchdir: self.f_button_changesketchdir,
            self.ui.button_changeupdatefile: self.f_button_changeupdatefile,
            # self.ui.button_sketch_process: self.f_button_sketch_process,
            self.ui.button_overlay_set: self.f_button_overlay_set,
            self.ui.button_overlay_mech_combine: self.f_overlay_mech_combine,
            self.ui.button_overlay_service1_overlay: self.f_button_overlay_service1_overlay,
            self.ui.button_overlay_markupfrom_change: self.f_button_overlay_markupfrom_change,
            self.ui.button_overlay_markupto_change: self.f_button_overlay_markupto_change,
            self.ui.button_overlay_copymarkup: self.f_button_overlay_copymarkup,
            self.ui.button_drawing_chooselogo: self.f_button_drawing_chooselogo,
            self.ui.button_drawing_update: self.f_button_drawing_update,
            self.ui.button_drawing_plot: self.f_button_drawing_plot,
            self.ui.button_drawing_plot_filefrom: self.f_button_drawing_plot_filefrom,
            self.ui.button_issue_copylink:self.f_button_issue_copylink,
            self.ui.button_issue_link:self.f_button_issue_link,
            self.ui.button_drawing_plot_fileclear:self.f_button_drawing_plot_fileclear,
            self.ui.button_issue_openfolder:self.f_button_issue_openfolder,
            self.ui.button_open_sketch: self.f_button_open_sketch,

            # self.ui.button_combine_3_open: partial(self.f_tablefile_open, table_name="Table_combine"),
            # self.ui.button_sketch_rescale_3_open: partial(self.f_tablefile_open, table_name="Table_rescale_O"),
            # self.ui.button_rescale_A_open: partial(self.f_tablefile_open, table_name="Table_rescale_A"),
            # self.ui.button_rescale_B_open: partial(self.f_tablefile_open, table_name="Table_rescale_B"),
            # self.ui.button_rescale_C_open: partial(self.f_tablefile_open, table_name="Table_rescale_C"),
            # self.ui.button_rescale_D_open: partial(self.f_tablefile_open, table_name="Table_rescale_D"),
            self.ui.button_overlay_service1_open: partial(self.f_tablefile_open, table_name="Table_overlay_service"),
            self.ui.button_overlay_mech_open: partial(self.f_tablefile_open, table_name="Table_overlay_mech"),

            self.ui.button_all_4_openfolder: partial(self.f_projectfolder_open, folder_name="Project"),
            self.ui.button_all_4_openbackfolder: partial(self.f_projectfolder_open, folder_name="Backup"),

            # self.ui.button_combine_1_up: partial(self.f_tableitem_up, table_name="Table_combine"),
            # self.ui.button_sketch_rescale_1_up: partial(self.f_tableitem_up, table_name="Table_rescale_O"),
            # self.ui.button_rescale_A_up: partial(self.f_tableitem_up, table_name="Table_rescale_A"),
            # self.ui.button_rescale_B_up: partial(self.f_tableitem_up, table_name="Table_rescale_B"),
            # self.ui.button_rescale_C_up: partial(self.f_tableitem_up, table_name="Table_rescale_C"),
            # self.ui.button_rescale_D_up: partial(self.f_tableitem_up, table_name="Table_rescale_D"),
            self.ui.button_overlay_service1_up: partial(self.f_tableitem_up, table_name="Table_overlay_service"),
            self.ui.button_overlay_mech_up: partial(self.f_tableitem_up, table_name="Table_overlay_mech"),

            # self.ui.button_combine_2_down: partial(self.f_tableitem_down, table_name="Table_combine"),
            # self.ui.button_sketch_rescale_2_down: partial(self.f_tableitem_down, table_name="Table_rescale_O"),
            # self.ui.button_rescale_A_down: partial(self.f_tableitem_down, table_name="Table_rescale_A"),
            # self.ui.button_rescale_B_down: partial(self.f_tableitem_down, table_name="Table_rescale_B"),
            # self.ui.button_rescale_C_down: partial(self.f_tableitem_down, table_name="Table_rescale_C"),
            # self.ui.button_rescale_D_down: partial(self.f_tableitem_down, table_name="Table_rescale_D"),
            self.ui.button_overlay_service1_down: partial(self.f_tableitem_down, table_name="Table_overlay_service"),
            self.ui.button_overlay_mech_down: partial(self.f_tableitem_down, table_name="Table_overlay_mech"),

            # self.ui.button_combine_4_delete: partial(self.f_tableitem_delete, table_name="Table_combine"),
            # self.ui.button_sketch_rescale_4_delete: partial(self.f_tableitem_delete, table_name="Table_rescale_O"),
            # self.ui.button_rescale_A_delete: partial(self.f_tableitem_delete, table_name="Table_rescale_A"),
            # self.ui.button_rescale_B_delete: partial(self.f_tableitem_delete, table_name="Table_rescale_B"),
            # self.ui.button_rescale_C_delete: partial(self.f_tableitem_delete, table_name="Table_rescale_C"),
            # self.ui.button_rescale_D_delete: partial(self.f_tableitem_delete, table_name="Table_rescale_D"),
            self.ui.button_overlay_service1_delete: partial(self.f_tableitem_delete, table_name="Table_overlay_service"),
            self.ui.button_overlay_mech_delete: partial(self.f_tableitem_delete, table_name="Table_overlay_mech"),

            # self.ui.button_combine_5_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_combine"),
            # self.ui.button_sketch_rescale_5_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_rescale_O"),
            # self.ui.button_rescale_A_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_rescale_A"),
            # self.ui.button_rescale_B_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_rescale_B"),
            # self.ui.button_rescale_C_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_rescale_C"),
            # self.ui.button_rescale_D_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_rescale_D"),
            self.ui.button_overlay_service1_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_overlay_service"),
            self.ui.button_overlay_mech_deleteall: partial(self.f_tableitem_deleteall, table_name="Table_overlay_mech"),

            self.ui.button_overlay_change1: partial(self.f_folder_change, folder_text_name="Overlay output"),
            self.ui.button_combine_1_change: partial(self.f_folder_change, folder_text_name="Combine output"),
            self.ui.button_overlay_mech_changefolder: partial(self.f_folder_change, folder_text_name="Overlay mech"),
            self.ui.button_filing_change: partial(self.f_folder_change, folder_text_name="Filing"),

            # self.ui.button_sketch_rescale_6_change: partial(self.f_folder_change, folder_text_name="Rescale O"),
            # self.ui.button_sketch_rescale_6_change: partial(self.open_function,
            #                                                 input_widget=self.ui.textedit_drawing_plot_filefrom_2,
            #                                                 folder_path=r"B:\02.Copilot\tmp"),







            self.ui.button_issue_changefolder:partial(self.f_folder_change, folder_text_name="Issue"),
            self.ui.button_drawing_plot_fileto: partial(self.f_folder_change, folder_text_name="Plot"),


            # self.ui.button_align_change_1: partial(self.f_file_change, file_dir_name="Rescale O"),
            # self.ui.button_align_change_a: partial(self.f_file_change, file_dir_name="Rescale A"),
            # self.ui.button_align_change_b: partial(self.f_file_change, file_dir_name="Rescale B"),
            # self.ui.button_align_change_c: partial(self.f_file_change, file_dir_name="Rescale C"),
            # self.ui.button_align_change_d: partial(self.f_file_change, file_dir_name="Rescale D"),

            self.ui.button_timesheet_sync: self.f_button_timesheet_sync,
            self.ui.button_timesheet_syncback: self.f_button_timesheet_syncback,

            self.ui.button_search_hide_1: partial(self.f_button_search_hide, region=1),
            self.ui.button_search_hide_2: partial(self.f_button_search_hide, region=2),
            self.ui.button_search_hide_3: partial(self.f_button_search_hide, region=3),
            self.ui.button_search_hide_4: partial(self.f_button_search_hide, region=4),
            self.ui.button_search_hide_5: partial(self.f_button_search_hide, region=5),
            self.ui.button_search_hide_6: partial(self.f_button_search_hide, region=6),


            self.ui.button_log_login:self.f_button_log_login,
            self.ui.button_log_confirm:self.f_button_log_confirm,
            self.ui.button_log_cancel:self.f_button_log_cancel,

            self.ui.button_timesheet_cal:self.f_calculate_sum_time_all,

        }
        self.connect_buttons()


        self.sketch_tab = Sketch_Tab(self.ui)
        self.drawing_tab = Drawing_Tab(self.ui)


        '''Button group set'''
        # self.button_group1 = QButtonGroup(self.ui)
        # button_group1_list = [self.ui.radiobutton_sketch_rescale_1_osc50, self.ui.radiobutton_sketch_rescale_1_osc75,
        #                       self.ui.radiobutton_sketch_rescale_1_osc100,
        #                       self.ui.radiobutton_sketch_rescale_1_osc150, self.ui.radiobutton_sketch_rescale_1_osc200,
        #                       self.ui.radiobutton_sketch_rescale_1_oscx]
        # self.button_group2 = QButtonGroup(self.ui)
        # button_group2_list = [self.ui.radiobutton_sketch_size_1_a4, self.ui.radiobutton_sketch_size_1_a3,
        #                       self.ui.radiobutton_sketch_size_1_a2,
        #                       self.ui.radiobutton_sketch_size_1_a1, self.ui.radiobutton_sketch_size_1_a0,
        #                       self.ui.radiobutton_sketch_size_1_a024,
        #                       self.ui.radiobutton_sketch_size_1_a026, self.ui.radiobutton_sketch_size_1_ax]
        # self.button_group_a1 = QButtonGroup(self.ui)
        # button_group_a1_list = [self.ui.radiobutton_a_1, self.ui.radiobutton_a_2, self.ui.radiobutton_a_3,
        #                         self.ui.radiobutton_a_4, self.ui.radiobutton_a_5, self.ui.radiobutton_a_6]
        # self.button_group_a2 = QButtonGroup(self.ui)
        # button_group_a2_list = [self.ui.radiobutton_a_7, self.ui.radiobutton_a_8, self.ui.radiobutton_a_9,
        #                         self.ui.radiobutton_a_10, self.ui.radiobutton_a_11,
        #                         self.ui.radiobutton_a_12, self.ui.radiobutton_a_13, self.ui.radiobutton_a_14, ]
        # self.button_group_b1 = QButtonGroup(self.ui)
        # button_group_b1_list = [self.ui.radiobutton_b_1, self.ui.radiobutton_b_2, self.ui.radiobutton_b_3,
        #                         self.ui.radiobutton_b_4, self.ui.radiobutton_b_5, self.ui.radiobutton_b_6]
        # self.button_group_b2 = QButtonGroup(self.ui)
        # button_group_b2_list = [self.ui.radiobutton_b_7, self.ui.radiobutton_b_8, self.ui.radiobutton_b_9,
        #                         self.ui.radiobutton_b_10, self.ui.radiobutton_b_11, self.ui.radiobutton_b_12,
        #                         self.ui.radiobutton_b_13, self.ui.radiobutton_b_14]
        # self.button_group_c1 = QButtonGroup(self.ui)
        # button_group_c1_list = [self.ui.radiobutton_c_1, self.ui.radiobutton_c_2, self.ui.radiobutton_c_3,
        #                         self.ui.radiobutton_c_4, self.ui.radiobutton_c_5, self.ui.radiobutton_c_6]
        # self.button_group_c2 = QButtonGroup(self.ui)
        # button_group_c2_list = [self.ui.radiobutton_c_7, self.ui.radiobutton_c_8, self.ui.radiobutton_c_9,
        #                         self.ui.radiobutton_c_10,
        #                         self.ui.radiobutton_c_11, self.ui.radiobutton_c_12, self.ui.radiobutton_c_13,
        #                         self.ui.radiobutton_c_14]
        # self.button_group_d1 = QButtonGroup(self.ui)
        # button_group_d1_list = [self.ui.radiobutton_d_1, self.ui.radiobutton_d_2, self.ui.radiobutton_d_3,
        #                         self.ui.radiobutton_d_4, self.ui.radiobutton_d_5, self.ui.radiobutton_d_6]
        # self.button_group_d2 = QButtonGroup(self.ui)
        # button_group_d2_list = [self.ui.radiobutton_d_7, self.ui.radiobutton_d_8, self.ui.radiobutton_d_9,
        #                         self.ui.radiobutton_d_10, self.ui.radiobutton_d_11,
        #                         self.ui.radiobutton_d_12, self.ui.radiobutton_d_13, self.ui.radiobutton_d_14]
        # self.button_group3 = QButtonGroup(self.ui)
        # button_group3_list = [self.ui.radiobutton_sketch_rescale_2_osc50, self.ui.radiobutton_sketch_rescale_2_osc100,
        #                       self.ui.radiobutton_sketch_rescale_2_oscx, ]
        # self.button_group4 = QButtonGroup(self.ui)
        # button_group4_list = [self.ui.radiobutton_sketch_size_2_a4, self.ui.radiobutton_sketch_size_2_a3,
        #                       self.ui.radiobutton_sketch_size_2_a2,
        #                       self.ui.radiobutton_sketch_size_2_a1, self.ui.radiobutton_sketch_size_2_a0,
        #                       self.ui.radiobutton_sketch_size_2_a024,
        #                       self.ui.radiobutton_sketch_size_2_a026, self.ui.radiobutton_sketch_size_2_x]
        self.button_group_tag = QButtonGroup(self.ui)
        self.button_group_tag.setExclusive(False)
        # self.button_group_tag_boxlist = [self.ui.box_tag_a_1, self.ui.box_tag_a_2, self.ui.box_tag_a_3, self.ui.box_tag_a_4,
        #                          self.ui.box_tag_a_4_1,
        #                          self.ui.box_tag_a_4_2, self.ui.checkbox_sketch_colorprocess_1_basement2,
        #                          self.ui.checkbox_sketch_colorprocess_2_basement1,
        #                          self.ui.checkbox_sketch_colorprocess_3_groundlevel,
        #                          self.ui.checkbox_sketch_colorprocess_4_level1,
        #                          self.ui.checkbox_sketch_colorprocess_5_level2,
        #                          self.ui.checkbox_sketch_colorprocess_6_level3,
        #                          self.ui.checkbox_sketch_colorprocess_7_level4,
        #                          self.ui.checkbox_sketch_colorprocess_8_level5,
        #                          self.ui.checkbox_sketch_colorprocess_9_level6,
        #                          self.ui.checkbox_sketch_colorprocess_10_level7,
        #                          self.ui.checkbox_sketch_colorprocess_11_level8,
        #                          self.ui.box_tag_a_5, self.ui.box_tag_a_6, self.ui.box_tag_a_7, self.ui.box_tag_a_8,
        #                          self.ui.box_tag_a_8_1, self.ui.box_tag_a_8_2,
        #                          self.ui.box_tag_a_8_3, self.ui.box_tag_a_8_4, self.ui.box_tag_a_8_5,
        #                          self.ui.box_tag_a_8_6, self.ui.checkbox_sketch_colorprocess_11_roof]
        # self.button_group_tag_textlist=[self.ui.text_sketch_colorprocess_1,self.ui.text_sketch_colorprocess_2,
        #                                 self.ui.text_sketch_colorprocess_3,self.ui.text_sketch_colorprocess_4,
        #                                 self.ui.text_sketch_colorprocess_4_1,self.ui.text_sketch_colorprocess_4_2,
        #                                 self.ui.text_color_tag_1,self.ui.text_color_tag_2,
        #                                 self.ui.text_color_tag_3,self.ui.text_color_tag_4,
        #                                 self.ui.text_color_tag_5,self.ui.text_color_tag_6,
        #                                 self.ui.text_color_tag_7,self.ui.text_color_tag_8,
        #                                 self.ui.text_color_tag_9,self.ui.text_color_tag_10,
        #                                 self.ui.text_color_tag_11,self.ui.text_sketch_colorprocess_5,
        #                                 self.ui.text_sketch_colorprocess_6,self.ui.text_sketch_colorprocess_7,
        #                                 self.ui.text_sketch_colorprocess_8,self.ui.text_sketch_colorprocess_8_1,
        #                                 self.ui.text_sketch_colorprocess_8_2,self.ui.text_sketch_colorprocess_8_3,
        #                                 self.ui.text_sketch_colorprocess_8_4,self.ui.text_sketch_colorprocess_8_5,
        #                                 self.ui.text_sketch_colorprocess_8_6,self.ui.text_color_tag_12]
        self.button_tag_group_num = 28
        self.button_group_tag2 = QButtonGroup(self.ui)
        self.button_group_tag2.setExclusive(False)
        self.button_group_tag2_boxlist = [self.ui.box_tag_b_1, self.ui.box_tag_b_2, self.ui.box_tag_b_3, self.ui.box_tag_b_4,
                                  self.ui.box_tag_b_4_1, self.ui.box_tag_b_4_2,
                                  self.ui.checkbox_drawing_colorprocess_1_basement2,
                                  self.ui.checkbox_drawing_colorprocess_2_basement1,
                                  self.ui.checkbox_drawing_colorprocess_3_groundlevel,
                                  self.ui.checkbox_drawing_colorprocess_4_level1,
                                  self.ui.checkbox_drawing_colorprocess_5_level2,
                                  self.ui.checkbox_drawing_colorprocess_6_level3,
                                  self.ui.checkbox_drawing_colorprocess_7_level4,
                                  self.ui.checkbox_drawing_colorprocess_8_level5,
                                  self.ui.checkbox_drawing_colorprocess_9_level6,
                                  self.ui.checkbox_drawing_colorprocess_10_level7,
                                  self.ui.checkbox_drawing_colorprocess_11_level8, self.ui.box_tag_b_5,
                                  self.ui.box_tag_b_6,
                                  self.ui.box_tag_b_7, self.ui.box_tag_b_8, self.ui.box_tag_b_8_1,
                                  self.ui.box_tag_b_8_2, self.ui.box_tag_b_8_3, self.ui.box_tag_b_8_4,
                                  self.ui.box_tag_b_8_5,
                                  self.ui.box_tag_b_8_6, self.ui.checkbox_drawing_colorprocess_11_roof]
        self.button_group_tag2_textlist=[self.ui.text_drawing_colorprocess_1,self.ui.text_drawing_colorprocess_2,
                                      self.ui.text_drawing_colorprocess_3,self.ui.text_drawing_colorprocess_4,
                                      self.ui.text_drawing_colorprocess_4_1,self.ui.text_drawing_colorprocess_4_2,
                                      self.ui.text_drawing_tag_1,self.ui.text_drawing_tag_2,
                                      self.ui.text_drawing_tag_3,self.ui.text_drawing_tag_4,
                                      self.ui.text_drawing_tag_5,self.ui.text_drawing_tag_6,
                                      self.ui.text_drawing_tag_7,self.ui.text_drawing_tag_8,
                                      self.ui.text_drawing_tag_9,self.ui.text_drawing_tag_10,
                                      self.ui.text_drawing_tag_11,self.ui.text_drawing_colorprocess_5,
                                      self.ui.text_drawing_colorprocess_6,self.ui.text_drawing_colorprocess_7,
                                      self.ui.text_drawing_colorprocess_8,self.ui.text_drawing_colorprocess_8_1,
                                      self.ui.text_drawing_colorprocess_8_2,self.ui.text_drawing_colorprocess_8_3,
                                      self.ui.text_drawing_colorprocess_8_4,self.ui.text_drawing_colorprocess_8_5,
                                      self.ui.text_drawing_colorprocess_8_6,self.ui.text_drawing_tag_12]
        self.button_group_cate = QButtonGroup(self.ui)
        button_group_cate_list = [self.ui.radiobutton_drawing_rest, self.ui.radiobutton_drawing_other]
        self.button_group_des = QButtonGroup(self.ui)
        button_group_des_list = [self.ui.radiobutton_drawing_approval, self.ui.radiobutton_drawing_construction,
                                 self.ui.checkbox_issue_type]
        self.button_overlay_group = QButtonGroup(self.ui)
        button_overlay_group_list = [self.ui.radiobutton_servicetype1, self.ui.radiobutton_servicetype2,
                                     self.ui.radiobutton_servicetype3,
                                     self.ui.radiobutton_servicetype4, self.ui.radiobutton_servicetype5,
                                     self.ui.radiobutton_servicetype6,
                                     self.ui.radiobutton_servicetype7, self.ui.radiobutton_servicetype8]
        self.button_group_overlay_service1_os = QButtonGroup(self.ui)
        button_group_overlay_service1_os_list = [self.ui.radiobutton_overlay_service1_os_1,
                                                 self.ui.radiobutton_overlay_service1_os_2,
                                                 self.ui.radiobutton_overlay_service1_os_3,
                                                 self.ui.radiobutton_overlay_service1_os_4,
                                                 self.ui.radiobutton_overlay_service1_os_5,
                                                 self.ui.radiobutton_overlay_service1_os_6, ]
        self.button_group_overlay_service1_outs = QButtonGroup(self.ui)
        button_group_overlay_service1_outs_list = [self.ui.radiobutton_overlay_service1_outs_1,
                                                   self.ui.radiobutton_overlay_service1_outs_2,
                                                   self.ui.radiobutton_overlay_service1_outs_3]
        self.button_group_overlay_service1_orisize = QButtonGroup(self.ui)
        button_group_overlay_service1_orisize_list = [self.ui.radiobutton_overlay_service1_orisize_1,
                                                      self.ui.radiobutton_overlay_service1_orisize_2,
                                                      self.ui.radiobutton_overlay_service1_orisize_3,
                                                      self.ui.radiobutton_overlay_service1_orisize_4,
                                                      self.ui.radiobutton_overlay_service1_orisize_5,
                                                      self.ui.radiobutton_overlay_service1_orisize_6,
                                                      self.ui.radiobutton_overlay_service1_orisize_7,
                                                      self.ui.radiobutton_overlay_service1_orisize_8]
        self.button_group_overlay_service1_outsize = QButtonGroup(self.ui)
        button_group_overlay_service1_outsize_list = [self.ui.radiobutton_overlay_service1_outsize_1,
                                                      self.ui.radiobutton_overlay_service1_outsize_2,
                                                      self.ui.radiobutton_overlay_service1_outsize_3,
                                                      self.ui.radiobutton_overlay_service1_outsize_4,
                                                      self.ui.radiobutton_overlay_service1_outsize_5,
                                                      self.ui.radiobutton_overlay_service1_outsize_6,
                                                      self.ui.radiobutton_overlay_service1_outsize_7,
                                                      self.ui.radiobutton_overlay_service1_outsize_8]
        self.button_group_modifycolor_sketch = QButtonGroup(self.ui)
        # button_group_modifycolor_sketch_list = [self.ui.radiobutton_color_sketch_o, self.ui.radiobutton_color_sketch_a,
        #                                         self.ui.radiobutton_color_sketch_b, self.ui.radiobutton_color_sketch_c,
        #                                         self.ui.radiobutton_color_sketch_d]


        self.button_overlay_customize_group = QButtonGroup(self.ui)
        button_overlay_customize_group_list = [self.ui.radiobutton_cus_1, self.ui.radiobutton_cus_2,self.ui.radiobutton_cus_3]

        '''Button group dict'''
        self.button_group_dict = {
            # self.button_group1: (button_group1_list, partial(self.f_buttongroup_scale, type="Sketch O")),
            #                       self.button_group2: (button_group2_list, partial(self.f_buttongroup_size_clicked, type="Sketch O origin")),
                                  # self.button_group_a1: (button_group_a1_list, partial(self.f_buttongroup_scale, type="Sketch A")),
                                  # self.button_group_a2: (button_group_a2_list, partial(self.f_buttongroup_size_clicked, type="Sketch A origin")),
                                  # self.button_group_b1: (button_group_b1_list, partial(self.f_buttongroup_scale, type="Sketch B")),
                                  # self.button_group_b2: (button_group_b2_list, partial(self.f_buttongroup_size_clicked, type="Sketch B origin")),
                                  # self.button_group_c1: (button_group_c1_list, partial(self.f_buttongroup_scale, type="Sketch C")),
                                  # self.button_group_c2: (button_group_c2_list, partial(self.f_buttongroup_size_clicked, type="Sketch C origin")),
                                  # self.button_group_d1: (button_group_d1_list, partial(self.f_buttongroup_scale, type="Sketch D")),
                                  # self.button_group_d2: (button_group_d2_list, partial(self.f_buttongroup_size_clicked, type="Sketch D origin")),
                                  # self.button_group3: (button_group3_list, partial(self.f_buttongroup_scale, type="Sketch output")),
                                  # self.button_group4: (button_group4_list, self.buttongroup4_clicked),
                                  # self.button_group_tag: (self.button_group_tag_boxlist, self.buttongroup_tag_click),
                                  self.button_group_tag2: (self.button_group_tag2_boxlist, self.buttongroup_tag_click2),
                                  self.button_group_cate: (button_group_cate_list, ''),
                                  self.button_group_des: (button_group_des_list, ''),
                                  self.button_overlay_group: (button_overlay_group_list, self.f_buttongroup_overlay_clicked),
                                  self.button_group_overlay_service1_os: (button_group_overlay_service1_os_list, partial(self.f_buttongroup_scale, type="Overlay origin")),
                                  self.button_group_overlay_service1_outs: (button_group_overlay_service1_outs_list, partial(self.f_buttongroup_scale, type="Overlay output")),
                                  # self.button_group_modifycolor_sketch: ("", self.f_button_group_modifycolor_sketch),
                                  self.button_group_overlay_service1_orisize: (button_group_overlay_service1_orisize_list,partial(self.f_buttongroup_size_clicked, type="Overlay origin")),
                                  self.button_group_overlay_service1_outsize: (button_group_overlay_service1_outsize_list,partial(self.f_buttongroup_size_clicked, type="Overlay output")),
                                  self.button_overlay_customize_group:(button_overlay_customize_group_list, ''),

                                  }
        self.connect_button_groups()
        '''Drag tables set'''
        table_style="""QTableWidget::item {color: black;height: 20px;} QTableWidget::item:selected {background-color: lightblue;color: black;}"""
        drag_table_list = {"Table1": ('table_combine_1_filename', partial(self.dropEvent_all, type="Combine")),
                           # "Table2": ('table_sketch_rescale_1_filename', partial(self.dropEvent_all, type="Sketch O")),
                           "Table3": ('table_overlay_service1', self.dropEvent_overlay_service1),
                           "Table4": ('table_overlay_mech', self.dropEvent_overlay_mech),
                           # "Sketch.1":("textedit_drawing_plot_filefrom_2", partial(self.drop_event_single,
                           #                                                   drag_in_widget=self.ui.textedit_drawing_plot_filefrom_2))
                           # "Table5": ('tableWidget_a', partial(self.dropEvent_all, type="Sketch A")),
                           # "Table6": ('tableWidget_b', partial(self.dropEvent_all, type="Sketch B")),
                           # "Table7": ('tableWidget_c', partial(self.dropEvent_all, type="Sketch C")),
                           # "Table8": ('tableWidget_d', partial(self.dropEvent_all, type="Sketch D")),
                           }
        for table_name in drag_table_list.keys():
            table_ui_name = drag_table_list[table_name][0]
            if table_name == "Table1":
                table = self.tableWidget1 = self.ui.findChild(QTableWidget, table_ui_name)
            elif table_name == "Sketch.1":
                table = self.tableWidget_combine_1_filename_2 = self.ui.findChild(QTableWidget, table_ui_name)
            # elif table_name == "Table2":
            #     table = self.tableWidget2 = self.ui.findChild(QTableWidget, table_ui_name)
            elif table_name == "Table3":
                table = self.tableWidget_overlay_service1 = self.ui.findChild(QTableWidget, table_ui_name)
            elif table_name == "Table4":
                table = self.tableWidget_overlay_mech = self.ui.findChild(QTableWidget, table_ui_name)
            # elif table_name == "Table5":
            #     table = self.tableWidget_a = self.ui.findChild(QTableWidget, table_ui_name)
            # elif table_name == "Table6":
            #     table = self.tableWidget_b = self.ui.findChild(QTableWidget, table_ui_name)
            # elif table_name == "Table7":
            #     table = self.tableWidget_c = self.ui.findChild(QTableWidget, table_ui_name)
            # elif table_name == "Table8":
            #     table = self.tableWidget_d = self.ui.findChild(QTableWidget, table_ui_name)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setAcceptDrops(True)
            table.dragMoveEvent = table.dragEnterEvent = self.dragEnterEvent
            table.dropEvent = drag_table_list[table_name][1]
            table.setStyleSheet(table_style)

        self.tableWidget_drawing = self.ui.findChild(QTableWidget, 'table_drawing')
        self.tableWidget_drawing.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.tableWidget3 = self.ui.findChild(QTableWidget, 'table_search_1')
        self.tableWidget3.setSortingEnabled(True)
        header = self.ui.table_search_1.horizontalHeader()
        header.setFixedHeight(40)
        self.tableWidget3.setSelectionBehavior(QTableWidget.SelectRows)
        self.tableWidget3.cellClicked.connect(self.selection_changed)
        self.tableWidget3.cellDoubleClicked.connect(self.f_tablesearch_doubleclick)
        self.tableWidget3.setStyleSheet(table_style)


        self.tableWidget4 = self.ui.findChild(QTableWidget, 'table_filling')
        self.tableWidget4.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget4.setStyleSheet(table_style)
        '''Text change trace'''
        self.text_to_function_copilot = {
            # self.ui.text_sketch_align_filedir: self.pagenum_cal,
            # self.ui.text_sketch_align_filename: self.text_sketch_align_filename_change,
            # self.ui.text_sketch_rescale_2_oscx:self.outscale_click,
            self.ui.text_overlay_service1_outs:self.f_text_overlay_service1_outs_change,
            # self.ui.text_sketch_rescale_2_outputcussize1:self.outsize_click,
            # self.ui.text_sketch_rescale_2_outputcussize2:self.outsize_click,
            self.ui.text_overlay_service_outsize1:self.f_overlay_service1_outsize,
            self.ui.text_overlay_service_outsize2:self.f_overlay_service1_outsize,
            self.ui.text_filing_name:self.f_change_filing_tablecolor,
            self.ui.text_search_1:self.f_all_search2,
            # self.ui.text_align_outfile:self.f_text_align_outfile,
            self.ui.text_overlay_servicetype:self.f_text_overlay_servicetype,
            self.ui.text_filing_folder:self.file_check_filing,
            self.ui.text_overlay_outfile:self.f_text_overlay_outfile_change,
            self.ui.text_issue_type:self.f_text_issue_type,
            self.ui.textedit_drawing_plot_filefrom:self.f_text_drawing_plot_filefrom,
            # self.ui.textedit_drawing_plot_filefrom_2: self.f_text_drawing_plot_filefrom,
            self.ui.text_drawing_plot_folderto:self.f_text_drawing_plot_folderto,
            # self.ui.text_color_tagnow:self.drawing_text_change_1,
            self.ui.text_drawing_sktechdir:self.drawing_text_change_5,
            # self.ui.text_sketch_rescale_1_oscx:partial(self.original_scale_text_change, name="Sketch O" ),
            # self.ui.text_rescale_a:partial(self.original_scale_text_change, name="Sketch A"),
            # self.ui.text_rescale_b:partial(self.original_scale_text_change, name="Sketch B"),
            # self.ui.text_rescale_c:partial(self.original_scale_text_change, name="Sketch C"),
            # self.ui.text_rescale_d:partial(self.original_scale_text_change, name="Sketch D"),
            self.ui.text_overlay_service1_os:partial(self.original_scale_text_change, name="Overlay"),
            # self.ui.text_sketch_rescale_7_outputcussize1:partial(self.original_size_text_change, name="Sketch O"),
            # self.ui.text_sketch_rescale_8_outputcussize2: partial(self.original_size_text_change, name="Sketch O"),
            # self.ui.text_resize1_a: partial(self.original_size_text_change, name="Sketch A"),
            # self.ui.text_resize2_a: partial(self.original_size_text_change, name="Sketch A"),
            # self.ui.text_resize1_b: partial(self.original_size_text_change, name="Sketch B"),
            # self.ui.text_resize2_b: partial(self.original_size_text_change, name="Sketch B"),
            # self.ui.text_resize1_c: partial(self.original_size_text_change, name="Sketch C"),
            # self.ui.text_resize2_c: partial(self.original_size_text_change, name="Sketch C"),
            # self.ui.text_resize1_d: partial(self.original_size_text_change, name="Sketch D"),
            # self.ui.text_resize2_d: partial(self.original_size_text_change, name="Sketch D"),
            self.ui.text_overlay_service_orisize1: partial(self.original_size_text_change, name="Overlay"),
            self.ui.text_overlay_service_orisize2: partial(self.original_size_text_change, name="Overlay"),
            # self.ui.text_sketch_colorprocess_1: partial(self.tags_text_change,name="Tag 1"),
            # self.ui.text_sketch_colorprocess_2: partial(self.tags_text_change,name="Tag 2"),
            # self.ui.text_sketch_colorprocess_3: partial(self.tags_text_change, name="Tag 3"),
            # self.ui.text_sketch_colorprocess_4: partial(self.tags_text_change, name="Tag 4"),
            # self.ui.text_sketch_colorprocess_4_1: partial(self.tags_text_change, name="Tag 4_1"),
            # self.ui.text_sketch_colorprocess_4_2: partial(self.tags_text_change, name="Tag 4_2"),
            # self.ui.text_sketch_colorprocess_5: partial(self.tags_text_change, name="Tag 5"),
            # self.ui.text_sketch_colorprocess_6: partial(self.tags_text_change, name="Tag 6"),
            # self.ui.text_sketch_colorprocess_7: partial(self.tags_text_change, name="Tag 7"),
            # self.ui.text_sketch_colorprocess_8: partial(self.tags_text_change, name="Tag 8"),
            # self.ui.text_sketch_colorprocess_8_1: partial(self.tags_text_change, name="Tag 8_1"),
            # self.ui.text_sketch_colorprocess_8_2: partial(self.tags_text_change, name="Tag 8_2"),
            # self.ui.text_sketch_colorprocess_8_3: partial(self.tags_text_change, name="Tag 8_3"),
            # self.ui.text_sketch_colorprocess_8_4: partial(self.tags_text_change, name="Tag 8_4"),
            # self.ui.text_sketch_colorprocess_8_5: partial(self.tags_text_change, name="Tag 8_5"),
            # self.ui.text_sketch_colorprocess_8_6: partial(self.tags_text_change, name="Tag 8_6"),
            # self.ui.text_color_pagenum: partial(self.drawing_text_change_all, name="Page num"),
            # self.ui.text_color_scale: partial(self.drawing_text_change_all, name="Scale"),
            # self.ui.text_color_papersize: partial(self.drawing_text_change_all, name="Paper size"),
            self.ui.text_all_1_searchbar:self.f_text_all_1_searchbar,
            self.ui.text_issue_folder:self.f_text_issue_folder,
            }
        self.text_to_function_copilot_2={}
        # for i in range(5):
        #     est_other_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_est_' + str(i + 1))
        #     self.text_to_function_copilot_2[est_other_text]=partial(self.f_calculate_sum_time, day=i + 1)
        #     act_other_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_act_' + str(i + 1))
        #     self.text_to_function_copilot_2[act_other_text]=partial(self.f_calculate_sum_time, day=i + 1)
        #     est_pro_text = self.ui.findChild(QLineEdit, 'text_timesheet_project_est_' + str(i + 1))
        #     self.text_to_function_copilot_2[est_pro_text] = partial(self.f_calculate_sum_time, day=i + 1)
        #     act_pro_text = self.ui.findChild(QLineEdit, 'text_timesheet_project_act_' + str(i + 1))
        #     self.text_to_function_copilot_2[act_pro_text] = partial(self.f_calculate_sum_time, day=i + 1)
        self.connect_texts()


        '''Frame original set (toolbutton name:(set fuction, frame name, show or not))'''
        self.toolbutton_to_function_copilot={
             # self.ui.toolbutton_color_1: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_color1,toolbutton_name=self.ui.toolbutton_color_1, true_text='-',false_text="+ Expand for more levels"), self.ui.frame_color1,True),
             # self.ui.toolbutton_color_2: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_color2,toolbutton_name=self.ui.toolbutton_color_2, true_text="-",false_text="+ Expand for more levels"), self.ui.frame_color2,True),
             self.ui.toolbutton_drawing_color_1: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_drawing_color1,toolbutton_name=self.ui.toolbutton_drawing_color_1, true_text="-",false_text="+ Expand for more levels"), self.ui.frame_drawing_color1,True),
             self.ui.toolbutton_drawing_color_2: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_drawing_color2,toolbutton_name=self.ui.toolbutton_drawing_color_2, true_text="-",false_text="+ Expand for more levels"), self.ui.frame_drawing_color2,True),
             # self.ui.toolbuton_modify_color: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_modify_color,toolbutton_name=self.ui.toolbuton_modify_color, true_text="Modify Color-",false_text="Modify Color+"), self.ui.frame_modify_color,True),
             self.ui.toolbutton_copymarkup: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_copymarkup,toolbutton_name=self.ui.toolbutton_copymarkup, true_text="Copy markup -",false_text="Copy markup +"), self.ui.frame_copymarkup,False),
             # self.ui.toolbutton_rescale1: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_rescale1,toolbutton_name=self.ui.toolbutton_rescale1, true_text="Block A -",false_text="Block A +"), self.ui.frame_rescale1,False),
             # self.ui.toolbutton_rescale2: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_rescale2,toolbutton_name=self.ui.toolbutton_rescale2, true_text="Block B -",false_text="Block B +"), self.ui.frame_rescale2,False),
             # self.ui.toolbutton_rescale3: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_rescale3,toolbutton_name=self.ui.toolbutton_rescale3, true_text="Block C -",false_text="Block C +"), self.ui.frame_rescale3,False),
             # self.ui.toolbutton_rescale4: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_rescale4,toolbutton_name=self.ui.toolbutton_rescale4, true_text="Block D -",false_text="Block D +"), self.ui.frame_rescale4,False),
             # self.ui.toolbutton_align_a: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_align_a,toolbutton_name= self.ui.toolbutton_align_a, true_text="Block A -",false_text="Block A +"), self.ui.frame_align_a,False),
             # self.ui.toolbutton_align_b: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_align_b,toolbutton_name=self.ui.toolbutton_align_b, true_text="Block B -",false_text="Block B +"), self.ui.frame_align_b,False),
             # self.ui.toolbutton_align_c: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_align_c,toolbutton_name=self.ui.toolbutton_align_c, true_text="Block C -",false_text="Block C +"), self.ui.frame_align_c,False),
             # self.ui.toolbutton_align_d: (partial(self.toggleVisibility_all, frame_name=self.ui.frame_align_d,toolbutton_name=self.ui.toolbutton_align_d, true_text="Block D -",false_text="Block D +"), self.ui.frame_align_d,False),
        }
        self.connect_toolbuttons()
        # self.ui.frame_modify_color.setVisible(False)

        '''Combobox change set'''
        # self.ui.combobox_color_sketch_o.currentIndexChanged.connect(self.f_combobox_color_1)
        self.ui.combox_drawing_update_des_a1.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=0))
        self.ui.combox_drawing_update_des_a2.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=1))
        self.ui.combox_drawing_update_des_a3.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=2))
        self.ui.combox_drawing_update_des_a4.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=3))
        self.ui.combox_drawing_update_des_a5.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=4))
        self.ui.combox_drawing_update_des_a6.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=5))
        self.ui.combox_drawing_update_des_a7.currentIndexChanged.connect(partial(self.f_combox_drawing_update_des_all, rev_num=6))
        '''Tab change set'''
        # self.tabs.currentChanged.connect(self.onTabChanged)
        '''Drawing frame content list'''
        self.rev_list = [self.ui.text_drawing_update_rev1, self.ui.text_drawing_update_rev2, self.ui.text_drawing_update_rev3,
                    self.ui.text_drawing_update_rev4, self.ui.text_drawing_update_rev5, self.ui.text_drawing_update_rev6,
                    self.ui.text_drawing_update_rev7]
        self.description_a_list = [self.ui.combox_drawing_update_des_a1, self.ui.combox_drawing_update_des_a2,
                              self.ui.combox_drawing_update_des_a3,
                              self.ui.combox_drawing_update_des_a4, self.ui.combox_drawing_update_des_a5,
                              self.ui.combox_drawing_update_des_a6,
                              self.ui.combox_drawing_update_des_a7]
        self.description_b_list = [self.ui.combox_drawing_update_des_b1, self.ui.combox_drawing_update_des_b2,
                              self.ui.combox_drawing_update_des_b3,
                              self.ui.combox_drawing_update_des_b4, self.ui.combox_drawing_update_des_b5,
                              self.ui.combox_drawing_update_des_b6,
                              self.ui.combox_drawing_update_des_b7]
        self.drw_list = [self.ui.combox_drawing_update_drw1, self.ui.combox_drawing_update_drw2,
                    self.ui.combox_drawing_update_drw3,
                    self.ui.combox_drawing_update_drw4, self.ui.combox_drawing_update_drw5,
                    self.ui.combox_drawing_update_drw6,
                    self.ui.combox_drawing_update_drw7]
        self.chk_list = [self.ui.combox_drawing_update_chk1, self.ui.combox_drawing_update_chk2,
                    self.ui.combox_drawing_update_chk3,
                    self.ui.combox_drawing_update_chk4, self.ui.combox_drawing_update_chk5,
                    self.ui.combox_drawing_update_chk6,
                    self.ui.combox_drawing_update_chk7]
        self.date_list = [self.ui.combox_drawing_update_date1, self.ui.combox_drawing_update_date2,
                     self.ui.combox_drawing_update_date3,
                     self.ui.combox_drawing_update_date4, self.ui.combox_drawing_update_date5,
                     self.ui.combox_drawing_update_date6,
                     self.ui.combox_drawing_update_date7]
        '''Table information dict'''
        self.overlay_service1_parentdir=''
        self.table_dict = {"Table_combine": (self.tableWidget1, self.ui.text_combine_1_foldername),
                           # "Table_rescale_O": (self.tableWidget2, self.ui.text_sketch_rescale_1_filedir),
                           # "Table_rescale_A": (self.tableWidget_a, self.ui.text_rescale_cf_1),
                           # "Table_rescale_B": (self.tableWidget_b, self.ui.text_rescale_cf_2),
                           # "Table_rescale_C": (self.tableWidget_c, self.ui.text_rescale_cf_3),
                           # "Table_rescale_D": (self.tableWidget_d, self.ui.text_rescale_cf_4),
                           "Table_overlay_service": (self.tableWidget_overlay_service1, self.overlay_service1_parentdir),
                           "Table_overlay_mech": (self.tableWidget_overlay_mech, self.ui.text_overlay_mech_foldername),}
        '''Folder dir change dict'''
        self.folder_text_dict={"Overlay output":self.ui.text_overlay_folder1,
                               "Combine output":self.ui.text_combine_op,
                               "Overlay mech":self.ui.text_overlay_mech_outfolder,
                               "Filing":self.ui.text_filing_folder,
                               # "Rescale O":self.ui.text_sketch_rescale_1_filesavedir,
                               # "Rescale A":self.ui.text_rescale_of_1,
                               # "Rescale B":self.ui.text_rescale_of_2,
                               # "Rescale C":self.ui.text_rescale_of_3,
                               # "Rescale D":self.ui.text_rescale_of_4,
                               # "Sketch Align output":self.ui.text_align_outfile,
                               # "Align output":self.ui.text_align_outfile,
                               "Issue":self.ui.text_issue_folder,
                               "Plot":self.ui.text_drawing_plot_folderto}
        '''File dir change dict'''
        self.file_dir_dict = {
            # "Rescale O": (self.ui.text_sketch_align_filedir, self.ui.text_sketch_align_filename),
                              # "Rescale A": (self.ui.text_folder_a, self.ui.text_file_a),
                              # "Rescale B": (self.ui.text_folder_b, self.ui.text_file_b),
                              # "Rescale C": (self.ui.text_folder_c, self.ui.text_file_c),
                              # "Rescale D": (self.ui.text_folder_d, self.ui.text_file_d),
        }
        '''Original scale change dict'''
        self.original_scale_change_dict = {
            # "Sketch O": (self.ui.text_sketch_rescale_1_oscx, self.button_group1),
                                           # "Sketch A": (self.ui.text_rescale_a, self.button_group_a1),
                                           # "Sketch B": (self.ui.text_rescale_b, self.button_group_b1),
                                           # "Sketch C": (self.ui.text_rescale_c, self.button_group_c1),
                                           # "Sketch D": (self.ui.text_rescale_d, self.button_group_d1),
                                           # "Overlay": (self.ui.text_overlay_service1_os, self.button_group_overlay_service1_os),
        }
        '''Original size change dict'''
        self.original_size_change_dict = {
            # "Sketch O": (self.ui.text_sketch_rescale_7_outputcussize1,self.ui.text_sketch_rescale_8_outputcussize2,self.button_group2),
                                           # "Sketch A": (self.ui.text_resize1_a.text(),self.ui.text_resize2_a.text(),self.button_group_a2),
                                           # "Sketch B": (self.ui.text_resize1_b.text(),self.ui.text_resize2_b.text(),self.button_group_b2),
                                           # "Sketch C": (self.ui.text_resize1_c.text(),self.ui.text_resize2_c.text(),self.button_group_c2),
                                           # "Sketch D": (self.ui.text_resize1_d.text(),self.ui.text_resize2_d.text(),self.button_group_d2),
                                           # "Overlay": (self.ui.text_overlay_service_orisize1.text(),self.ui.text_overlay_service_orisize2.text(),self.button_group_overlay_service1_orisize),
        }
        '''Tag text change dict'''
        # self.tag_text_dict={"Tag 1":(self.ui.text_sketch_colorprocess_1,self.ui.box_tag_a_1),
        #                     "Tag 2":(self.ui.text_sketch_colorprocess_2,self.ui.box_tag_a_2),
        #                     "Tag 3": (self.ui.text_sketch_colorprocess_3,self.ui.box_tag_a_3),
        #                     "Tag 4": (self.ui.text_sketch_colorprocess_4,self.ui.box_tag_a_4),
        #                     "Tag 4_1": (self.ui.text_sketch_colorprocess_4_1,self.ui.box_tag_a_4_1),
        #                     "Tag 4_2": (self.ui.text_sketch_colorprocess_4_2,self.ui.box_tag_a_4_2),
        #                     "Tag 5": (self.ui.text_sketch_colorprocess_5,self.ui.box_tag_a_5),
        #                     "Tag 6": (self.ui.text_sketch_colorprocess_6,self.ui.box_tag_a_6),
        #                     "Tag 7": (self.ui.text_sketch_colorprocess_7,self.ui.box_tag_a_7),
        #                     "Tag 8": (self.ui.text_sketch_colorprocess_8,self.ui.box_tag_a_8),
        #                     "Tag 8_1": (self.ui.text_sketch_colorprocess_8_1,self.ui.box_tag_a_8_1),
        #                     "Tag 8_2": (self.ui.text_sketch_colorprocess_8_2,self.ui.box_tag_a_8_2),
        #                     "Tag 8_3": (self.ui.text_sketch_colorprocess_8_3,self.ui.box_tag_a_8_3),
        #                     "Tag 8_4": (self.ui.text_sketch_colorprocess_8_4,self.ui.box_tag_a_8_4),
        #                     "Tag 8_5": (self.ui.text_sketch_colorprocess_8_5,self.ui.box_tag_a_8_5),
        #                     "Tag 8_6": (self.ui.text_sketch_colorprocess_8_6,self.ui.box_tag_a_8_6),}
        '''Drawing info dict'''
        # self.drawing_text_change_dict={
            # "Page num":(self.ui.text_color_pagenum,self.ui.text_drawing_2),
                                       # "Scale":(self.ui.text_color_scale,self.ui.text_drawing_3),
                                       # "Paper size": (self.ui.text_color_papersize,self.ui.text_drawing_4)
                                       # }
        '''Customized size text change'''
        self.buttongroup_size_click_dict = {
            "Overlay origin": (self.button_group_overlay_service1_orisize, self.ui.text_overlay_service_orisize1,self.ui.text_overlay_service_orisize2),
                                            "Overlay output": (self.button_group_overlay_service1_outsize,self.ui.text_overlay_service_outsize1,self.ui.text_overlay_service_outsize2),
                                            # "Sketch O origin": (self.button_group2, self.ui.text_sketch_rescale_7_outputcussize1,self.ui.text_sketch_rescale_8_outputcussize2),
                                            # "Sketch A origin": (self.button_group_a2, self.ui.text_resize1_a, self.ui.text_resize2_a),
                                            # "Sketch B origin": (self.button_group_b2, self.ui.text_resize1_b, self.ui.text_resize2_b),
                                            # "Sketch C origin": (self.button_group_c2, self.ui.text_resize1_c, self.ui.text_resize2_c),
                                            # "Sketch D origin": (self.button_group_d2, self.ui.text_resize1_d, self.ui.text_resize2_d),
            }
        '''Table dropevent dict'''
        self.dropevent_dict = {#"Sketch O": (self.tableWidget2, self.ui.text_sketch_rescale_1_filedir, self.ui.text_sketch_rescale_1_filesavedir,self.ui.text_sketch_rescale_2_filename),
                               # "Sketch A": (self.tableWidget_a, self.ui.text_rescale_cf_1, self.ui.text_rescale_of_1,self.ui.text_rescale_outfile_1),
                               # "Sketch B": (self.tableWidget_b, self.ui.text_rescale_cf_2, self.ui.text_rescale_of_2,self.ui.text_rescale_outfile_2),
                               # "Sketch C": (self.tableWidget_c, self.ui.text_rescale_cf_3, self.ui.text_rescale_of_3,self.ui.text_rescale_outfile_3),
                               # "Sketch D": (self.tableWidget_d, self.ui.text_rescale_cf_4, self.ui.text_rescale_of_4,self.ui.text_rescale_outfile_4),
                               "Combine": (self.tableWidget1, self.ui.text_combine_1_foldername, self.ui.text_combine_op,self.ui.text_combine_3_filenameadd), }
        '''Customized scale text change'''
        self.buttongroup_scale_dict = {
            "Overlay output": (self.button_group_overlay_service1_outs, 3, self.ui.text_overlay_service1_outs),
            "Overlay origin": (self.button_group_overlay_service1_os, 6, self.ui.text_overlay_service1_os),
            # "Sketch O": (self.button_group1, 6, self.ui.text_sketch_rescale_1_oscx),
            # "Sketch A": (self.button_group_a1, 6, self.ui.text_rescale_a),
            # "Sketch B": (self.button_group_b1, 6, self.ui.text_rescale_b),
            # "Sketch C": (self.button_group_c1, 6, self.ui.text_rescale_c),
            # "Sketch D": (self.button_group_d1, 6, self.ui.text_rescale_d),
            # "Sketch output": (self.button_group3, 3, self.ui.text_sketch_rescale_2_oscx),
            }
        '''Get table file dict'''
        self.gettablefile_dict = {"Combine": self.ui.table_combine_1_filename,
                                  # "Sketch.1": self.ui.table_combine_1_filename_2,
                                  "Overlay mech": self.tableWidget_overlay_mech,
                                  "Overlay service": self.ui.table_overlay_service1,
                                  # "Sketch O": self.ui.table_sketch_rescale_1_filename,
                                  # "Sketch A": self.ui.tableWidget_a,
                                  # "Sketch B": self.ui.tableWidget_b,
                                  # "Sketch C": self.ui.tableWidget_c,
                                  # "Sketch D": self.ui.tableWidget_d,
                                  }
        '''Enter search connect'''
        self.ui.text_all_1_searchbar.returnPressed.connect(self.f_all_search)
        '''Table size set'''
        self.ui.table_drawing_linkcontent.setColumnWidth(0, 1000)
        self.ui.table_drawing_foldercontent.setColumnWidth(0, 1000)
        table=self.ui.table_timesheet_user
        table.setColumnWidth(1, 100)
        table.setColumnWidth(2, 150)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        width_money=80
        column_len_list=([80, 60, 290]+# 0-2
                         [80, 90, 45, 80, 100, 60]+# 3-8
                         [85, 85, 85, 85, 85, 85]+# 9-14
                         [100, 120, 320]+# 15-17
                         [40, 100, 90, 190, 100, 90, 190]+# 18-24
                         [55, 55, width_money, width_money, width_money, width_money, width_money, width_money]+# 25-32
                         [0 for _ in range(19)]+# 33-51
                         [width_money for _ in range(9)]+# 52-60
                         [60 for _ in range(8)]+# 61-68
                         [0 for _ in range(10)]+# 69-78
                         [width_money,width_money]+#79-80
                         [width_money for _ in range(12)]+#81-92
                         [0 for _ in range(10)]+# 93-102
                         [50 for _ in range(6)] +  # 103-108
                         [0,0]+ # 109-110
                         [50 for _ in range(6)]+# 111-116
                         [0 for _ in range(40)]# 117-156
                         )
        for i in range(len(column_len_list)):
            self.tableWidget3.setColumnWidth(i, column_len_list[i])
        table=self.ui.table_scope_all
        table.setColumnWidth(0, 200)
        table.setColumnWidth(1, 200)
        table.horizontalHeader().setStretchLastSection(True)
        self.f_set_table_style(table)
        table=self.ui.table_plot_link_and_date
        table.setColumnWidth(1, 70)
        table.setColumnHidden(2, True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

        self.ui.table_plot_link_and_date.itemClicked.connect(self.f_table_plot_link_and_date)

        self.ui.frame_logo.hide()

        '''Bridge frame hide'''
        frame_bridge_top_1 = self.ui.findChild(QFrame, 'frame_pro_stage')
        frame_bridge_top_2 = self.ui.findChild(QFrame, 'frame_pro_func')
        frame_bridge_top_3 = self.ui.findChild(QFrame, 'frame_pro_invoice')
        frame_bridge_top_1.hide()
        frame_bridge_top_2.hide()
        frame_bridge_top_3.hide()

        self.ui.frame_login.hide()
        self.ui.tabWidget.setTabVisible(10, False)
        self.ui.tabWidget.setTabVisible(11, False)
        self.ui.tabWidget.setTabVisible(12, False)

        self.f_button_search_hide(region=2)
        self.f_button_search_hide(region=3)
        self.f_button_search_hide(region=4)
        self.f_button_search_hide(region=5)
        self.f_button_search_hide(region=6)

        button_index_list=[2,3,4,5,6]
        for index in button_index_list:
            button=self.ui.findChild(QPushButton, 'button_search_hide_'+str(index))
            button.hide()

        # self.ui.toolbutton_rescale2.hide()
        # self.ui.toolbutton_rescale3.hide()
        # self.ui.toolbutton_rescale4.hide()


        self.ui.combobox_search_history.currentIndexChanged.connect(self.f_choose_from_search_history)

        self.ui.tabWidget.setTabVisible(7, False)

        self.ui.text_all_3_projectname.setCursorPosition(0)

        self.ui.frame_link_files.hide()


        '''==================================================Timesheet============================================='''
        self.Frame_Mon= self.ui.findChild(QFrame, "Frame_Mon")
        self.layout_Frame_Mon = self.Frame_Mon.layout()
        self.Frame_Tue= self.ui.findChild(QFrame, "Frame_Tue")
        self.layout_Frame_Tue = self.Frame_Tue.layout()
        self.Frame_Wed= self.ui.findChild(QFrame, "Frame_Wed")
        self.layout_Frame_Wed = self.Frame_Wed.layout()
        self.Frame_Thu= self.ui.findChild(QFrame, "Frame_Thu")
        self.layout_Frame_Thu = self.Frame_Thu.layout()
        self.Frame_Fri= self.ui.findChild(QFrame, "Frame_Fri")
        self.layout_Frame_Fri = self.Frame_Fri.layout()



    def eventFilter(self, obj, event):
        if obj == self.line_edit and event.type() == QEvent.FocusIn:
            print(1111)
        return super().eventFilter(obj, event)

    '''================================Connection set=================================='''
    '''Button connection'''
    def connect_buttons(self):
        for button, function in self.button_to_function_copilot.items():
            button.clicked.connect(function)
    '''Text connection'''
    def connect_texts(self):
        for text, function in self.text_to_function_copilot.items():
            text.textChanged.connect(function)
        for text, function in self.text_to_function_copilot_2.items():
            text.textChanged.connect(function)
    '''Toolbutton connection'''
    def connect_toolbuttons(self):
        for toolbutton, function in self.toolbutton_to_function_copilot.items():
            function_to_connect=function[0]
            frame_name=function[1]
            show_ornot=function[2]
            toolbutton.clicked.connect(function_to_connect)
            frame_name.setVisible(show_ornot)
    '''Button group connection'''
    def connect_button_groups(self):
        for button_group in self.button_group_dict.keys():
            button_group_list=self.button_group_dict[button_group][0]
            function=self.button_group_dict[button_group][1]
            for i in range(len(button_group_list)):
                button_group.addButton(button_group_list[i], i+1)
            if function!='':
                button_group.buttonClicked.connect(function)

    '''Button disconnection'''
    def disconnect_buttons(self):
        for button, function in self.button_to_function_copilot.items():
            button.clicked.disconnect(function)
    '''Text disconnection'''
    def disconnect_texts(self):
        for text, function in self.text_to_function_copilot.items():
            text.textChanged.disconnect(function)
    '''Toolbutton disconnection'''
    def disconnect_toolbuttons(self):
        for toolbutton, function in self.toolbutton_to_function_copilot.items():
            function_to_connect=function[0]
            frame_name=function[1]
            show_ornot=function[2]
            toolbutton.clicked.disconnect(function_to_connect)
            frame_name.setVisible(show_ornot)
    '''Button group disconnection'''
    def disconnect_button_groups(self):
        for button_group in self.button_group_dict.keys():
            button_group_list=self.button_group_dict[button_group][0]
            function=self.button_group_dict[button_group][1]
            for i in range(len(button_group_list)):
                button_group.addButton(button_group_list[i], i+1)
            if function!='':
                button_group.buttonClicked.disconnect(function)
    '''Disconnect all'''
    def disconnect_all(self):
        try:
            self.disconnect_buttons()
        except:
            pass
        try:
            self.disconnect_texts()
        except:
            pass
        try:
            self.disconnect_toolbuttons()
        except:
            pass
        try:
            self.disconnect_button_groups()
        except:
            pass
    '''Connect all'''
    def connect_all(self):
        try:
            self.connect_buttons()
        except:
            pass
        try:
            self.connect_texts()
        except:
            pass
        try:
            self.connect_toolbuttons()
        except:
            pass
        try:
            self.connect_button_groups()
        except:
            pass
    '''Frame show or not function'''
    def toggleVisibility_all(self,frame_name,toolbutton_name,true_text,false_text):
        try:
            frame_name.setVisible(not frame_name.isVisible())
            if frame_name.isVisible()==True:
                toolbutton_name.setText(true_text)
            else:
                toolbutton_name.setText(false_text)
        except:
            traceback.print_exc()
    '''Tab change id'''
    def switch_tab(self, page_name):
        try:
            page_name_dict={'Search':9}
            index=page_name_dict[page_name]
            # self.tabs = self.ui.tabWidget
            # self.tabs.setCurrentIndex(index)
        except:
            traceback.print_exc()
    '''Tab change'''
    def onTabChanged(self):
        # try:
        copilot_tab_page=[0,1,2,3,4,5,6,7,8,9]
        bridge_tab_page=[10,11,12]
        current_tab=self.tabs.currentIndex()
        '''Search bar'''
        frame_bridge_top_1 = self.ui.findChild(QFrame, 'frame_pro_stage')
        frame_bridge_top_2 = self.ui.findChild(QFrame, 'frame_pro_func')
        frame_bridge_top_3 = self.ui.findChild(QFrame, 'frame_pro_invoice')
        frame_copilot_top_1 = self.ui.findChild(QFrame, 'qframe_button')
        frame_stretch=self.ui.findChild(QFrame, 'frame_stretch')
        frame_stretch.hide()
        if self.last_tab in bridge_tab_page and current_tab in copilot_tab_page:
            frame_bridge_top_1.hide()
            frame_bridge_top_2.hide()
            frame_bridge_top_3.hide()
            frame_copilot_top_1.show()
        elif self.last_tab in copilot_tab_page and current_tab in bridge_tab_page:
            frame_copilot_top_1.hide()
            frame_bridge_top_1.show()
            frame_bridge_top_2.show()
            frame_bridge_top_3.show()
            status_all=1
            update_time_all='2024-01-01 00:00'
            # try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM bridge.update_exe_status")
            exe_update_database = mycursor.fetchall()
            for exe_update_history in exe_update_database:
                update_time=exe_update_history[1]
                status=exe_update_history[2]
                if status==0:
                    status_all=0
                time1=datetime.strptime(update_time_all, '%Y-%m-%d %H:%M')
                time2=datetime.strptime(update_time, '%Y-%m-%d %H:%M')
                if time2>time1:
                    update_time_all=update_time
            self.ui.findChild(QLineEdit, 'text_database_update_time').setText(update_time_all)
                # except:
                #     traceback.print_exc()
                # finally:
                #     mycursor.close()
                #     mydb.close()
            quo_num_copilot=self.ui.text_all_1_projectnumber.text()
            text_quo_bridge=self.ui.findChild(QLineEdit, 'text_proinfo_1_quonum').text()
            if quo_num_copilot!=text_quo_bridge:
                stats2.f_button_pro_search(quo_num_copilot)
            if current_tab == 9:
                frame_stretch.show()
            '''Change to Drawing'''
            if current_tab == 4:
                '''Set tags and table'''
                tags = []
                # for i in range(len(self.button_group_tag_boxlist)):
                #     if self.button_group_tag_boxlist[i].isChecked():
                #         self.button_group_tag2_boxlist[i].setChecked(True)
                #         tag=self.button_group_tag_textlist[i].text()
                #         self.button_group_tag2_textlist[i].setText(tag)
                #         tags.append(tag)
                #     else:
                #         self.button_group_tag2_boxlist[i].setChecked(False)
                self.set_draw_table(tags)
                tag_fold_sketch1=False
                tag_fold_sketch2 = False
                for i in range(6):
                    if self.button_group_tag2.button(i + 1).isChecked():
                        tag_fold_sketch1 =True
                        break
                for i in range(10):
                    if self.button_group_tag2.button(i + 18).isChecked():
                        tag_fold_sketch2 = True
                        break
                if tag_fold_sketch1:
                    self.ui.frame_drawing_color1.setVisible(True)
                    self.ui.toolbutton_drawing_color_1.setText("-")
                else:
                    self.ui.frame_drawing_color1.setVisible(False)
                    self.ui.toolbutton_drawing_color_1.setText("+ Expand for more levels")
                if tag_fold_sketch2:
                    self.ui.frame_drawing_color2.setVisible(True)
                    self.ui.toolbutton_drawing_color_2.setText("-")
                else:
                    self.ui.frame_drawing_color2.setVisible(False)
                    self.ui.toolbutton_drawing_color_2.setText("+ Expand for more levels")
                '''Set drawing name'''
                if self.ui.text_drawing_sktechdir.text()!='':
                    drawing_name=os.path.join(Path(self.ui.text_drawing_sktechdir.text()).parent.absolute(),self.ui.text_all_3_projectname.text()+'-Mechanical Drawings.pdf')
                    self.ui.text_drawing_output.setText(drawing_name)
            '''Change to Sketch - Tag frame show or not'''
            # if current_tab == 3:
            #     tag_fold_sketch1=False
            #     tag_fold_sketch2 = False
                # for i in range(6):
                #     if self.button_group_tag.button(i + 1).isChecked():
                #         tag_fold_sketch1 =True
                #         break
                # for i in range(10):
                #     if self.button_group_tag.button(i + 18).isChecked():
                #         tag_fold_sketch2 = True
                #         break
                # if tag_fold_sketch1:
                #     self.ui.frame_color1.setVisible(True)
                #     self.ui.toolbutton_color_1.setText("-")
                # else:
                #     self.ui.frame_color1.setVisible(False)
                #     self.ui.toolbutton_color_1.setText("+ Expand for more levels")
                # if tag_fold_sketch2:
                #     self.ui.frame_color2.setVisible(True)
                #     self.ui.toolbutton_color_2.setText("-")
                # else:
                #     self.ui.frame_color2.setVisible(False)
                #     self.ui.toolbutton_color_2.setText("+ Expand for more levels")

            self.last_tab=current_tab
        # except:
        #     traceback.print_exc()
    '''=================================Table function set==============================='''
    '''Set color of table'''
    def f_set_table_style(self, table):
        try:
            for row in range(table.rowCount()):
                for column in range(table.columnCount()):
                    if table.item(row, column) == None:
                        item = QTableWidgetItem()
                    else:
                        item = table.item(row, column)
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
        except:
            print(table, row, column)
            traceback.print_exc()
    '''Open file in table'''
    def f_tablefile_open(self,table_name):
        try:
            tableWidget,file_dir=self.table_dict[table_name][0],self.table_dict[table_name][1].text()
            selected_items = tableWidget.selectedItems()
            if selected_items:
                content = selected_items[0].text()
            file_name = os.path.join(file_dir, content)
            open_in_bluebeam(file_name)
        except:
            traceback.print_exc()
    '''Up move file in table'''
    def f_tableitem_up(self,table_name):
        try:
            table=self.table_dict[table_name][0]
            selected_items = table.selectedItems()
            if selected_items:
                row = selected_items[0].row()
                if row!=0:
                    content=selected_items[0].text()
                    table.removeRow(row)
                    table.insertRow(row-1)
                    table.setItem(row-1, 0, QTableWidgetItem(content))
                    table.setCurrentCell(row-1,0)
        except:
            traceback.print_exc()
    '''Down move file in table'''
    def f_tableitem_down(self,table_name):
        try:
            table = self.table_dict[table_name][0]
            selected_items = table.selectedItems()
            if selected_items:
                row = selected_items[0].row()
                if int(row)!=int(table.rowCount())-1:
                    content=selected_items[0].text()
                    table.removeRow(row)
                    table.insertRow(row+1)
                    table.setItem(row+1, 0, QTableWidgetItem(content))
                    table.setCurrentCell(row+1,0)
        except:
            traceback.print_exc()
    '''Delete file in table'''
    def f_tableitem_delete(self,table_name):
        try:
            table = self.table_dict[table_name][0]
            selected_items = table.selectedItems()
            if selected_items:
                row = selected_items[0].row()
                table.removeRow(row)
        except:
            traceback.print_exc()
    '''Delete all files in table'''
    def f_tableitem_deleteall(self,table_name):
        try:
            table = self.table_dict[table_name][0]
            table.setRowCount(0)
        except:
            traceback.print_exc()
    '''DropEvent set'''
    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()
    '''Table drop event for sketch O~D and combine'''
    def drop_event_single(self, event, drag_in_widget):
        file_name = event.mimeData().text().rstrip('\n').split('\n')[0].lstrip("file:///")
        if drag_in_widget.rowCount()==0:
            drag_in_widget.insertRow(0)
        drag_in_widget.setItem(0, 0, QTableWidgetItem(file_name))


    def dropEvent_all(self, event,type):
        try:
            table=self.dropevent_dict[type][0]
            parent_folder_text=self.dropevent_dict[type][1]
            output_folder_text=self.dropevent_dict[type][2]
            output_file_text = self.dropevent_dict[type][3]
            for line in event.mimeData().text().rstrip('\n').split('\n'):
                dropped_text = line.lstrip("file:///")
                row_position = table.rowCount()
                if row_position==0:
                    if type=="Sketch O" or type=="Combine":
                        pro_name = dropped_text.split('/')[1]
                        loc_project_divide = pro_name.find('-')
                        if loc_project_divide != -1:
                            proj_num = pro_name[:loc_project_divide]
                        self.f_search_fileadded(proj_num)
                    parent_dir = str(Path(dropped_text).parent.absolute())
                    parent_folder_text.setText(parent_dir)
                    if type=="Combine":
                        output_folder_text.setText(parent_dir)
                        output_file_text.setText(str(date.today().strftime("%Y%m%d")) + '-Combined.pdf')
                    else:
                        output_folder_text.setText(self.backup_folder)
                        output_file_text.setText(str(self.ui.text_all_3_projectname.text()+'-Mechanical '+type+'.pdf'))
                file_name=str(Path(dropped_text).name)
                table.insertRow(row_position)
                table.setItem(row_position, 0, QTableWidgetItem(file_name))

            event.acceptProposedAction()
        except:
            traceback.print_exc()
    '''Table drop event for overlay mech'''
    def dropEvent_overlay_mech(self, event):
        try:
            for line in event.mimeData().text().rstrip('\n').split('\n'):
                dropped_text = line.lstrip("file:///")
                row_position = self.tableWidget_overlay_mech.rowCount()
                if row_position==0:
                    parent_dir = str(Path(dropped_text).parent.absolute())
                    file_directory=Path(dropped_text).parent
                    pro_name_overlay_file=list(file_directory.parts)[1]
                    loc_project_overlay_divide = pro_name_overlay_file.find('-')
                    if loc_project_overlay_divide != -1:
                        proj_overlay_num = pro_name_overlay_file[:loc_project_overlay_divide]
                        if proj_overlay_num!=self.ui.text_all_2_quotationnumber.text() and proj_overlay_num!=self.ui.text_all_1_projectnumber.text():
                            message('Wrong overlay folder, please choose file again.',self.ui)
                            return
                    self.ui.text_overlay_mech_foldername.setText(parent_dir)
                    self.ui.text_overlay_mech_outfolder.setText(os.path.join(self.ui.text_overlay_folder1.text(),self.ui.text_overlay_folder2.text()))
                    self.ui.text_overlay_mech_filenameadd.setText(str(date.today().strftime("%Y%m%d")) + '-Combined Mech.pdf')
                file_name=str(Path(dropped_text).name)
                self.tableWidget_overlay_mech.insertRow(row_position)
                self.tableWidget_overlay_mech.setItem(row_position, 0, QTableWidgetItem(file_name))

            combinefile_list =self.get_table_files("Overlay mech")
            current_pdf_parent_directory = str(self.ui.text_overlay_mech_foldername.text())
            page_num_sum=0
            paper_sizes=''
            for i in range(len(combinefile_list)):
                filename=os.path.join(current_pdf_parent_directory, combinefile_list[i])
                page_num_i=get_number_of_page(filename)
                page_num_sum+=page_num_i
                paper_size=get_paper_size(filename, 0)
                if paper_size!=paper_sizes:
                    if paper_sizes=='':
                        paper_sizes = paper_size
                    else:
                        paper_sizes=paper_sizes+'&'+paper_size
            self.ui.text_overlay_mech_pages.setText(str(page_num_sum))
            self.ui.text_overlay_mech_size.setText(str(paper_sizes))
            event.acceptProposedAction()
        except:
            traceback.print_exc()
    '''Table drop event for overlay service'''
    def dropEvent_overlay_service1(self,event):
        try:
            for line in event.mimeData().text().rstrip('\n').split('\n'):
                dropped_text = line.lstrip("file:///")
                row_position = self.tableWidget_overlay_service1.rowCount()
                if row_position==0:
                    self.overlay_service1_parentdir = str(Path(dropped_text).parent.absolute())
                    try:
                        folder_name = self.overlay_service1_parentdir
                        for i in range(10):
                            print(str(Path(folder_name).parent.absolute())[-8:])
                            if str(Path(folder_name).parent.absolute()).upper()[-8:] == 'EXTERNAL':
                                break
                            else:
                                folder_name = Path(folder_name).parent.absolute()
                        try:
                            service_name=self.service_name+'-'
                        except:
                            service_name=''
                        try:
                            layername=service_name+str(Path(folder_name).name).split('-')[0]
                        except:
                            layername=service_name+str(date.today().strftime("%Y%m%d"))
                        self.ui.text_overlay_layer.setText(layername)
                    except:
                        traceback.print_exc()
                file_name=str(Path(dropped_text).name)
                self.tableWidget_overlay_service1.insertRow(row_position)
                self.tableWidget_overlay_service1.setItem(row_position, 0, QTableWidgetItem(file_name))
            event.acceptProposedAction()
        except:
            traceback.print_exc()
    '''Get table files'''
    def get_table_files(self,tablename):
        try:
            table=self.gettablefile_dict[tablename]
            combinefile_list = []
            for row in range(table.rowCount()):
                item = table.item(row, 0)
                if item is not None:
                    combinefile_list.append(item.text())
            return combinefile_list
        except:
            traceback.print_exc()
    '''===============================Folder and file function set=========================='''
    '''Change and reset folder'''


    def open_function(self, input_widget, folder_path):
        if hasattr(self, "current_folder_address"):
            input_file_dir = QFileDialog.getOpenFileName(None, "Choose File", self.current_folder_address)[0]
            output_folder = os.path.join(folder_path, self.quotation_number)
            create_directory(output_folder)
        else:
            return
            input_file_dir = QFileDialog.getOpenFileName(None, "Choose File", ".")[0]
            # need to fix this part because right now in the bridge the current folder address is mess up
            current_folder = Path(input_file_dir).parts[1]
            result = get_value_from_table_with_filter("projects", "current_folder_address", current_folder)
            # if len(result)==0:

        if input_file_dir != '':
            # output_filename = get_timestamp()+self.project_name+".pdf"
            # output_file_dir = os.path.join(output_folder, output_filename)
            input_widget.setText(input_file_dir)


    def f_folder_change(self,folder_text_name):
        try:
            file_direc_read = QFileDialog.getExistingDirectory(None, "Choose file", self.current_folder_address)
        except:
            file_direc_read = QFileDialog.getExistingDirectory(None, "Choose file", ".")
        if file_direc_read != '':
            folder_text=self.folder_text_dict[folder_text_name]
            folder_text.setText(file_direc_read)
            if folder_text_name=='Issue':
                directory = Path(file_direc_read).parent
                pro_name = list(directory.parts)[1]
                loc_project_divide=pro_name.find('-')
                if loc_project_divide != -1:
                    proj_num = pro_name[:loc_project_divide]
                self.f_search_fileadded(proj_num)
                folder_text.setText(file_direc_read)

    '''Get filename'''
    def f_get_filename(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.ReadOnly
            try:
                fileName, _ = QFileDialog.getOpenFileName(self.ui, "QFileDialog.getOpenFileName()",self.current_folder_address,"PDF Files (*.pdf *.PDF)", options=options)
            except:
                fileName, _ = QFileDialog.getOpenFileName(self.ui, "QFileDialog.getOpenFileName()", "","PDF Files (*.pdf *.PDF)", options=options)
            return fileName
        except:
            traceback.print_exc()
    '''Change and reset file'''
    def f_file_change(self,file_dir_name):
        try:
            fileName=self.f_get_filename()
            folder=str(Path(fileName).parent.absolute())
            file=str(Path(fileName).name)
            if file!='':
                self.file_dir_dict[file_dir_name][0].setText(folder)
                self.file_dir_dict[file_dir_name][1].setText(file)
        except:
            traceback.print_exc()
    '''Copy file to SS folder'''
    def copy2ss(self,file):
        if file_exists(file):
            ss_name = os.path.join(self.current_folder_address, 'ss',
                                   date.today().strftime("%Y%m%d") + '-' + Path(file).name)
            if file_exists(ss_name):
                for i in range(100):
                    new_ss_name = ss_name[:-4] + '(' + str(i + 1) + ').pdf'
                    if file_exists(new_ss_name) == False:
                        ss_name = new_ss_name
                        break
            ss_dir = os.path.join(self.current_folder_address, 'ss')
            os.makedirs(ss_dir, exist_ok=True)
            shutil.copy(file, ss_name)
    '''================================Scale and size change set============================='''
    '''Original scale change for rescale'''
    def original_scale_text_change(self,name):
        try:
            text_name,buttongroup=self.original_scale_change_dict[name][0],self.original_scale_change_dict[name][1]
            if text_name.text()!='':
                buttongroup.button(6).setChecked(True)
            else:
                buttongroup.button(6).setChecked(False)
        except:
            traceback.print_exc()
    '''Original size change for rescale'''
    def original_size_text_change(self,name):
        try:
            text1,text2,buttongroup=self.original_size_change_dict[name][0],self.original_size_change_dict[name][1],self.original_size_change_dict[name][2]
            if text1.text()=='' and text2.text()=='':
                buttongroup.button(8).setChecked(False)
            else:
                buttongroup.button(8).setChecked(True)
        except:
            traceback.print_exc()
    '''Output scale change for rescale'''
    def outscale_click(self):
        return
        # try:
        #     if str(self.ui.text_sketch_rescale_2_oscx.text())!='':
        #         self.button_group3.button(3).setChecked(True)
        #         scale = '1:' + str(self.ui.text_sketch_rescale_2_oscx.text())
        #         self.ui.text_color_scale.setText(scale)
        #     else:
        #         self.button_group3.button(3).setChecked(False)
        # except:
        #     traceback.print_exc()
    '''Output size change for rescale'''
    def outsize_click(self):
        try:
            if str(self.ui.text_sketch_rescale_2_outputcussize1.text())=='' and str(self.ui.text_sketch_rescale_2_outputcussize2.text())=='':
                self.button_group4.button(8).setChecked(False)
            else:
                self.button_group4.button(8).setChecked(True)
                papersize=self.ui.text_sketch_rescale_2_outputcussize1.text()+':'+self.ui.text_sketch_rescale_2_outputcussize2.text()
                # self.ui.text_color_papersize.setText(papersize)
        except:
            traceback.print_exc()
    '''Output scale change for overlay service rescale'''
    def f_text_overlay_service1_outs_change(self):
        try:
            if str(self.ui.text_overlay_service1_outs.text())!='':
                self.button_group_overlay_service1_outs.button(3).setChecked(True)
            else:
                self.button_group_overlay_service1_outs.button(3).setChecked(False)
        except:
            traceback.print_exc()
    '''Output size change for overlay service rescale'''
    def f_overlay_service1_outsize(self):
        try:
            if str(self.ui.text_overlay_service_outsize1.text())=='' and str(self.ui.text_overlay_service_outsize2.text())=='':
                self.button_group_overlay_service1_outsize.button(8).setChecked(False)
            else:
                self.button_group_overlay_service1_outsize.button(8).setChecked(True)
        except:
            traceback.print_exc()
    '''Text set with block signal set'''
    def textset_with_blocksignal(self,text_name):
        try:
            text_name.blockSignals(True)
            text_name.setText('')
            text_name.blockSignals(False)
        except:
            traceback.print_exc()
    def textlistset_with_blocksignal(self,text_list):
        for text in text_list:
            self.textset_with_blocksignal(text)
    def textset_content_with_blocksignal(self,text_name,content):
        try:
            text_name.blockSignals(True)
            text_name.setText(content)
            text_name.blockSignals(False)
        except:
            traceback.print_exc()
    '''Buttongroup customized size text change'''
    def f_buttongroup_size_clicked(self,button,type):
        try:
            buttongroup=self.buttongroup_size_click_dict[type][0]
            text1=self.buttongroup_size_click_dict[type][1]
            text2=self.buttongroup_size_click_dict[type][2]
            button_id =buttongroup.id(button)
            if button_id!=8:
                self.textset_with_blocksignal(text1)
                self.textset_with_blocksignal(text2)
        except:
            traceback.print_exc()
    '''Buttongroup customized scale text change'''
    def f_buttongroup_scale(self,button,type):
        try:
            buttongroup=self.buttongroup_scale_dict[type][0]
            cutomize_id=self.buttongroup_scale_dict[type][1]
            cus_text=self.buttongroup_scale_dict[type][2]
            button_id = buttongroup.id(button)
            if button_id!=cutomize_id:
                cus_text.setText('')
            # if type=="Sketch output":
                # if button_id==1:
                #     self.ui.text_color_scale.setText('1:50')
                # elif button_id==2:
                #     self.ui.text_color_scale.setText('1:100')
        except:
            traceback.print_exc()
    '''Button group list clear'''
    def buttongroup_list_clear(self,buttongroup_list):
        for buttongroup in buttongroup_list:
            buttongroup.setExclusive(False)
            for button in buttongroup.buttons():
                button.setChecked(False)
            buttongroup.setExclusive(True)
    '''Button list clear'''
    def button_list_clear(self,button_list):
        for button in button_list:
            button.setChecked(False)
    '''===================================Project info frame==================================='''
    '''Log in set'''
    def f_button_log_login(self):
        try:
            if os.path.exists(conf["password_file"]):
                with open(conf["password_file"], 'r', encoding='utf-8') as file:
                    file_content = file.read()
                    content = ast.literal_eval(file_content)
                    user_name=content['user_name']
                    password = content['password']
                    self.ui.text_log_email.setText(user_name)
                    self.ui.text_log_password.setText(password)
            self.ui.frame_login.show()
        except:
            traceback.print_exc()
    '''Choose search history'''
    def f_choose_from_search_history(self):
        try:
            search_history=self.ui.combobox_search_history.currentText()
            if search_history!='':
                self.ui.combobox_search_history.currentIndexChanged.disconnect(self.f_choose_from_search_history)
                quo_num=search_history.split('-')[0]
                self.ui.text_all_1_searchbar.setText(quo_num)
                self.f_all_search()
                self.ui.combobox_search_history.currentIndexChanged.connect(self.f_choose_from_search_history)
        except:
            traceback.print_exc()
    def f_button_log_confirm(self):
        email_show_list = ['felix@pcen.com.au', 'Daniel@forgebc.com.au',
                           'admin@pcen.com.au','wayne@forgebc.com.au',
                           'engineer4@forgebc.com.au',
                           'yitong@forgebc.com.au', 'Engineer1@forgebc.com.au',
                           'engineer2@forgebc.com.au', 'engineer3@forgebc.com.au',
                           'engineer5@forgebc.com.au', ]

        try:
            email=self.ui.text_log_email.text()
            password=self.ui.text_log_password.text()
            if email=='' or password=='':
                message('Email or password not correct',self.ui)
                return
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.users WHERE user_email = %s AND user_password=%s", (email,password,))
                user_database=mycursor.fetchall()
                if len(user_database)==1:
                    self.ui.text_user_email.setText(email)
                    user_type=user_database[0][3]
                    asana_id=user_database[0][4]
                    table = self.ui.table_timesheet_user
                    table.setRowCount(0)
                    if user_type==1:
                        self.ui.text_user_type.setText('Admin')
                        self.ui.tabWidget.setTabVisible(10, True)
                        self.ui.tabWidget.setTabVisible(11, True)
                        self.ui.tabWidget.setTabVisible(12, True)
                        button_index_list = [2, 3, 4, 5, 6]
                        for index in button_index_list:
                            button = self.ui.findChild(QPushButton, 'button_search_hide_' + str(index))
                            button.show()
                        self.f_button_search_hide(region=2)
                        self.f_button_search_hide(region=3)
                        self.f_button_search_hide(region=4)
                        self.f_button_search_hide(region=5)
                        self.f_button_search_hide(region=6)

                        mycursor.execute("SELECT * FROM bridge.users")
                        user_all_database = mycursor.fetchall()
                        row=0
                        for user_i_database in user_all_database:
                            user_email=user_i_database[0]
                            if user_email in email_show_list:
                                table.insertRow(row)
                                category_dict={0:'Engineer',1:'Admin'}
                                category=category_dict[user_i_database[3]]
                                table.setItem(row,0, QTableWidgetItem(str(user_email)))
                                table.setItem(row,1, QTableWidgetItem(str(category)))
                                table.setItem(row,2, QTableWidgetItem(str(user_i_database[4])))
                                row+=1
                    else:
                        self.ui.text_user_type.setText('Engineer')
                        self.ui.tabWidget.setTabVisible(10, True)
                        self.ui.tabWidget.setTabVisible(11, True)
                        button_hide_name_list=['button_pro_state_genfee','button_pro_state_emailfee','button_pro_state_chasefee','button_pro_func_database','button_pro_func_updatexero','button_pro_func_refreshxero','button_pro_func_loginxero']
                        for button_name in button_hide_name_list:
                            button=self.ui.findChild(QPushButton, button_name)
                            button.hide()
                        frame_fee_hide=self.ui.findChild(QFrame, 'frame_fee_hide')
                        frame_fee_hide.hide()


                        table.insertRow(0)
                        table.setItem(0, 0, QTableWidgetItem(str(email)))
                        table.setItem(0, 1, QTableWidgetItem(self.ui.text_user_type.text()))
                        table.setItem(0, 2, QTableWidgetItem(str(asana_id)))
                    self.f_set_table_style(self.ui.table_timesheet_user)
                    self.ui.frame_login.hide()
                    user_password_dict={}
                    user_password_dict['user_name']=email
                    user_password_dict['password']=password
                    if not os.path.exists(conf["password_file"]):
                        os.makedirs(os.path.dirname(conf["password_file"]), exist_ok=True)
                    with open(conf["password_file"], 'w') as file:
                        file.write(str(user_password_dict))
                else:
                    message('Email or password not correct',self.ui)
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
        except:
            traceback.print_exc()

    def f_button_log_cancel(self):
        try:
            self.ui.frame_login.hide()
        except:
            traceback.print_exc()


    '''Open project and backupfolder'''
    def f_projectfolder_open(self,folder_name):
        try:
            folder_name_dict={"Project":self.current_folder_address,"Backup":os.path.join(conf["backup_folder"],Path(self.current_folder_address).name)}
            folder_dir=folder_name_dict[folder_name]
            quo_no=str(self.ui.text_all_1_projectnumber.text())
            if len(quo_no)==0:
                message('Please login a Project before open folder',self.ui)
                return
            if folder_name=="Backup":
                create_directory(folder_dir)
            if file_exists(self.current_folder_address):
                open_folder(folder_dir)
            else:
                message('Folder does not exist',self.ui)
        except:
            traceback.print_exc()
    '''Search content change'''
    def f_text_all_1_searchbar(self):
        try:
            search_content=self.ui.text_all_1_searchbar.text()
        except:
            traceback.print_exc()


    '''Search bar gst or not'''
    def f_checkbox_search_gst(self):
        try:
            if self.ui.checkbox_search_gst.isChecked()==False:
                if self.checkbox_search_gst_state==True:
                    self.checkbox_search_gst_state=False
                    table=self.ui.table_search_1
                    for row in range(table.rowCount()):
                        for column in list(range(27,32))+list(range(52,69))+[79,80,81,82]:
                            item = table.item(row, column)
                            original_background = item.background()
                            if item:
                                item_content=item.text()
                                if is_float(item_content):
                                    new_amount=round(float(item_content)/1.1,1)
                                    new_item = QTableWidgetItem(str(new_amount))
                                    new_item.setBackground(original_background)
                                    new_item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                                    table.setItem(row, column, new_item)
                        for column in [84,86,88,90,92]:
                            item = table.item(row, column)
                            item_last = table.item(row, column-1)
                            original_background = item.background()
                            if item:
                                item_last_color=item_last.background().color().name()
                                if item_last_color!='#cc9933':
                                    item_content = item.text()
                                    if is_float(item_content):
                                        new_amount = round(float(item_content) / 1.1, 1)
                                        new_item = QTableWidgetItem(str(new_amount))
                                        new_item.setBackground(original_background)
                                        new_item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                                        table.setItem(row, column, new_item)
            else:
                self.f_all_search_reset()
                self.checkbox_search_gst_state=True
        except:
            traceback.print_exc()


    '''Open asana'''
    def f_all_openasana(self):
        try:
            quo_no = str(self.ui.text_all_1_projectnumber.text())
            if len(quo_no) == 0:
                message('Please login a Project before open asana',self.ui)
                return
            if self.data_json["Asana_url"]==None or self.data_json["Asana_url"]=='':
                message('Asana id does not exist',self.ui)
            else:
                if self.data_json["Asana_url"].find('task'):
                    link=self.data_json["Asana_url"] + "?focus=true"
                else:
                    link=self.data_json["Asana_url"] + "?focus=true"
                open_link_with_edge(link)
        except:
            traceback.print_exc()

    '''Function when close software'''
    def on_close(self):
        try:
            quo_no = str(self.ui.text_all_1_projectnumber.text())
            if len(quo_no) != 0:
                self.savepdf()
            stats2.on_close()
        except:
            traceback.print_exc()
    '''Save project info when close software'''
    def savepdf(self):
        try:
            quo_num = self.ui.text_all_1_projectnumber.text()
            if quo_num=='':
                return

            selected_id1 = 0
            # self.button_group1.checkedId()
            input_scale_list = ['50', '75', '100', '150', '200']
            if selected_id1==6:
                input_scale=None
                custom_input_scale=str(self.ui.text_sketch_rescale_1_oscx.text())
            else:
                input_scale=input_scale_list[selected_id1 - 1]
                custom_input_scale=None

            selected_id2 = self.button_group2.checkedId()
            input_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24', 'A0-26']
            if selected_id2 == 8:
                input_size=None
                custom_input_size_x=str(self.ui.text_sketch_rescale_7_outputcussize1.text())
                custom_input_size_y=str(self.ui.text_sketch_rescale_8_outputcussize2.text())
            else:
                input_size=input_size_list[selected_id2 - 1]
                custom_input_size_x=None
                custom_input_size_y=None

            selected_id3 = self.button_group3.checkedId()
            output_scale_list = ['50', '100']
            if selected_id3 == 3:
                output_scale=None
                custom_output_scale=str(self.ui.text_sketch_rescale_2_oscx.text())
            else:
                output_scale=output_scale_list[selected_id3 - 1]
                custom_output_scale=None

            selected_id4 = self.button_group4.checkedId()
            output_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24', 'A0-26']
            if selected_id4==8:
                output_size=None
                custom_output_size_x=str(self.ui.text_sketch_rescale_2_outputcussize1.text())
                custom_output_size_y=str(self.ui.text_sketch_rescale_2_outputcussize1.text())
            else:
                output_size=output_size_list[selected_id4 - 1]
                custom_output_size_x=None
                custom_output_size_y=None

            tag_name_list=[]
            tag_include_list=[]
            for i in range(len(self.button_group_tag_textlist)):
                tag_name=self.button_group_tag_textlist[i].text()
                tag_name_list.append(tag_name)
                # if self.button_group_tag_boxlist[i].isChecked():
                #     include=1
                # else:
                #     include=0
                # tag_include_list.append(include)
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("UPDATE bridge.copilot_sizes SET input_scale = %s, input_size = %s, _scale = %s,"
                                 "_size_x  = %s, _size_y = %s, output_scale = %s, output_size = %s, m_scale = %s, m_size_x = %s, m_size_y = %s WHERE quotation_number=%s",
                    (input_scale,input_size,custom_input_scale,custom_input_size_x,custom_input_size_y,
                     output_scale,output_size,custom_output_scale,custom_output_size_x,custom_output_size_y,quo_num))
                for i in range(len(tag_include_list)):
                    mycursor.execute("UPDATE bridge.copilot_tags SET tag_level = %s, include = %s WHERE quotation_number=%s AND row_index=%s",
                                     (tag_name_list[i], tag_include_list[i], quo_num, i))


                mydb.commit()
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
        except:
            traceback.print_exc()
    '''=====================================Search tab========================================='''
    '''Search project'''
    def function_to_run(self):
        try:
            # if processing_or_not == False:
            self.ui.text_all_1_searchbar.setText('')
            self.f_all_search()
            self.clipboard_history=''
        except:
            traceback.print_exc()
    '''Search project and open folder'''
    def function_to_run2(self):
        try:
            print(processing_or_not)
            if processing_or_not==False:
                pro_name_ori=self.ui.text_all_3_projectname.text()
                clipboard_content = pyperclip.paste()
                self.ui.text_all_1_searchbar.setText(clipboard_content)
                self.f_all_search()
                pro_name_new=self.ui.text_all_3_projectname.text()
                self.f_projectfolder_open("Project")
        except:
            traceback.print_exc()
    '''Search - Double click to open folder'''
    def f_tablesearch_doubleclick(self):
        try:
            self.selection_changed()
            self.f_projectfolder_open("Project")
        except:
            traceback.print_exc()
    '''Project serach'''
    def projectsearch(self,searchcontent):
        try:
            searchresult={}
            pro_dict={}
            data = []
            for index,item in self.projects_all.items():
                for title in item:
                    if not title is None and str(searchcontent).lower() in str(title).lower():
                        data.append(item)
                        break
            if len(data)==1:
                searchresult['state']='fully_found'
            elif len(data)>1:
                searchresult['state']='part_found'
            else:
                searchresult['state'] = 'cannot_found'
            if len(data)>0:
                for i in range(len(data)):
                    quo_num=data[i][0]
                    pro_dict[quo_num]=data[i]

            searchresult['projectlist'] = data
            return searchresult

        except:
            traceback.print_exc()
    '''Search - search and set'''
    def f_all_search(self):
        try:
            self.ui.table_search_1.setSortingEnabled(False)
            search_content = str(self.ui.text_all_1_searchbar.text().strip())
            try:
                if search_content[:2]=='P:':
                    search_content=search_content.split("\\")[1].split('-')[0]
            except:
                pass
            try:
                self.ui.text_search_1.textChanged.disconnect(self.f_all_search2)
                self.ui.text_search_1.setText('')
                self.ui.text_search_1.textChanged.connect(self.f_all_search2)
            except:
                pass
            if search_content=='':
                self.change_table(self.projects_all)
            elif search_content=='cal show':
                self.ui.tabWidget.setTabVisible(7, True)
            elif search_content=='cal hide':
                self.ui.tabWidget.setTabVisible(7, False)
            else:
                search_result = self.projectsearch(search_content)
                project_list=search_result['projectlist']
                project_dict={}
                for i in range(len(project_list)):
                    project_dict[project_list[i][0]]=project_list[i]
                self.project_dict=project_dict
                if search_result['state']=='fully_found':
                    self.set_quotation_number(search_result['projectlist'][0])
                    self.change_table(project_dict)
                    for column in range(self.ui.table_search_1.columnCount()):
                        item = self.ui.table_search_1.item(0, column)
                        item.setSelected(True)
                elif search_result['state']=='part_found':
                    try:
                        self.ui.text_search_1.textChanged.disconnect(self.f_all_search2)
                        self.ui.text_search_1.setText(search_content)
                        self.ui.text_search_1.textChanged.connect(self.f_all_search2)
                    except:
                        self.ui.text_search_1.setText(search_content)
                    self.change_table(project_dict)
                    self.set_quotation_number(search_result['projectlist'][-1])
                    for row in range(self.ui.table_search_1.rowCount()):
                        quo_num = self.ui.table_search_1.item(row, 0).text()
                        if search_result['projectlist'][-1][0]==quo_num:
                            for column in range(self.ui.table_search_1.columnCount()):
                                item = self.ui.table_search_1.item(row, column)
                                item.setSelected(True)
                            break
                elif search_result['state']=='cannot_found':
                    try:
                        self.ui.text_search_1.textChanged.disconnect(self.f_all_search2)
                        self.ui.text_search_1.setText(search_content)
                        self.ui.text_search_1.textChanged.connect(self.f_all_search2)
                    except:
                        self.ui.text_search_1.setText(search_content)
                    self.change_table({})
            self.switch_tab('Search')
            self.ui.text_all_1_searchbar.setText('')
            self.ui.table_search_1.setSortingEnabled(True)
            self.ui.table_search_1.sortItems(0,Qt.DescendingOrder)
            self.ui.text_all_1_searchbar.setFocus()
        except:
            traceback.print_exc()
    '''Search - Reset search table'''
    def f_all_search_reset(self):
        try:
            self.ui.table_search_1.setSortingEnabled(False)
            if self.project_dict=={}:
                self.change_table(self.projects_all)
            else:
                self.change_table(self.project_dict)
            self.ui.table_search_1.setSortingEnabled(True)
        except:
            traceback.print_exc()

    '''Search - search and set'''
    def f_all_search2(self):
        try:
            self.ui.table_search_1.setSortingEnabled(False)
            search_content = str(self.ui.text_search_1.text().strip())
            if search_content=='':
                pass
            else:
                search_result = self.projectsearch(search_content)
                project_list=search_result['projectlist']
                project_dict={}
                for i in range(len(project_list)):
                    project_dict[project_list[i][0]]=project_list[i]
                if search_result['state']=='fully_found':
                    self.set_quotation_number(search_result['projectlist'][0])
                    self.change_table(project_dict)
                    for column in range(self.ui.table_search_1.columnCount()):
                        item = self.ui.table_search_1.item(0, column)
                        item.setSelected(True)
                elif search_result['state']=='part_found':
                    self.change_table(project_dict)
            self.ui.table_search_1.setSortingEnabled(True)
        except:
            traceback.print_exc()

    '''Search - Right click to copy'''
    def showContextMenu(self, pos):
        try:
            contextMenu = QMenu(self.ui)
            copyAction = QAction("Copy", self.ui)
            copyAction.triggered.connect(self.copySelection)
            contextMenu.addAction(copyAction)
            contextMenu.exec_(self.ui.table_search_1.viewport().mapToGlobal(pos))
        except:
            traceback.print_exc()
    '''Search table content copy'''
    def copySelection(self):
        try:
            selected_indexes = self.ui.table_search_1.selectedIndexes()
            if len(selected_indexes) > 0:
                selected_rows = list(set(index.row() for index in selected_indexes))
                rows_data = []
                for row in selected_rows:
                    row_data = [self.ui.table_search_1.item(row, col).text() for col in range(self.ui.table_search_1.columnCount())]
                    rows_data.append('\t'.join(row_data))
                clipboard = QApplication.clipboard()
                clipboard.setText('\n'.join(rows_data))
        except:
            traceback.print_exc()


    '''Search - Right click to copy'''
    def showContextMenu2(self, pos):
        try:
            contextMenu = QMenu(self.ui)
            copyAction = QAction("Copy", self.ui)
            copyAction.triggered.connect(self.copySelection2)
            contextMenu.addAction(copyAction)
            contextMenu.exec_(self.ui.table_kitchen_duct.viewport().mapToGlobal(pos))
        except:
            traceback.print_exc()
    '''Search table content copy'''
    def copySelection2(self):
        try:
            selected_indexes = self.ui.table_kitchen_duct.selectedIndexes()
            if len(selected_indexes) > 0:
                selected_rows = list(set(index.row() for index in selected_indexes))
                rows_data = []
                for row in selected_rows:
                    row_data = [self.ui.table_kitchen_duct.item(row, col).text() for col in range(self.ui.table_kitchen_duct.columnCount())]
                    rows_data.append('\t'.join(row_data))
                clipboard = QApplication.clipboard()
                clipboard.setText('\n'.join(rows_data))
        except:
            traceback.print_exc()


    '''Search add to search history'''
    def add_search_history(self):
        try:
            quo_num=self.ui.text_all_1_projectnumber.text()
            pro_num=self.ui.text_all_2_quotationnumber.text()
            pro_name=self.ui.text_all_3_projectname.text()
            search_data=quo_num+'-'+pro_num+'-'+pro_name
            if quo_num!='':
                search_history_num_list=[]
                for pro_info in self.search_history:
                    if pro_info!='':
                        quo_num_i=pro_info.split('-')[0]
                        search_history_num_list.append(quo_num_i)
                if quo_num not in search_history_num_list:
                    self.search_history.insert(1,search_data)
            if len(self.search_history)>6:
                self.search_history=self.search_history[:6]
            self.ui.combobox_search_history.clear()
            self.ui.combobox_search_history.addItems(self.search_history)
        except:
            traceback.print_exc()

    '''Search - Generate search bar'''
    def f_correct_invoice(self,invoice_num):
        try:
            if is_float(invoice_num):
                return invoice_num
            else:
                return None
        except:
            traceback.print_exc()
            return None

    def generate_search_bar(self,order):
        # to_change 1
        table_total_column=157
        self.inv_state_dict={}
        self_inv_state_dict={}
        self.bill_state_dict={}
        self_bill_state_dict={}
        self.bill_gst_dict={}
        self_bill_gst_dict = {}
        self.projects_all={}
        self_projects_all = {}

        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM bridge.projects")
            project_database = mycursor.fetchall()
            for i in range(len(project_database)):
                try:
                    pro_info_list_i = [None for _ in range(table_total_column)]
                    project_i_database = project_database[i]
                    quotation_number = project_i_database[0]
                    client_id = project_i_database[17]
                    main_contact_id = project_i_database[18]
                    mycursor.execute("SELECT * FROM bridge.emails WHERE quotation_number = %s", (quotation_number,))
                    email_database = mycursor.fetchall()[0]
                    if client_id!=None:
                        mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                        client_database = mycursor.fetchall()[0]
                    else:
                        client_database=[None for _ in range(8)]
                    if main_contact_id!=None:
                        mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (main_contact_id,))
                        main_contact_database = mycursor.fetchall()[0]
                    else:
                        main_contact_database=[None for _ in range(8)]
                    mycursor.execute("SELECT * FROM bridge.invoices WHERE quotation_number = %s", (quotation_number,))
                    invoice_database=mycursor.fetchall()
                    inv_list=[None for _ in range(8)]
                    inv_number_list=[None for _ in range(8)]
                    paid_fee=0
                    sent_and_paid_fee=0
                    overdue_fee=0
                    invoice_state_list=[None for _ in range(8)]

                    if len(invoice_database)>0:
                        for j in range(len(invoice_database)):
                            invoice_number=invoice_database[j][0]
                            invoice_amount=invoice_database[j][8]
                            state=invoice_database[j][2]
                            index=invoice_database[j][9]
                            if invoice_amount!=None:
                                inv_list[index]=round(invoice_amount*1.1,1)
                                inv_number_list[index]=invoice_number
                            if is_float(invoice_database[j][7]):
                                payment_amount = float(invoice_database[j][7])
                            else:
                                payment_amount=0
                            if is_float(invoice_database[j][8]):
                                fee_amount=float(invoice_database[j][8])*1.1
                            else:
                                fee_amount=0
                            paid_fee+=payment_amount
                            if state in ['Sent', 'Paid']:
                                overdue_fee+=(fee_amount-payment_amount)
                                sent_and_paid_fee+=fee_amount
                            color_dict = {"Backlog":(210,210,210),"Sent":(255,0,0),"Paid":(0,128,0),"Voided":(128,0,128)}
                            invoice_state_list[index]=color_dict[state]
                    self_inv_state_dict[quotation_number]=invoice_state_list

                    service_types_list = ['Mechanical Service', 'CFD Service', 'Electrical Service',
                                          'Hydraulic Service', 'Fire Service', 'Mechanical Review',
                                          'Miscellaneous', 'Installation', 'Variation']
                    mycursor.execute("SELECT * FROM bridge.fee_item WHERE quotation_number = %s", (quotation_number,))
                    fee_item_database=mycursor.fetchall()
                    service_fee_list=[None for _ in range(9)]
                    total_fee_amount=0
                    service_types_with_kitchen_list = ['Mechanical Service', 'CFD Service', 'Electrical Service',
                                                       'Hydraulic Service', 'Fire Service', 'Mechanical Service',
                                                       'Mechanical Review', 'Miscellaneous', 'Installation']
                    project_service_full_list=[]
                    for j in range(9):
                        if project_i_database[j + 8] == 1:
                            project_service_full_list.append(service_types_with_kitchen_list[j])
                    project_service_full_list.append('Variation')
                    for j in range(len(fee_item_database)):
                        fee_item_database_i=fee_item_database[j]
                        amount=fee_item_database_i[3]
                        if is_float(amount):
                            amount=round(float(amount)*1.1,1)
                            service=fee_item_database_i[1]
                            if service in project_service_full_list:
                                index=service_types_list.index(service)
                                if service_fee_list[index]==None:
                                    fee_old=0
                                else:
                                    fee_old=float(service_fee_list[index])
                                service_fee_list[index]=round(fee_old+amount,1)
                                total_fee_amount+=amount

                    mycursor.execute("SELECT * FROM bridge.bills WHERE quotation_number = %s", (quotation_number,))
                    bill_database=mycursor.fetchall()
                    bill_service_list=[None for _ in range(5)]
                    bill_amount_list=[None for _ in range(5)]
                    bill_state_list=[None for _ in range(5)]
                    bill_gst_list=[None for _ in range(5)]
                    total_paid_bill=0
                    total_bill_amount=0
                    color_dict = {"Draft": (210, 210, 210), "Awaiting Approval": (255, 0, 0), "Paid": (0, 128, 0),
                                  "Voided": (128, 0, 128), "Awaiting Payment": (255, 165, 0)}
                    color_dict_bill_gst={1:(204, 153, 51),0:(51, 153, 204)}
                    if len(bill_database)>0:
                        for j in range(len(bill_database)):
                            bill_database_i=bill_database[j]
                            bill_num=char2num(bill_database_i[0][-1])
                            if bill_num in [0,1,2,3,4]:
                                service=bill_database_i[9]
                                service_combo=bill_database_i[8]
                                if bill_database_i[10] is None:
                                    amount=0
                                else:
                                    amount=round(float(bill_database_i[10])*1.1,1)
                                total_bill_amount+=amount
                                state=bill_database_i[2]
                                gst=bill_database_i[11]
                                bill_service_list[bill_num]=service_combo
                                bill_amount_list[bill_num]=amount
                                bill_state_list[bill_num]=color_dict[state]
                                if gst==1:
                                    bill_gst_list[bill_num]=color_dict_bill_gst[gst]
                                if state=="Paid":
                                    total_paid_bill+=amount
                    self_bill_state_dict[quotation_number]=bill_state_list
                    self_bill_gst_dict[quotation_number] = bill_gst_list

                    email_name_dict = {'felix@pcen.com.au': 'FY', 'Daniel@forgebc.com.au': 'DW',
                                       'admin@pcen.com.au': 'Iza', 'alex@forgebc.com.au': 'AX',
                                       'engineer4@forgebc.com.au': 'SL',
                                       'yitong@forgebc.com.au': 'YX', 'Engineer1@forgebc.com.au': 'Intern',
                                       'engineer2@forgebc.com.au': 'Intern', 'engineer3@forgebc.com.au': 'Intern',
                                       'engineer5@forgebc.com.au': 'Intern', }
                    worker_time_dict = {'FY': [0, 0], 'DW': [0, 0], 'Iza': [0, 0], 'AX': [0, 0], 'SL': [0, 0], 'YX': [0, 0],
                                        'Intern': [0, 0]}
                    mycursor.execute("SELECT * FROM bridge.man_power WHERE quotation_number = %s", (quotation_number,))
                    man_power_database=mycursor.fetchall()
                    if len(man_power_database)>0:
                        for j in range(len(man_power_database)):
                            email_from_database=man_power_database[j][1]
                            est_from_database=man_power_database[j][2]
                            act_from_database = man_power_database[j][3]
                            worker_name=email_name_dict[email_from_database]
                            if is_float(est_from_database):
                                old_time=worker_time_dict[worker_name][0]
                                worker_time_dict[worker_name][0]=old_time+int(est_from_database)
                            if is_float(act_from_database):
                                old_time=worker_time_dict[worker_name][1]
                                worker_time_dict[worker_name][1]=old_time+int(act_from_database)

                    # to_change 2
                    pro_info_list_i[0] = quotation_number
                    pro_info_list_i[1] = project_i_database[1]
                    pro_info_list_i[2] = project_i_database[2]
                    pro_info_list_i[4] = project_i_database[5]
                    pro_info_list_i[5] = project_i_database[6]
                    pro_info_list_i[6] = project_i_database[7]
                    status_database=project_i_database[3]
                    status_dict={'Set Up':'01.Set Up','Gen Fee Proposal':'02.Gen Fee Proposal','Email Fee Proposal':'03.Email Fee Proposal',
                                 'Chase Fee Acceptance':'04.Chase Fee Acceptance','Design':'05.Design','Pending':'06.Pending',
                                 'DWG drawings':'07.DWG drawings','Done':'08.Done','Installation':'09.Installation',
                                 'Construction Phase':'10.Construction Phase','Quote Unsuccessful':'11.Quote Unsuccessful'}
                    real_status=status_database
                    try:
                        real_status=status_dict[status_database]
                    except:
                        pass
                    pro_info_list_i[3] =real_status
                    service_list = ['Me', 'CFD', 'El', 'Hy', 'Fi', 'Ki', 'MR', 'Mis', 'In']
                    services = ''
                    for j in range(9):
                        if project_i_database[j + 8] == 1:
                            services += service_list[j] + ','
                    if services!='':
                        services=services[:-1]
                    pro_info_list_i[7] = services
                    pro_info_list_i[9] = email_database[1]
                    pro_info_list_i[10] = email_database[2]
                    pro_info_list_i[11] = project_i_database[27]
                    pro_info_list_i[12] = email_database[3]
                    pro_info_list_i[13] = email_database[4]
                    pro_info_list_i[14] = email_database[5]
                    pro_info_list_i[15] = project_i_database[19]
                    pro_info_list_i[16] = project_i_database[20]
                    pro_info_list_i[17] = project_i_database[21]
                    pro_info_list_i[18] = project_i_database[4][0]
                    pro_info_list_i[19] = client_database[2]
                    pro_info_list_i[20] = client_database[3]
                    pro_info_list_i[21] = client_database[4]
                    pro_info_list_i[22] = main_contact_database[2]
                    pro_info_list_i[23] = main_contact_database[3]
                    pro_info_list_i[24] = main_contact_database[4]
                    if order==2:
                        all_subtasks = {}
                        # todo: get info from asana
                    else:
                        pro_info_list_i[103] = worker_time_dict['FY'][0]
                        pro_info_list_i[104] = worker_time_dict['FY'][1]
                        pro_info_list_i[105] = worker_time_dict['DW'][0]
                        pro_info_list_i[106] = worker_time_dict['DW'][1]
                        pro_info_list_i[107] = worker_time_dict['Iza'][0]
                        pro_info_list_i[108] = worker_time_dict['Iza'][1]
                        pro_info_list_i[109] = worker_time_dict['AX'][0]
                        pro_info_list_i[110] = worker_time_dict['AX'][1]
                        pro_info_list_i[111] = worker_time_dict['SL'][0]
                        pro_info_list_i[112] = worker_time_dict['SL'][1]
                        pro_info_list_i[113] = worker_time_dict['YX'][0]
                        pro_info_list_i[114] = worker_time_dict['YX'][1]
                        pro_info_list_i[115] = worker_time_dict['Intern'][0]
                        pro_info_list_i[116] = worker_time_dict['Intern'][1]
                    # todo
                    if worker_time_dict['FY'][1]<worker_time_dict['DW'][1]:
                        pro_info_list_i[8] = 'DW'
                    else:
                        pro_info_list_i[8] = 'FY'
                    pro_info_list_i[61] = self.f_correct_invoice(inv_list[0])
                    pro_info_list_i[62] = self.f_correct_invoice(inv_list[1])
                    pro_info_list_i[63] = self.f_correct_invoice(inv_list[2])
                    pro_info_list_i[64] = self.f_correct_invoice(inv_list[3])
                    pro_info_list_i[65] = self.f_correct_invoice(inv_list[4])
                    pro_info_list_i[66] = self.f_correct_invoice(inv_list[5])
                    pro_info_list_i[67] = self.f_correct_invoice(inv_list[6])
                    pro_info_list_i[68] = self.f_correct_invoice(inv_list[7])
                    pro_info_list_i[69] = self.f_correct_invoice(inv_number_list[0])
                    pro_info_list_i[70] = self.f_correct_invoice(inv_number_list[1])
                    pro_info_list_i[71] = self.f_correct_invoice(inv_number_list[2])
                    pro_info_list_i[72] = self.f_correct_invoice(inv_number_list[3])
                    pro_info_list_i[73] = self.f_correct_invoice(inv_number_list[4])
                    pro_info_list_i[74] = self.f_correct_invoice(inv_number_list[5])
                    pro_info_list_i[75] = self.f_correct_invoice(inv_number_list[6])
                    pro_info_list_i[76] = self.f_correct_invoice(inv_number_list[7])
                    pro_info_list_i[79] = round(paid_fee,1)
                    pro_info_list_i[80] = round(overdue_fee,1)
                    pro_info_list_i[52] = service_fee_list[0]
                    pro_info_list_i[53] = service_fee_list[1]
                    pro_info_list_i[54] = service_fee_list[2]
                    pro_info_list_i[55] = service_fee_list[3]
                    pro_info_list_i[56] = service_fee_list[4]
                    pro_info_list_i[57] = service_fee_list[5]
                    pro_info_list_i[58] = service_fee_list[6]
                    pro_info_list_i[59] = service_fee_list[7]
                    pro_info_list_i[60] = service_fee_list[8]
                    pro_info_list_i[81] = round(total_bill_amount,1)
                    pro_info_list_i[83] = bill_service_list[0]
                    pro_info_list_i[84] = bill_amount_list[0]
                    pro_info_list_i[85] = bill_service_list[1]
                    pro_info_list_i[86] = bill_amount_list[1]
                    pro_info_list_i[87] = bill_service_list[2]
                    pro_info_list_i[88] = bill_amount_list[2]
                    pro_info_list_i[89] = bill_service_list[3]
                    pro_info_list_i[90] = bill_amount_list[3]
                    pro_info_list_i[91] = bill_service_list[4]
                    pro_info_list_i[92] = bill_amount_list[4]
                    # pro_info_list_i[82] = '$'+str(round(total_paid_bill,1))
                    # pro_info_list_i[31] = '$'+str(round(total_fee_amount,1))
                    pro_info_list_i[82] = str(round(total_paid_bill,1))
                    pro_info_list_i[31] = str(round(total_fee_amount,1))
                    if total_fee_amount == 0:
                        invoiced_per = '0%'
                    else:
                        invoiced_per = str(int(sent_and_paid_fee * 100 / total_fee_amount)) + '%'
                    pro_info_list_i[26] = invoiced_per
                    try:
                        txt_name = os.path.join(conf["database_dir"], 'finan_data.txt')
                        with open(txt_name, 'r', encoding='utf-8') as file:
                            file_content = file.read()
                            content = ast.literal_eval(file_content)
                        salary_per_hour_list = content['salary_per_hour_list']
                        k1 = content['k1']
                        k2 = content['k2']
                    except:
                        traceback.print_exc()
                        message('Get financial data failed',self.ui)
                        salary_per_hour_list = [300, 280, 150, 150, 300, 150, 120]
                        k1=0.3
                        k2=1
                    salary_sum = 0
                    for j in range(len(salary_per_hour_list)):
                        salary_sum+=salary_per_hour_list[j]*pro_info_list_i[104+j*2]/60
                    # pro_info_list_i[27] = '$'+str(int(salary_sum))
                    # pro_info_list_i[28] = '$'+str(round(total_bill_amount,1))
                    # pro_info_list_i[29] = '$'+str(round(int(salary_sum)*k2,1))
                    # pro_info_list_i[30] = '$'+str(round(total_fee_amount-int(salary_sum)-total_bill_amount-int(salary_sum)*k2,1))
                    pro_info_list_i[27] = str(int(salary_sum))
                    pro_info_list_i[28] = str(round(total_bill_amount,1))
                    pro_info_list_i[29] = str(round(int(salary_sum)*k2,1))
                    pro_info_list_i[30] = str(round(sent_and_paid_fee-int(salary_sum)-total_bill_amount-int(salary_sum)*k2,1))
                    if total_fee_amount==0:
                        progress_per='0%'
                    else:
                        progress_per=str(int((int(salary_sum)+total_bill_amount)*100/total_fee_amount/k1))+'%'
                    pro_info_list_i[25] = progress_per
                    if sent_and_paid_fee!=0:
                        profit_per=str(round((sent_and_paid_fee-int(salary_sum)-total_bill_amount-int(salary_sum)*k2)*100/sent_and_paid_fee,1))+'%'
                    else:
                        profit_per='0%'
                    pro_info_list_i[32]=profit_per

                    self_projects_all[pro_info_list_i[0]]=pro_info_list_i


                    self.inv_state_dict = self_inv_state_dict
                    self.bill_state_dict = self_bill_state_dict
                    self.bill_gst_dict = self_bill_gst_dict
                    self.projects_all = self_projects_all
                except:
                    print(quotation_number)
                    traceback.print_exc()

        except:
            print(quotation_number)
            traceback.print_exc()
            return
        finally:
            mycursor.close()
            mydb.close()


    '''Search - Set project info'''
    def set_quotation_number(self,project_info):
        try:
            self.f_all_clearup()
        except:
            return
        try:
            self.disconnect_all()
        except:
            return
        try:
            self.ui.setEnabled(False)
            self.ui.text_all_1_projectnumber.setText(project_info[0])
            self.ui.text_all_2_quotationnumber.setText(project_info[1])
            self.ui.text_all_3_projectname.setText(project_info[2])
            self.ui.text_all_3_projectname.setCursorPosition(0)
            self.add_search_history()
            self.load_bridge_json()
            self.quotation_number = self.data_json["Quotation_number"]
            self.current_folder_address = os.path.join(conf["project_folder"], self.data_json["Current_folder_address"])
            self.project_name = self.data_json["Project_name"]
            self.load_pdfjson()
            self.ui.table_combine_1_filename.setRowCount(0)
            external_folder=os.path.join(self.current_folder_address,'External')
            self.ui.text_filing_folder.setText(external_folder)
            external_folder_name=date.today().strftime("%Y%m%d") +'- Drawing'
            self.ui.text_filing_name.setText(external_folder_name)
            self.file_check_filing()
            self.ui.text_overlay_folder1.setText(os.path.join(self.current_folder_address,'Drafting'))
            overlay_folder_name=date.today().strftime("%Y%m%d") +'-Overlay'
            self.ui.text_overlay_folder2.setText(overlay_folder_name)
            try:
                selected_items = self.tableWidget3.selectedItems()
                self.change_content([selected_items[4].text(), selected_items[5].text(), selected_items[6].text(),
                                     selected_items[8].text()])
            except:
                pass
            self.backup_folder=os.path.join(conf["backup_folder"], Path(self.data_json["Current_folder_address"]).name)

            create_directory(self.backup_folder)
            ss_folder=os.path.join(self.backup_folder,'ss')
            print(ss_folder)
            create_directory(ss_folder)
            project_type_dict={8:'Mechanical Service',9:'CFD Service',10:'Electrical Service',11:'Hydraulic Service',
                               12:'Fire Service',13:'Mechanical Service',14:'Mechanical Review',15:'Miscellaneous',16:'Installation'}
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.projects WHERE quotation_number = %s", (project_info[0],))
                project_database = mycursor.fetchall()
                if len(project_database)==1:
                    project_database=project_database[0]
                    proposal_type=project_database[6]
                    service_list=[]
                    for index,service_name in project_type_dict.items():
                        include=project_database[index]
                        if include==1:
                            service_list.append(service_name)
                    service_list=list(set(service_list))
                    if len(service_list)>0:
                        placeholders = ', '.join(['%s'] * len(service_list))
                        query = f"SELECT * FROM bridge.scope WHERE quotation_number = %s AND minor_or_major = %s AND service IN ({placeholders}) AND include = %s"
                        params = (project_info[0], proposal_type) + tuple(service_list) + (1,)
                        mycursor.execute(query, params)
                        scope_database = mycursor.fetchall()
                    else:
                        scope_database = []
            except:
                traceback.print_exc()
                return
            finally:
                mycursor.close()
                mydb.close()
            table=self.ui.table_scope_all
            type_index=1
            type_index_dict={"Extent":1,"Clarifications":2,"Deliverables":3}
            if len(scope_database)>1:
                for i in range(len(scope_database)):
                    service=scope_database[i][2]
                    type=scope_database[i][3]
                    content=scope_database[i][5]
                    row_position=table.rowCount()
                    table.insertRow(row_position)
                    type_index_now=type_index_dict[type]
                    if type_index_now!=type_index:
                        type_index=type_index_now
                    else:
                        table.setItem(row_position, 0, QTableWidgetItem(service))
                        table.setItem(row_position, 1, QTableWidgetItem(type))
                        table.setItem(row_position, 2, QTableWidgetItem(content))
            self.f_set_table_style(table)
        except:
            traceback.print_exc()
        finally:
            self.ui.setEnabled(True)
        try:
            self.f_get_links_from_database()
            self.ui.frame_link_files.hide()
        except:
            traceback.print_exc()
        try:
            self.connect_all()
        except:
            traceback.print_exc()
    '''Search - Set project info'''
    def f_search_fileadded(self,search_content):
        try:
            if search_content=='':
                return
            else:
                search_result = self.projectsearch(search_content)
                if search_result['state']=='fully_found':
                    self.set_quotation_number(search_result['projectlist'][0])
                else:
                    return
        except:
            traceback.print_exc()
    '''Search - Get bridge database'''
    def load_bridge_json(self):
        try:
            self.data_json = {}
            quo_num=self.ui.text_all_1_projectnumber.text()
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.projects WHERE quotation_number = %s",(quo_num,))
                project_database=mycursor.fetchall()
            except:
                traceback.print_exc()
                return
            finally:
                mycursor.close()
                mydb.close()
            self.data_json["Quotation_number"]=project_database[0][0]
            self.data_json["Project_name"]=project_database[0][2]
            self.data_json["Asana_url"]=project_database[0][24]
            self.data_json["Current_folder_address"]=project_database[0][25]
        except:
            traceback.print_exc()
    '''Search - Load project sketch info'''
    def load_project(self, quotation_number):
        project = format_output(get_value_from_table_with_filter("projects", "quotation_number", quotation_number))[quotation_number]
        self.current_folder_address = os.path.join(self.working_dir, project["current_folder_address"])
        self.data_json = {}
        self.data_json["Quotation_number"] = project["quotation_number"]
        self.data_json["Project_name"] = project["project_name"]
        self.data_json["Asana_url"] = project["asana_url"]
        self.data_json["Current_folder_address"] = os.path.join(self.working_dir, project["current_folder_address"])
        self.ui.text_all_1_projectnumber.setText(project["quotation_number"])
        self.ui.text_all_2_quotationnumber.setText(project["project_number"])
        self.ui.text_all_3_projectname.setText(project["project_name"])

        self.sketch_tab.current_folder = os.path.join(self.working_dir, project["current_folder_address"])
        self.sketch_tab.project_name = project["project_name"]
    def load_pdfjson(self):
        try:
            quo_num=self.ui.text_all_1_projectnumber.text()
            if quo_num=='':
                return
            self.ui.text_sketch_rescale_1_oscx.setText('')
            text_set_none_list=[self.ui.text_sketch_rescale_7_outputcussize1,self.ui.text_sketch_rescale_8_outputcussize2,self.ui.text_sketch_rescale_2_oscx,self.ui.text_sketch_rescale_2_outputcussize2]
            for text in text_set_none_list:
                self.textset_with_blocksignal(text)
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.copilot_sizes WHERE quotation_number = %s",(quo_num,))
                sizes_database=mycursor.fetchall()
                mycursor.execute("SELECT * FROM bridge.copilot_tags WHERE quotation_number = %s",(quo_num,))
                tags_database = mycursor.fetchall()[:27]
            except:
                traceback.print_exc()
                return
            finally:
                mycursor.close()
                mydb.close()
            if len(sizes_database)!=1:
                return
            sizes_database=sizes_database[0]
            input_scale=str(sizes_database[1])
            input_size=str(sizes_database[2])
            custom_input_scale=str(sizes_database[3])
            custom_input_size_x=str(sizes_database[4])
            custom_input_size_y=str(sizes_database[5])
            output_scale=str(sizes_database[6])
            output_size=str(sizes_database[7])
            custom_output_scale=str(sizes_database[8])
            custom_output_size_x=str(sizes_database[9])
            custom_output_size_y=str(sizes_database[10])
            input_scale_list=['50','75','100','150','200']
            if input_scale in input_scale_list:
                input_scale_id=input_scale_list.index(input_scale)+1
            else:
                input_scale_id=6
                self.ui.text_sketch_rescale_1_oscx.setText(custom_input_scale)
            try:
                pass
                # self.button_group1.button(input_scale_id).setChecked(True)
            except:
                traceback.print_exc()

            input_size_list=['A4','A3','A2','A1','A0','A0-24','A0-26']
            if input_size in input_size_list:
                input_size_id=input_size_list.index(input_size)+1
            else:
                input_size_id=8
                self.textset_content_with_blocksignal(self.ui.text_sketch_rescale_7_outputcussize1, custom_input_size_x)
                self.textset_content_with_blocksignal(self.ui.text_sketch_rescale_8_outputcussize2, custom_input_size_y)
            try:
                self.button_group2.button(input_size_id).setChecked(True)
            except:
                traceback.print_exc()
            output_scale_list=['50','100']
            if output_scale in output_scale_list:
                output_scale_id=output_scale_list.index(output_scale)+1
                # self.ui.text_color_scale.setText('1:'+output_scale_list[output_scale_id-1])
            else:
                output_scale_id=3
                self.ui.text_sketch_rescale_2_oscx.setText(custom_output_scale)
                # self.ui.text_color_scale.setText(custom_output_scale)
            try:
                self.button_group3.button(output_scale_id).setChecked(True)
            except:
                traceback.print_exc()
            output_size_list=['A4','A3','A2','A1','A0','A0-24','A0-26']
            if output_size in output_size_list:
                output_size_id=output_size_list.index(output_size)+1
                # self.ui.text_color_papersize.setText(output_size_list[output_size_id - 1])
            else:
                output_size_id=8
                self.textset_content_with_blocksignal(self.ui.text_sketch_rescale_2_outputcussize1,custom_output_size_x)
                self.textset_content_with_blocksignal(self.ui.text_sketch_rescale_2_outputcussize2,custom_output_size_y)
                # self.ui.text_color_papersize.setText(custom_output_size_x + ':' + custom_output_size_y)
            try:
                self.button_group4.button(output_size_id).setChecked(True)
            except:
                traceback.print_exc()
            # self.ui.table_sketch_rescale_1_filename.setRowCount(0)
            # self.ui.tableWidget_a.setRowCount(0)
            # self.ui.tableWidget_b.setRowCount(0)
            # self.ui.tableWidget_c.setRowCount(0)
            # self.ui.tableWidget_d.setRowCount(0)

            for i in range(len(tags_database)):
                tag_level=tags_database[i][1]
                include=tags_database[i][2]
                row_index=tags_database[i][3]
                # try:
                #     self.button_group_tag_textlist[row_index].setText(tag_level)
                #     if include==1:
                #         self.button_group_tag_boxlist[row_index].setChecked(True)
                #     else:
                #         self.button_group_tag_boxlist[row_index].setChecked(False)
                # except:
                #     traceback.print_exc()
        except:
            traceback.print_exc()

    '''Search - '''
    def f_button_search_hide(self, region):
        # to_change 4
        try:
            region_hide_dict={1:range(4,18),2:range(18,25),3:range(25,52),4:range(52,81),5:range(81,103),6:range(103,157)}
            button=self.ui.findChild(QPushButton, 'button_search_hide_'+str(region))
            button_text=button.text()
            if button_text[-1]=='-':
                for column in region_hide_dict[region]:
                    self.tableWidget3.setColumnHidden(column, True)
                button_name_new=button_text[:-1]+'+'
                button.setText(button_name_new)
                button.setStyleSheet(conf["Gray button style"])

            else:
                for column in region_hide_dict[region]:
                    self.tableWidget3.setColumnHidden(column, False)
                button_name_new=button_text[:-1]+'-'
                button.setText(button_name_new)
                color_region_dict={1:"Color button style 1",2:"Color button style 2",3:"Color button style 3",
                                   4:"Color button style 4",5:"Color button style 5",6:"Color button style 6",}
                button.setStyleSheet(conf[color_region_dict[region]])

        except:
            traceback.print_exc()


    '''Search - Set table style'''
    def change_table(self,project_dict):
        try:
            self.ui.checkbox_search_gst.stateChanged.disconnect(self.f_checkbox_search_gst)
        except:
            pass
        try:
            self.ui.checkbox_search_gst.setChecked(True)
            # to_change 3
            self.ui.table_search_1.setRowCount(0)
            project_color_dict={'05.Design':(131,201,169),'06.Pending':(240,106,106),'07.DWG drawings':(236,141,113),
                                '08.Done':(69,115,210),'09.Installation':(78,203,196),'10.Construction Phase':(242,111,178),
                                '11.Quote Unsuccessful':(249,170,239)}
            i=-1
            if project_dict=={}:
                return
            for index,item in project_dict.items():
                i+=1
                self.ui.table_search_1.setRowCount(i + 1)
                for j in range(len(item)):
                    if item[j]==None:
                        content=''
                    else:
                        content=str(item[j])
                    if j==0:
                        quotation_number=content
                    newItem = QTableWidgetItem(content)
                    if j in [2,17,21,24]:
                        newItem.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    else:
                        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    newItem.setForeground(QColor(255, 255, 255))
                    self.ui.table_search_1.setItem(i, j, newItem)
                    if i % 2 == 0:
                        self.ui.table_search_1.item(i, j).setBackground(QColor(240, 240, 240))
                    else:
                        self.ui.table_search_1.item(i, j).setBackground(QColor(255, 255, 255))
                    if j==3:
                        if content in project_color_dict.keys():
                            rdg1=project_color_dict[content][0]
                            rdg2 = project_color_dict[content][1]
                            rdg3 = project_color_dict[content][2]
                            self.ui.table_search_1.item(i, j).setBackground(QColor(rdg1, rdg2, rdg3))
                    if j in [61,62,63,64,65,66,67,68]:
                        try:
                            color_list=self.inv_state_dict[quotation_number][j-61]
                            rdg1=color_list[0]
                            rdg2=color_list[1]
                            rdg3=color_list[2]
                            self.ui.table_search_1.item(i, j).setBackground(QColor(rdg1, rdg2, rdg3))
                        except:
                            pass
                    if j in [83,85,87,89,91]:
                        bill_index=[83,85,87,89,91].index(j)
                        if self.bill_gst_dict[index][bill_index]!=None:
                            color=self.bill_gst_dict[index][bill_index]
                            rdg1=color[0]
                            rdg2=color[1]
                            rdg3=color[2]
                            self.ui.table_search_1.item(i, j).setBackground(QColor(rdg1, rdg2, rdg3))
                    if j in [84,86,88,90,92]:
                        bill_index=[84,86,88,90,92].index(j)
                        if self.bill_state_dict[index][bill_index]!=None:
                            color=self.bill_state_dict[index][bill_index]
                            rdg1=color[0]
                            rdg2=color[1]
                            rdg3=color[2]
                            self.ui.table_search_1.item(i, j).setBackground(QColor(rdg1, rdg2, rdg3))
            try:
                font = QFont()
                font.setPointSize(conf['Table font'])
                self.tableWidget3.setFont(font)
                if self.tableWidget3.rowCount()!=0:
                    for row in range(self.tableWidget3.rowCount()):
                        self.tableWidget3.setRowHeight(row, 20)
                print('CHANGE TABLE')
            except:
                pass
        except:
            traceback.print_exc()
        self.ui.checkbox_search_gst.stateChanged.connect(self.f_checkbox_search_gst)

    '''Search - Refresh'''
    def f_search_refresh2(self):
        try:
            self.generate_search_bar(order=1)
            self.f_search_refresh()
        except:
            traceback.print_exc()

    '''Search - Refresh'''
    def f_search_refresh(self):
        try:
            self.ui.table_search_1.setSortingEnabled(False)
            self.change_table(self.projects_all)
            self.ui.table_search_1.setSortingEnabled(True)

        except:
            traceback.print_exc()
    '''Search - Set project content when table click'''
    def selection_changed(self):
        try:
            selected_items = self.tableWidget3.selectedItems()
            if selected_items:
                qua_no = selected_items[0].text()
                proj_no = selected_items[1].text()
                address_no = selected_items[2].text()
                self.set_quotation_number([qua_no,proj_no,address_no])
        except:
            traceback.print_exc()
    '''Search - Set project info in Drawing'''
    def change_content(self,bridge_content):
        try:
            self.ui.text_drawing_scale.setText(bridge_content[0])
            self.ui.text_drawing_type.setText(bridge_content[1])
            self.ui.text_drawing_status.setText(bridge_content[2])
            self.ui.text_drawing_service.setText(bridge_content[3])
        except:
            traceback.print_exc()
    '''Text list set clear'''
    def text_list_clear(self,text_list):
        for text in text_list:
            text.setText('')
    '''Table list set clear'''
    def table_list_clear(self,table_list):
        for table in table_list:
            table.setRowCount(0)
    '''Clear up all'''
    def f_button_all_clearup(self):
        try:
            self.f_all_clearup()
            self.f_all_search()
        except:
            traceback.print_exc()

    def f_all_clearup(self):
        try:
            self.disconnect_all()
        except:
            traceback.print_exc()
        try:
            if self.ui.text_all_1_projectnumber.text()=='':
                return
            '''Save project info'''
            self.on_close()
            '''Project search frame reset'''
            text_list_pro=[self.ui.text_all_1_projectnumber,self.ui.text_all_2_quotationnumber,self.ui.text_all_3_projectname,self.ui.text_all_1_searchbar]
            self.text_list_clear(text_list_pro)
            '''Scope tab reset'''
            self.ui.table_scope_all.setRowCount(0)
            '''Filing tab reset'''
            text_list_filing=[self.ui.text_filing_folder,self.ui.text_filing_name]
            self.text_list_clear(text_list_filing)
            table_list_filing=[self.ui.table_filling]
            self.table_list_clear(table_list_filing)
            '''Combine tab reset'''
            text_list_combine = [self.ui.text_combine_1_foldername,self.ui.text_combine_op,self.ui.text_combine_3_filenameadd]
            self.text_list_clear(text_list_combine)
            table_list_combine=[self.ui.table_combine_1_filename]
            self.table_list_clear(table_list_combine)
            '''Sketch - Rescale reset'''
            text_list_rescale = [self.ui.text_sketch_rescale_1_filedir,self.ui.text_sketch_rescale_1_filesavedir,
                                 # self.ui.text_sketch_rescale_2_filename,
                                 self.ui.text_rescale_cf_1,self.ui.text_rescale_of_1,self.ui.text_rescale_outfile_1,
                                 self.ui.text_rescale_cf_2,self.ui.text_rescale_of_2,self.ui.text_rescale_outfile_2,
                                 self.ui.text_rescale_cf_3,self.ui.text_rescale_of_3,self.ui.text_rescale_outfile_3,
                                 self.ui.text_rescale_cf_4,self.ui.text_rescale_of_4,self.ui.text_rescale_outfile_4,
                                 self.ui.text_sketch_rescale_1_oscx,self.ui.text_rescale_a,self.ui.text_rescale_b,self.ui.text_rescale_c,self.ui.text_rescale_d,
                                 self.ui.text_sketch_rescale_2_oscx]
            self.text_list_clear(text_list_rescale)
            # table_list_rescale = [self.ui.table_sketch_rescale_1_filename,self.ui.tableWidget_a,self.ui.tableWidget_b,self.ui.tableWidget_c,self.ui.tableWidget_d]
            # self.table_list_clear(table_list_rescale)
            text_block_list=[self.ui.text_sketch_rescale_7_outputcussize1,self.ui.text_sketch_rescale_8_outputcussize2,self.ui.text_resize1_a,
                             self.ui.text_resize2_a,self.ui.text_resize1_b,self.ui.text_resize2_b,self.ui.text_resize1_c,self.ui.text_resize2_c,
                             self.ui.text_resize1_d,self.ui.text_resize2_d,self.ui.text_sketch_rescale_2_outputcussize1,self.ui.text_sketch_rescale_2_outputcussize2]
            self.textlistset_with_blocksignal(text_block_list)
            buttongroup_list_rescale=[
                # self.button_group1,
                #                       self.button_group2,self.button_group3,self.button_group4,
            ]
                                      # self.button_group_a1,self.button_group_a2, self.button_group_b1,self.button_group_b2,self.button_group_c1,self.button_group_c2,self.button_group_d1,self.button_group_d2]
            self.buttongroup_list_clear(buttongroup_list_rescale)
            '''Sketch - Align reset'''
            text_list_align = [self.ui.text_sketch_align_filedir,self.ui.text_sketch_align_filename,self.ui.text_folder_a,self.ui.text_file_a,
                               self.ui.text_folder_b,self.ui.text_file_b,self.ui.text_folder_c,self.ui.text_file_c,self.ui.text_folder_d,self.ui.text_file_d,
                               self.ui.text_align_outfile,
                               # self.ui.text_color_tagnow,
                               # self.ui.text_color_pagenum,
                               # self.ui.text_color_scale,
                               # self.ui.text_color_papersize
                               ]
            self.text_list_clear(text_list_align)
            self.text_list_clear(self.button_group_tag_textlist[1:6])
            self.text_list_clear(self.button_group_tag_textlist[17:27])
            # for box in self.button_group_tag_boxlist:
            #     box.setChecked(False)
            '''Drawing tab reset'''
            text_list_drawing = [self.ui.text_drawing_sktechdir,self.ui.text_drawing_2,self.ui.text_drawing_3,self.ui.text_drawing_4,self.ui.text_issue_type,
                                 self.ui.text_logo_dir,self.ui.text_drawing_scale,self.ui.text_drawing_type,self.ui.text_drawing_status,self.ui.text_drawing_service,
                                 self.ui.text_drawing_output,self.ui.text_update_filepath,self.ui.text_drawing_update_totalpage,self.ui.text_drawing_update_coverpage,
                                 self.ui.text_drawing_update_drawingpage,self.ui.text_other_update_drawingpage,self.ui.text_drawing_update_papersize,
                                 self.ui.text_drawing_update_project,self.ui.text_drawing_update_scale,self.ui.text_drawing_update_prono,self.ui.text_drawing_update_issuetype,]
            self.text_list_clear(text_list_drawing)
            self.text_list_clear(self.rev_list)
            self.text_list_clear(self.description_b_list)
            self.text_list_clear(self.drw_list)
            self.text_list_clear(self.chk_list)
            self.text_list_clear(self.date_list)
            table_list_drawing = [self.ui.table_drawing]
            self.table_list_clear(table_list_drawing)
            button_list_drawing=[self.ui.radiobutton_drawing_approval,self.ui.radiobutton_drawing_construction,self.ui.checkbox_issue_type,
                                 self.ui.checkbox_drawing_movesketch,self.ui.checkbox_plot_folder]
            self.button_list_clear(button_list_drawing)
            for content in self.description_a_list:
                content.setCurrentIndex(0)
            self.ui.label_drawing_update_logo.setText("Logo")
            tags = []
            # for i in range(len(self.button_group_tag_boxlist)):
            #     if self.button_group_tag_boxlist[i].isChecked():
            #         tag=self.button_group_tag_textlist[i].text()
            #         tags.append(tag)
            self.set_draw_table(tags)
            '''Plot tab reset'''
            self.f_button_drawing_plot_fileclear()
            text_list_plot = [self.ui.text_issue_folder,self.ui.text_issue_link,self.ui.text_plot_link_folder]
            self.text_list_clear(text_list_plot)
            table_list_plot = [self.ui.table_drawing_foldercontent,self.ui.table_drawing_linkcontent,self.ui.table_plot_link_and_date]
            self.table_list_clear(table_list_plot)
            '''Overlay tab reset'''
            text_list_overlay = [self.ui.text_overlay_folder1,self.ui.text_overlay_folder2,self.ui.text_overlay_mech_foldername,self.ui.text_overlay_mech_outfolder,
                                 self.ui.text_overlay_mech_filenameadd,self.ui.text_overlay_mech_pages,self.ui.text_overlay_mech_size,self.ui.text_overlay_outfile,
                                 self.ui.text_overlay_outfile_2,self.ui.text_overlay_outfile_3,self.ui.text_overlay_layer,self.ui.text_overlay_markupfrom_file,
                                 self.ui.text_overlay_markupto_file]
            self.text_list_clear(text_list_overlay)
            table_list_overlay = [self.ui.table_overlay_mech,self.ui.table_overlay_service1]
            self.table_list_clear(table_list_overlay)
            button_list_overlay = [self.ui.checkbox_overlay_overlay,self.ui.checkbox_overlay_update]
            self.button_list_clear(button_list_overlay)
        except:
            traceback.print_exc()
        try:
            self.connect_buttons()
            self.connect_texts()
            self.connect_toolbuttons()
            self.connect_button_groups()
        except:
            traceback.print_exc()
    '''======================================Filing tab==========================================='''
    '''Filing - Set color when similar content found'''
    def f_change_filing_tablecolor(self):
        try:
            keyword_list1 = find_keyword(str(self.ui.text_filing_name.text()))
            for row in range(self.ui.table_filling.rowCount()):
                item = self.ui.table_filling.item(row, 0)
                keyword_list2 = find_keyword(item.text())
                if set(keyword_list1).intersection(set(keyword_list2)):
                    self.ui.table_filling.item(row, 0).setBackground(QColor(240, 240, 240))
                else:
                    self.ui.table_filling.item(row, 0).setBackground(QColor(170, 170, 170))
        except:
            traceback.print_exc()
    '''Filing - Reset file table'''
    def set_filing_table(self,filelist):
        try:
            self.ui.table_filling.setRowCount(0)
            for file in filelist:
                row_position = self.ui.table_filling.rowCount()
                self.ui.table_filling.insertRow(row_position)
                self.ui.table_filling.setItem(row_position, 0, QTableWidgetItem(file))
        except:
            traceback.print_exc()
    '''Filing - Get file names'''
    def file_check_filing(self):
        try:
            folder_name =self.ui.text_filing_folder.text()
            file_list = os.listdir(folder_name)
            self.set_filing_table(file_list)
        except:
            pass
    '''Filing - Set new folder'''
    def f_filing_set(self):
        try:
            fn1=self.ui.text_filing_folder.text()
            fn2=self.ui.text_filing_name.text()
            folder_name = os.path.join(fn1, fn2)
            if os.path.exists(folder_name):
                open_folder(folder_name)
            else:
                create_directory(folder_name)
                open_folder(folder_name)
        except:
            traceback.print_exc()
    '''======================================Combine tab==========================================='''
    '''Combine - Combine'''
    def f_combine_combine(self):
        try:
            combinefile_list=self.get_table_files("Combine")
            current_pdf_parent_directory = str(self.ui.text_combine_1_foldername.text())
            output_file_dir = str(self.ui.text_combine_op.text())
            name_extension=str(self.ui.text_combine_3_filenameadd.text())
            combine_pdf_dir=os.path.join(output_file_dir, name_extension)
            self.copy2ss(combine_pdf_dir)
            combinefile_list_ab=[]
            for i in range(len(combinefile_list)):
                combinefile_list_ab.append(os.path.join(current_pdf_parent_directory, combinefile_list[i]))
            for file in combinefile_list_ab:
                if is_pdf_open(file)==True:
                    message('Please close the pdf while combining',self.ui)
                    return
            if is_pdf_open(combine_pdf_dir) == True:
                message('Please close the pdf while combining',self.ui)
                return
            if os.path.exists(combine_pdf_dir):
                overwrite=messagebox('Overwrite',f'{combine_pdf_dir} \n exists, do you want to overwrite?',self.ui)
                if not overwrite:
                    return
            self.progress_dialog = ProgressDialog("Combining, open in Bluebeam when done.",self.ui)
            self.combine_pro = Combine(self.progress_dialog,combinefile_list_ab, combine_pdf_dir)
            self.combine_pro.start()
            self.progress_dialog.exec_()
            flatten_pdf(combine_pdf_dir, combine_pdf_dir)
            self.ui.table_combine_1_filename.setRowCount(0)
            self.ui.text_combine_3_filenameadd.setText('')
            self.ui.text_combine_1_foldername.setText('')
            self.ui.text_combine_op.setText('')
            open_in_bluebeam(combine_pdf_dir)
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''======================================Sketch tab==========================================='''
    '''Tags change in Sketch'''
    def tags_text_change(self,name):
        pass
        # try:
        #     tag_text,tag_box=self.tag_text_dict[name][0],self.tag_text_dict[name][1]
        #     if tag_text.text()!='':
        #         tag_box.setChecked(True)
        #     else:
        #         tag_box.setChecked(False)
        #     tag_num_now = 0
        #     # for i in range(self.button_tag_group_num):
        #     #     if self.button_group_tag.button(i + 1).isChecked():
        #     #         tag_num_now += 1
        #     self.ui.text_color_tagnow.setText(str(tag_num_now))
        # except:
        #     traceback.print_exc()
    '''Sketch - Align and stitch'''
    def f_button_sketch_process(self):
        self.color_process=False
        self.f_sketch_align_align()
        try:
            if self.color_process:
                grayonoff=self.ui.checkbox_sketch_colorprocess_19_grayscale.isChecked()
                tagonoff = self.ui.checkbox_sketch_colorprocess_19_tag.isChecked()
                lumionoff= self.ui.checkbox_luminosity.isChecked()
                changecoloronoff=self.ui.frame_modify_color.isVisible()
                # if tagonoff:
                #     if int(self.ui.text_color_tagnow.text()) != int(self.ui.text_color_pagenum.text()):
                #         message('Tag amount is not correct',self.ui)
                #         return
                if lumionoff or changecoloronoff:
                    self.f_lumi_color()
                if grayonoff or tagonoff:
                    time.sleep(1)
                    self.f_sketch_colorprocess_addtags()
            open_in_bluebeam(self.ui.text_align_outfile.text())
        except:
            traceback.print_exc()
    '''Sketch - File name change'''
    def text_sketch_align_filename_change(self):
        self.pagenum_cal()
        self.f_alignout_file()

    '''Sketch - Lumi and color'''
    def f_lumi_color(self):
        try:
            lumionoff = self.ui.checkbox_luminosity.isChecked()
            changecoloronoff=self.ui.frame_modify_color.isVisible()
            file_sketch = self.ui.text_align_outfile.text()
            pagenum = get_number_of_page(file_sketch)
            lumi=self.ui.text_sketch_lumi.text()
            lumi_set="0.45"
            if is_float(lumi):
                if float(lumi)<=1 and float(lumi)>=0:
                    lumi_set=str(lumi)
            wait_prcocess=f_waitinline('0000')
            if wait_prcocess != 'ok now':
                set = messagebox('Waiting confirm','Someone is ' + wait_prcocess + ', do you want to wait?',self.ui)
                if set:
                    file_time=str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans=os.path.join(conf["trans_dir"],file_time)
                    txt_name=os.path.join(file_dir_trans,'file_names_lumicolor.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    colors_change=[]
                    try:
                        for i_c in range(len(self.color_change)):
                            colors_change.append(self.color_change[i_c].upper())
                            new_colors=adjust_hex_color(self.color_change[i_c],1)
                            colors_change=colors_change+new_colors
                            colors_change= remove_duplicates_keep_order(colors_change)
                    except:
                        pass
                    with open(txt_name, 'w') as file:
                        file_names={'lumionoff':lumionoff,'changecoloronoff':changecoloronoff,'file':file_sketch,'pagenum':pagenum,'color_change':colors_change,'lumi_set':lumi_set}
                        file.write(str(file_names))
                    self.progress_dialog = ProgressDialog('Someone is '+wait_prcocess+', your task is waiting in line.',self.ui)
                    self.waitinline = Waitinline(self.progress_dialog,file_time)
                    self.waitinline.start()
                    self.progress_dialog.exec_()
                else:
                    return
            else:
                file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                txt_name = os.path.join(file_dir_trans, 'file_names_lumicolor.txt')
                os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                colors_change = []
                try:
                    for i_c in range(len(self.color_change)):
                        colors_change.append(self.color_change[i_c].upper())
                        new_colors = adjust_hex_color(self.color_change[i_c], 1)
                        colors_change = colors_change + new_colors
                        colors_change = remove_duplicates_keep_order(colors_change)
                except:
                    pass
                with open(txt_name, 'w') as file:
                    file_names = {'lumionoff': lumionoff, 'changecoloronoff': changecoloronoff, 'file': file_sketch,
                                  'pagenum': pagenum, 'color_change': colors_change, 'lumi_set': lumi_set}
                    file.write(str(file_names))
            self.progress_dialog = ProgressDialog('Modifying color and changing luminosity',self.ui)
            self.lumicolor = Process2PC(self.progress_dialog,["LumiColor",file_dir_trans])
            self.lumicolor.start()
            self.progress_dialog.exec_()
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False
    '''Sketch - Align file name from sketch'''
    def f_alignout_file(self):
        try:

            if self.ui.text_sketch_align_filename.text()[-12:]=='Sketch O.pdf':
                filename=str(Path(self.current_folder_address).name).split('-')[1]
                outputname =os.path.join(self.current_folder_address,filename+'-Mechanical Sketch.pdf')
            else:
                outputname=os.path.join(self.ui.text_sketch_align_filedir.text(),'Aligned.pdf')
            self.ui.text_align_outfile.setText(outputname)
        except:
            traceback.print_exc()
    '''Style change when different sketch changed in modify color'''
    # def f_button_group_modifycolor_sketch(self, button):
    def f_button_group_modifycolor_sketch(self):
        try:
            self.ui.combobox_color_sketch_o.clear()
            for child in self.frame_sourcecolor1.children():
                if isinstance(child, QWidget):
                    self.frame_sourcecolor1.layout().removeWidget(child)
                    child.deleteLater()
            for child in self.frame_sourcecolor2.children():
                if isinstance(child, QWidget):
                    self.frame_sourcecolor2.layout().removeWidget(child)
                    child.deleteLater()
            self.ui.label_color_inpage.setText("")
            self.ui.label_color_outpage.setText("")
            # button_id = self.button_group_modifycolor_sketch.id(button)
            # combobox_list=[self.ui.combobox_color_sketch_o,self.ui.combobox_color_sketch_a,
            #                self.ui.combobox_color_sketch_b,self.ui.combobox_color_sketch_c,
            #                self.ui.combobox_color_sketch_d]
            # if button_id==1:
            file_dir=os.path.join(self.ui.text_sketch_align_filedir.text(),self.ui.text_sketch_align_filename.text())
            # elif button_id == 2:
            #     file_dir = os.path.join(self.ui.text_folder_a.text(), self.ui.text_file_a.text())
            # elif button_id == 3:
            #     file_dir = os.path.join(self.ui.text_folder_b.text(), self.ui.text_file_b.text())
            # elif button_id == 4:
            #     file_dir = os.path.join(self.ui.text_folder_c.text(), self.ui.text_file_c.text())
            # elif button_id == 5:
            #     file_dir = os.path.join(self.ui.text_folder_d.text(), self.ui.text_file_d.text())
            pages = get_number_of_page(file_dir)
            # for box in combobox_list:
            #     box.setEnabled(False)
            #     box.setStyleSheet("background-color: rgb(200, 200, 200);")
            # combobox_list[button_id-1].setEnabled(True)
            # combobox_list[button_id - 1].setStyleSheet("background-color: white;")
            # for i in range(pages):
            #     new_item_text='Page '+str(i+1)
            #     combobox_list[button_id-1].addItem(new_item_text)
        except:
            traceback.print_exc()
    '''Sketch modify color svg load'''
    def load_svg(self,num):
        try:
            size = QSize(400, 300)
            pixmap = QPixmap(size)
            pixmap.fill(Qt.white)
            painter = QPainter(pixmap)
            if num==1:
                self.renderer1.render(painter)
                painter.end()
                self.ui.label_color_inpage.setPixmap(pixmap)
            else:
                self.renderer2.render(painter)
                painter.end()
                self.ui.label_color_outpage.setPixmap(pixmap)
        except:
            traceback.print_exc()
    '''Generate color toolbutton when modify color'''
    def add_tool_button(self,color,name1,name2,num):
        try:
            tool_button = QToolButton()
            tool_button.setObjectName(name1)
            tool_button.setStyleSheet(f"background-color: {color};width:20;height:20;")
            self.layout1.addWidget(tool_button)
            self.tool_button1_list.append(tool_button)
            tool_button2 = QToolButton()
            tool_button2.setObjectName(name2)
            tool_button2.setStyleSheet(f"background-color: {color};width:20;height:20;")
            self.layout2.addWidget(tool_button2)
            self.tool_button2_list.append(tool_button2)
            tool_button2.clicked.connect(self.changecolor(name2,num))
        except:
            traceback.print_exc()
    '''Change color toolbox when clicked'''
    def changecolor(self, name,num):
        def callback():
            try:
                background_color = self.ui.findChild(QToolButton, name).palette().color(self.ui.findChild(QToolButton, name).backgroundRole())
                background_color_string = f'rgb({background_color.red()}, {background_color.green()}, {background_color.blue()})'
                if background_color_string=='rgb(255, 255, 255)':
                    self.ui.findChild(QToolButton, name).setText('')
                    self.ui.findChild(QToolButton, name).setStyleSheet(f'background-color: {self.svg_colors[num]}; width:20;height:20;')
                    self.color_onoff[num]=False
                else:
                    self.ui.findChild(QToolButton, name).setText('/')
                    self.ui.findChild(QToolButton, name).setStyleSheet(f'background-color: rgb(255,255,255); width:20;height:20;')
                    self.color_onoff[num] = True
                self.f_change_svg_color()
            except:
                traceback.print_exc()
        return callback
    '''Define color toolbutton when modify color'''
    def add_color_button(self, colors):
        try:
            self.tool_button1_list = []
            self.tool_button2_list = []
            for i in range(len(colors)):
                self.add_tool_button(colors[i],'button_colorpool_'+str(i),'button_colorpool2_'+str(i),i)
        except:
            traceback.print_exc()
    '''Change svg color in modify color'''
    def f_change_svg_color(self):
        try:
            colors=[]
            self.color_change = []
            for i in range(len(self.svg_colors)):
                if self.color_onoff[i]==True:
                    colors.append(self.svg_colors[i])
                    self.color_change.append(self.svg_colors[i])
            color_origin_pic = os.path.join(conf["c_temp_dir"], "color_origin_pic.svg")
            color_changed_pic = os.path.join(conf["c_temp_dir"], "color_changed_pic.svg")
            if len(colors)!=0:
                shutil.copy(color_origin_pic, color_changed_pic)
                svg_set_opacity_multicolor(color_changed_pic, colors, color_changed_pic)
            else:
                shutil.copy(color_origin_pic,color_changed_pic)
            self.renderer2.load(color_changed_pic)
            self.load_svg(2)
        except:
            traceback.print_exc()
    '''Set svg page in modify color when page changes'''
    def f_combobox_color_1(self):
        try:
            for child in self.frame_sourcecolor1.children():
                if isinstance(child, QWidget):
                    self.frame_sourcecolor1.layout().removeWidget(child)
                    child.deleteLater()
            for child in self.frame_sourcecolor2.children():
                if isinstance(child, QWidget):
                    self.frame_sourcecolor2.layout().removeWidget(child)
                    child.deleteLater()
            selected_text = self.ui.combobox_color_sketch_o.currentText()
            page=int(selected_text[5:])
            file_dir=os.path.join(self.ui.text_sketch_align_filedir.text(),self.ui.text_sketch_align_filename.text())
            color_origin_pic = os.path.join(conf["c_temp_dir"], "color_origin_pic.svg")
            with open(color_origin_pic, 'wb') as f:
                PDFTools_v2.pdf_page_to_svg(file_dir, page-1, f)
            self.renderer1.load(color_origin_pic)
            self.load_svg(1)
            colors=list(PDFTools_v2.pdf_color(file_dir, page-1))
            self.color_onoff=[]
            for k in range(len(colors)):
                self.color_onoff.append(False)
            self.svg_colors = colors
            self.add_color_button(colors)
            self.color_change=[]
        except:
            traceback.print_exc()
    '''Sketch - Count tags when clicked'''
    def buttongroup_tag_click(self):
        try:
            tag_num_now=0
            # for i in range(self.button_tag_group_num):
            #     if self.button_group_tag.button(i+1).isChecked():
            #         tag_num_now+=1
            # self.ui.text_color_tagnow.setText(str(tag_num_now))
        except:
            traceback.print_exc()
    '''Sketch - Add tags'''
    def f_sketch_colorprocess_addtags(self):
        try:
            gray_onoff=self.ui.checkbox_sketch_colorprocess_19_grayscale.isChecked()
            tag_onoff = self.ui.checkbox_sketch_colorprocess_19_tag.isChecked()
            tags = []
            # for i in range(len(self.button_group_tag_boxlist)):
            #     if self.button_group_tag_boxlist[i].isChecked():
            #         tag = self.button_group_tag_textlist[i].text()
            #         tags.append(tag)
            align_file=self.ui.text_align_outfile.text()
            # output_scale=self.ui.text_color_scale.text().split(':')[1]
            # size_output=self.ui.text_color_papersize.text()
            # cus4=output_scale
            # cus5=0
            # cus6=0
            # if size_output.find(':')!=-1:
            #     cus5 = size_output.split(':')[0]
            #     cus6 = size_output.split(':')[1]
            #     size_output='custom'
            # tag_size_dict = {'output_scale': output_scale, 'output_size': size_output, 'cus_list': [cus4, cus5, cus6]}
            # if is_pdf_open(align_file) == True:
            #     message('Please close the pdf while coloring',self.ui)
            #     return
            # self.copy2ss(align_file)
            # self.progress_dialog = ProgressDialog("Adding tags, open in Bluebeam when done.",self.ui)
            # self.addtag = Addtag(self.progress_dialog, align_file, gray_onoff, tags,tag_size_dict,tag_onoff)
            # self.addtag.start()
            # self.progress_dialog.exec_()
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Sketch - Page number calculation'''
    def pagenum_cal(self):
        try:
            print('cal')
            file_all = os.path.join(self.ui.text_sketch_align_filedir.text(), self.ui.text_sketch_align_filename.text())
            print(file_all)
            if file_exists(file_all):
                page_num_all = get_number_of_page(file_all)
                print(page_num_all)
                # self.ui.text_color_pagenum.setText(str(page_num_all))
        except:
            traceback.print_exc()
    '''Sketch output size change'''
    def buttongroup4_clicked(self,button):
        try:
            button_id = self.button_group4.id(button)
            input_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24', 'A0-26']
            if button_id!=8:
                self.textset_with_blocksignal(self.ui.text_sketch_rescale_2_outputcussize1)
                self.textset_with_blocksignal(self.ui.text_sketch_rescale_2_outputcussize2)
                papersize=input_size_list[button_id-1]
                # self.ui.text_color_papersize.setText(papersize)
        except:
            traceback.print_exc()
    '''Sketch - Rescale'''
    def test_sketch_rescale(self):
        # input_file_dir = self.ui.text_sketch_rescale_1_filedir.text()
        # output_file_dir = self.ui.text_sketch_rescale_1_filesavedir.text()
        # shutil.copytree(input_file_dir, output_file_dir)
        # self.ui.text_sketch_rescale_1_filedir_2.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number, get_timestamp()+"-"+self.project_name+"-aligned.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_2.setText(path)
        # self.ui.text_sketch_rescale_1_filedir_3.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number, get_timestamp()+"-"+self.project_name+"-color_processed.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_3.setText(path)
        # self.ui.text_sketch_rescale_1_filedir_4.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number, get_timestamp() + "-" + self.project_name + "-markup_coped.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_4.setText(path)
        # self.ui.text_sketch_rescale_1_filedir_5.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number,get_timestamp() + "-" + self.project_name + "-erased.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_5.setText(path)
        self.ui.toolBox.setCurrentIndex(1)

    def test_sketch_align(self):
        # input_file_dir = self.ui.text_sketch_rescale_1_filedir_2.text()
        # output_file_dir = self.ui.text_sketch_rescale_1_filesavedir_2.text()
        # # shutil.copytree(input_file_dir, output_file_dir)
        # self.ui.text_sketch_rescale_1_filedir_3.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number, get_timestamp()+"-"+self.project_name+"-color_processed.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_3.setText(path)
        # self.ui.text_sketch_rescale_1_filedir_4.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number, get_timestamp() + "-" + self.project_name + "-markup_coped.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_4.setText(path)
        # self.ui.text_sketch_rescale_1_filedir_5.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number,get_timestamp() + "-" + self.project_name + "-erased.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_5.setText(path)
        self.ui.toolBox.setCurrentIndex(2)

    def test_sketch_color_process(self):
        # input_file_dir = self.ui.text_sketch_rescale_1_filedir_3.text()
        # output_file_dir = self.ui.text_sketch_rescale_1_filesavedir_3.text()
        # shutil.copytree(input_file_dir, output_file_dir)
        # self.ui.text_sketch_rescale_1_filedir_4.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number, get_timestamp() + "-" + self.project_name + "-markup_coped.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_4.setText(path)
        # self.ui.text_sketch_rescale_1_filedir_5.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number,get_timestamp() + "-" + self.project_name + "-erased.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_5.setText(path)
        self.ui.toolBox.setCurrentIndex(3)

    def test_sketch_copy_markup(self):
        # input_file_dir = self.ui.text_sketch_rescale_1_filedir_4.text()
        # output_file_dir = self.ui.text_sketch_rescale_1_filesavedir_4.text()
        # shutil.copytree(input_file_dir, output_file_dir)
        # self.ui.text_sketch_rescale_1_filedir_5.setText(output_file_dir)
        # path = os.path.join(self.backup_disk, self.quotation_number,get_timestamp() + "-" + self.project_name + "-erased.pdf")
        # self.ui.text_sketch_rescale_1_filesavedir_5.setText(path)
        self.ui.toolBox.setCurrentIndex(4)

    def test_sketch_erase(self):
        return
        # input_file_dir = self.ui.text_sketch_rescale_1_filedir_5.text()
        # output_file_dir = self.ui.text_sketch_rescale_1_filesavedir_5.text()
        # shutil.copytree(input_file_dir, output_file_dir)
    def f_sketch_rescale_process(self):
        try:
            self.sketch_rescale_function()
            # if self.ui.frame_rescale1.isVisible()==True:
            #     self.f_rescale_i('a')
            # if self.ui.frame_rescale2.isVisible()==True:
            #     self.f_rescale_i('b')
            # if self.ui.frame_rescale3.isVisible()==True:
            #     self.f_rescale_i('c')
            # if self.ui.frame_rescale4.isVisible()==True:
            #     self.f_rescale_i('d')
            self.ui.toolBox.setCurrentIndex(1)
        except:
            traceback.print_exc()
    '''Sketch - Align sketch O process'''
    def f_sketch_align_align_1(self,filename):
        try:
            self.progress_dialog = ProgressDialog("Aligning Sketch O, open in Bluebeam when done.",self.ui)
            self.align = Align(self.progress_dialog, filename)
            self.align.start()
            self.progress_dialog.exec_()
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False
    '''Sketch - Align sketch A~D process'''
    def f_align_i(self,file):
        try:
            self.progress_dialog = ProgressDialog("Aligning Sketch A~D, open in Bluebeam when done.",self.ui)
            self.align = Align(self.progress_dialog, file)
            self.align.start()
            self.progress_dialog.exec_()
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False
    '''Sketch - Align - Align O~D'''
    def f_sketch_align_align(self):
        try:
            align_name1=os.path.join(self.ui.text_sketch_align_filedir.text(),self.ui.text_sketch_align_filename.text())
            if file_exists(align_name1) == False:
                message('Input file does not exist',self.ui)
                return
            combine_list = [align_name1]
            if is_pdf_open(align_name1) == True:
                message('Please close the pdf while combining',self.ui)
                return
            self.f_sketch_align_align_1(align_name1)
            frame_list=[self.ui.frame_align_a,self.ui.frame_align_b,self.ui.frame_align_c,self.ui.frame_align_d]
            folder_list=[self.ui.text_folder_a.text(),self.ui.text_folder_b.text(),self.ui.text_folder_c.text(),self.ui.text_folder_d.text()]
            file_list=[self.ui.text_file_a.text(),self.ui.text_file_b.text(),self.ui.text_file_c.text(),self.ui.text_file_d.text()]
            for i in range(len(frame_list)):
                if frame_list[i].isVisible() == True:
                    align_name_i=os.path.join(folder_list[i],file_list[i])
                    if is_pdf_open(align_name_i) == True:
                        message('Please close the pdf while combining',self.ui)
                        return
                    self.f_align_i(align_name_i)
                    combine_list.append(align_name_i)
            outputname=self.ui.text_align_outfile.text()
            self.copy2ss(outputname)
            if is_pdf_open(outputname) == True:
                message('Please close the pdf while combining',self.ui)
                return
            if len(combine_list)>1:
                page_file1=get_number_of_page(combine_list[0])
                page_file2 = get_number_of_page(combine_list[1])
                if page_file1!=page_file2:
                    message('Sketch O and sketch A must have same pages.',self.ui)
                    return
                wait_prcocess = f_waitinline('0000')
                if wait_prcocess != 'ok now':
                    set = messagebox('Waiting confirm','Someone is ' + wait_prcocess + ', do you want to wait?',self.ui)
                    if set:
                        file_time=str(datetime.now().strftime("%Y%m%d%H%M%S"))
                        file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                        txt_name = os.path.join(file_dir_trans, 'file_names_overlay.txt')
                        os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                        with open(txt_name, 'w') as file:
                            file_contetnt = {'File1': combine_list[0], 'File2': combine_list[1], 'File3': outputname,
                                             'File4': '', 'Color': None, 'Name': 'Block'}
                            file.write(str(file_contetnt))
                        self.progress_dialog = ProgressDialog('Someone is ' + wait_prcocess + ', your task is waiting in line.',self.ui)
                        self.waitinline = Waitinline(self.progress_dialog, file_time)
                        self.waitinline.start()
                        self.progress_dialog.exec_()
                    else:
                        return
                else:
                    file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                    txt_name = os.path.join(file_dir_trans, 'file_names_overlay.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    with open(txt_name, 'w') as file:
                        file_contetnt = {'File1': combine_list[0], 'File2': combine_list[1], 'File3': outputname,
                                         'File4': '', 'Color': None, 'Name': 'Block'}
                        file.write(str(file_contetnt))
                self.progress_dialog = ProgressDialog("Stitching sketches, open in Bluebeam when done.",self.ui)
                self.align_combine=Process2PC(self.progress_dialog,["Align_combine",file_dir_trans])
                self.align_combine.start()
                self.progress_dialog.exec_()
            else:
                shutil.copy(align_name1, outputname)
            time.sleep(2)
            flatten_pdf(outputname,outputname)
            self.color_process = True
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Sketch - Rescale skecth A~D (to do 1)'''
    def f_rescale_i(self,id):
        try:
            '''get file list and dir'''
            if id=='a':
                file_list=self.get_table_files("Sketch A")
                current_pdf_parent_directory = str(self.ui.text_rescale_cf_1.text())
                output_folder=str(self.ui.text_rescale_of_1.text())
                output_file_name=str(self.ui.text_rescale_outfile_1.text())
                id1 = self.button_group_a1.checkedId()
                id2 = self.button_group_a2.checkedId()
                cus1 = str(self.ui.text_rescale_a.text())
                cus2 = str(self.ui.text_resize1_a.text())
                cus3 = str(self.ui.text_resize2_a.text())
            elif id=='b':
                file_list = self.get_table_files("Sketch B")
                current_pdf_parent_directory = str(self.ui.text_rescale_cf_2.text())
                output_folder = str(self.ui.text_rescale_of_2.text())
                output_file_name = str(self.ui.text_rescale_outfile_2.text())
                id1 = self.button_group_b1.checkedId()
                id2 = self.button_group_b2.checkedId()
                cus1 = str(self.ui.text_rescale_b.text())
                cus2 = str(self.ui.text_resize1_b.text())
                cus3 = str(self.ui.text_resize2_b.text())
            elif id=='c':
                file_list = self.get_table_files("Sketch C")
                current_pdf_parent_directory = str(self.ui.text_rescale_cf_3.text())
                output_folder = str(self.ui.text_rescale_of_3.text())
                output_file_name = str(self.ui.text_rescale_outfile_3.text())
                id1 = self.button_group_c1.checkedId()
                id2 = self.button_group_c2.checkedId()
                cus1 = str(self.ui.text_rescale_c.text())
                cus2 = str(self.ui.text_resize1_c.text())
                cus3 = str(self.ui.text_resize2_c.text())
            elif id == 'd':
                file_list = self.get_table_files("Sketch D")
                current_pdf_parent_directory = str(self.ui.text_rescale_cf_4.text())
                output_folder = str(self.ui.text_rescale_of_4.text())
                output_file_name = str(self.ui.text_rescale_outfile_4.text())
                id1 = self.button_group_d1.checkedId()
                id2 = self.button_group_d2.checkedId()
                cus1 = str(self.ui.text_rescale_d.text())
                cus2 = str(self.ui.text_resize1_d.text())
                cus3 = str(self.ui.text_resize2_d.text())
            file_list_ab=[]
            for i in range(len(file_list)):
                file_list_ab.append(os.path.join(current_pdf_parent_directory, file_list[i]))
            '''get group selections'''
            id3 = self.button_group3.checkedId()
            id4 = self.button_group4.checkedId()
            id_list=[id1,id2,id3,id4]
            cus4=str(self.ui.text_sketch_rescale_2_oscx.text())
            cus5=str(self.ui.text_sketch_rescale_2_outputcussize1.text())
            cus6=str(self.ui.text_sketch_rescale_2_outputcussize2.text())
            cus_list=[cus1,cus2,cus3,cus4,cus5,cus6]
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            '''confirm right content'''
            if not all([f.endswith(".pdf") or f.endswith(".PDF") for f in file_list_ab]):
                message('Please input at least one pdf file in the dialog box',self.ui)
                return
            output_file = os.path.join(output_folder,output_file_name)
            if file_exists(output_file):
                ss_name = os.path.join(Path(output_file).parent.absolute(),
                                       date.today().strftime("%Y%m%d") +'-'+ Path(output_file).name)
                if file_exists(ss_name):
                    for i in range(100):
                        new_ss_name = ss_name[:-4] + '(' + str(i+1) + ').pdf'
                        if file_exists(new_ss_name) == False:
                            ss_name = new_ss_name
                            break
                ss_dir = output_folder
                os.makedirs(ss_dir, exist_ok=True)
                shutil.copy(output_file, ss_name)

            if len(file_list_ab) == 1:
                input_file = file_list_ab[0]
                if is_pdf_open(input_file) == True:
                    message('Please close the pdf while rescaling',self.ui)
                    return
                if id == 'a':
                    self.ui.tableWidget_a.setRowCount(0)
                elif id == 'b':
                    self.ui.tableWidget_b.setRowCount(0)
                elif id == 'c':
                    self.ui.tableWidget_c.setRowCount(0)
                elif id == 'd':
                    self.ui.tableWidget_d.setRowCount(0)
            else:
                combine_pdf_dir =output_file
                for file in file_list_ab:
                    if is_pdf_open(file) == True:
                        message('Please close the pdf while rescaling',self.ui)
                        return
                if is_pdf_open(combine_pdf_dir) == True:
                    message('Please close the pdf while rescaling',self.ui)
                    return
                combine_pdf(file_list_ab, combine_pdf_dir)
                input_file = combine_pdf_dir
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            '''get process size change and scale change'''
            input_scale_list = ['50', '75', '100', '150', '200', 'custom']
            input_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24','A0-26', 'custom']
            output_scale_list = ['50', '100', 'custom']
            output_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24','A0-26', 'custom']
            if id_list[0] == 6:
                input_scale = cus_list[0]
            else:
                input_scale = input_scale_list[id_list[0]-1]
            if id_list[2] == 3:
                output_scale = cus_list[3]
            else:
                output_scale = output_scale_list[id_list[2]-1]
            page_size_json_dir = os.path.join(conf["database_dir"], "page_size.json")
            all_page_size = json.load(open(page_size_json_dir))
            if id_list[1] == 8:
                size_input='custom'
                if not (is_float(cus_list[1]) and is_float(cus_list[2])):
                    message('Please enter correct input size width and height',self.ui)
                    return
                input_size_x=convert_mm_to_pixel(float(cus_list[1]))
                input_size_y=convert_mm_to_pixel(float(cus_list[2]))
            else:
                size_input=input_size_list[id_list[1]-1]
                input_size_x = convert_mm_to_pixel(all_page_size[size_input]["width"])
                input_size_y = convert_mm_to_pixel(all_page_size[size_input]["height"])
            if id_list[3] == 8:
                size_output='custom'
                if not (is_float(cus_list[4]) and is_float(cus_list[5])):
                    message('Please enter correct input size width and height',self.ui)
                    return
                output_size_x=convert_mm_to_pixel(float(cus_list[4]))
                output_size_y=convert_mm_to_pixel(float(cus_list[5]))

            else:
                size_output=output_size_list[id_list[3]-1]
                output_size_x = convert_mm_to_pixel(all_page_size[size_output]["width"])
                output_size_y = convert_mm_to_pixel(all_page_size[size_output]["height"])
            self.align_files.append(output_file)

        except Exception as e:
            traceback.print_exc()
            raise e


        try:
            size_list=[input_scale, input_size_x, input_size_y,output_scale, output_size_x, output_size_y]
            if is_pdf_open(output_file) == True:
                message('Please close the pdf while rescaling',self.ui)
                return
            self.tag_size_dict={'output_scale':output_scale,'output_size':size_output,'cus_list':cus_list[3:]}
            self.progress_dialog = ProgressDialog("Rescaling, open in Bluebeam when done.",self.ui)
            self.process = Process(self.progress_dialog, input_file,size_list,output_file)
            self.process.start()
            self.progress_dialog.exec_()
            if id=='a':
                self.ui.tableWidget_a.setRowCount(0)
                folder_dir = self.ui.text_rescale_of_1.text()
                file_name = self.ui.text_rescale_outfile_1.text()
                self.ui.text_rescale_cf_1.setText('')
                self.ui.text_rescale_of_1.setText('')
                self.ui.text_rescale_outfile_1.setText('')
                self.ui.text_folder_a.setText(folder_dir)
                self.ui.text_file_a.setText(file_name)
                self.ui.frame_align_a.setVisible(True)
                self.ui.toolbutton_align_a.setText("Block A -")
            elif id=='b':
                self.ui.tableWidget_b.setRowCount(0)
                folder_dir = self.ui.text_rescale_of_2.text()
                file_name = self.ui.text_rescale_outfile_2.text()
                self.ui.text_rescale_cf_2.setText('')
                self.ui.text_rescale_of_2.setText('')
                self.ui.text_rescale_outfile_2.setText('')
                self.ui.text_folder_b.setText(folder_dir)
                self.ui.text_file_b.setText(file_name)
                self.ui.frame_align_b.setVisible(True)
                self.ui.toolbutton_align_b.setText("Block B -")
            elif id=='c':
                self.ui.tableWidget_c.setRowCount(0)
                folder_dir = self.ui.text_rescale_of_3.text()
                file_name = self.ui.text_rescale_outfile_3.text()
                self.ui.text_rescale_cf_3.setText('')
                self.ui.text_rescale_of_3.setText('')
                self.ui.text_rescale_outfile_3.setText('')
                self.ui.text_folder_c.setText(folder_dir)
                self.ui.text_file_c.setText(file_name)
                self.ui.frame_align_c.setVisible(True)
                self.ui.toolbutton_align_c.setText("Block C -")
            elif id=='d':
                self.ui.tableWidget_d.setRowCount(0)
                folder_dir = self.ui.text_rescale_of_4.text()
                file_name = self.ui.text_rescale_outfile_4.text()
                self.ui.text_rescale_cf_4.setText('')
                self.ui.text_rescale_of_4.setText('')
                self.ui.text_rescale_outfile_4.setText('')
                self.ui.text_folder_d.setText(folder_dir)
                self.ui.text_file_d.setText(file_name)
                self.ui.frame_align_d.setVisible(True)
                self.ui.toolbutton_align_d.setText("Block D -")
            self.align_file_i=output_file
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False
    '''Sketch - Rescale skecth O (to do 1)'''
    def sketch_rescale_function(self):
        input_file_dir = self.ui.text_sketch_rescale_1_filedir.text()
        output_file_dir = self.ui.text_sketch_rescale_1_filesavedir.text()
        input_scale_list = ['50', '75', '100', '150', '200', 'custom']
        input_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24', 'A0-26', 'custom']
        output_scale_list = ['50', '100', 'custom']
        output_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24', 'A0-26', 'custom']
        # id1 = self.button_group1.checkedId()
        id1 = 0
        id2 = self.button_group2.checkedId()
        id3 = self.button_group3.checkedId()
        id4 = self.button_group4.checkedId()
        id_list = [id1, id2, id3, id4]
        cus1 = str(self.ui.text_sketch_rescale_1_oscx.text())
        cus2 = str(self.ui.text_sketch_rescale_7_outputcussize1.text())
        cus3 = str(self.ui.text_sketch_rescale_8_outputcussize2.text())
        cus4 = str(self.ui.text_sketch_rescale_2_oscx.text())
        cus5 = str(self.ui.text_sketch_rescale_2_outputcussize1.text())
        cus6 = str(self.ui.text_sketch_rescale_2_outputcussize2.text())
        cus_list = [cus1, cus2, cus3, cus4, cus5, cus6]
        page_size_json_dir = os.path.join(conf["database_dir"], "page_size.json")
        all_page_size = json.load(open(page_size_json_dir))
        if id_list[0] == 6:
            input_scale = cus_list[0]
        else:
            input_scale = input_scale_list[id_list[0] - 1]
        if id_list[2] == 3:
            output_scale = cus_list[3]
        else:
            output_scale = output_scale_list[id_list[2] - 1]
        if id_list[1] == 8:
            size_input = 'custom'
            if not (is_float(cus_list[1]) and is_float(cus_list[2])):
                message('Please enter correct input size width and height', self.ui)
                return
            input_size_x = convert_mm_to_pixel(float(cus_list[1]))
            input_size_y = convert_mm_to_pixel(float(cus_list[2]))

        else:
            size_input = input_size_list[id_list[1] - 1]
            input_size_x = convert_mm_to_pixel(all_page_size[size_input]["width"])
            input_size_y = convert_mm_to_pixel(all_page_size[size_input]["height"])

        if id_list[3] == 8:
            size_output = 'custom'
            if not (is_float(cus_list[4]) and is_float(cus_list[5])):
                message('Please enter correct input size width and height', self.ui)
                return
            output_size_x = convert_mm_to_pixel(float(cus_list[4]))
            output_size_y = convert_mm_to_pixel(float(cus_list[5]))

        else:
            size_output = output_size_list[id_list[3] - 1]
            output_size_x = convert_mm_to_pixel(all_page_size[size_output]["width"])
            output_size_y = convert_mm_to_pixel(all_page_size[size_output]["height"])
        if file_exists(output_file_dir):
            ss_name = os.path.join(Path(output_file_dir ).parent.absolute(),
                                   date.today().strftime("%Y%m%d") + '-' + Path(output_file_dir).name)
            if file_exists(ss_name):
                for i in range(100):
                    new_ss_name = ss_name[:-4] + '(' + str(i + 1) + ').pdf'
                    if file_exists(new_ss_name) == False:
                        ss_name = new_ss_name
                        break
            ss_dir = str(self.ui.text_sketch_rescale_1_filesavedir.text())
            os.makedirs(ss_dir, exist_ok=True)
            shutil.copy(output_file_dir, ss_name)

        try:
            size_list = [input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y]
            output_file = Path(self.ui.text_sketch_rescale_1_filesavedir.text())
            # self.ui.text_sketch_rescale_2_filename.text())
            if is_pdf_open(output_file) == True:
                message('Please close the pdf while rescaling', self.ui)
                return
            self.tag_size_dict = {'output_scale': output_scale, 'output_size': size_output, 'cus_list': cus_list[3:]}
            self.progress_dialog = ProgressDialog("Rescaling, open in Bluebeam when done.", self.ui)
            self.process = Process(self.progress_dialog, input_file_dir, size_list, output_file)
            self.process.start()
            self.progress_dialog.exec_()
            # self.ui.table_sketch_rescale_1_filename.setRowCount(0)
            folder_dir = self.ui.text_sketch_rescale_1_filesavedir.text()
            # file_name=self.ui.text_sketch_rescale_2_filename.text()
            self.ui.text_sketch_rescale_1_filedir.setText('')
            self.ui.text_sketch_rescale_1_filesavedir.setText('')
            # self.ui.text_sketch_rescale_2_filename.setText('')
            self.align_file = output_file
            '''set align and color page file name'''
            self.ui.text_sketch_rescale_1_filedir_2.setText(output_file_dir)
            output_file_name = os.path.join(self.current_folder_address, self.project_name+"-Mechanical Sketch.pdf")
            self.ui.text_sketch_rescale_1_filesavedir_2.setText(output_file_name)
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False
            print()
    def f_sketch_rescale_process_1(self):
        try:
            '''get file list and dir'''
            file_list=self.get_table_files("Sketch O")
            current_pdf_parent_directory = str(self.ui.text_sketch_rescale_1_filedir.text())
            file_list_ab=[]
            for i in range(len(file_list)):
                file_list_ab.append(os.path.join(current_pdf_parent_directory, file_list[i]))
            '''get quo_no and group selections'''
            quo_no = str(self.ui.text_all_1_projectnumber.text())
            # id1 = self.button_group1.checkedId()
            id1 = 0
            id2 = self.button_group2.checkedId()
            id3 = self.button_group3.checkedId()
            id4 = self.button_group4.checkedId()
            id_list=[id1,id2,id3,id4]
            cus1=str(self.ui.text_sketch_rescale_1_oscx.text())
            cus2=str(self.ui.text_sketch_rescale_7_outputcussize1.text())
            cus3=str(self.ui.text_sketch_rescale_8_outputcussize2.text())
            cus4=str(self.ui.text_sketch_rescale_2_oscx.text())
            cus5=str(self.ui.text_sketch_rescale_2_outputcussize1.text())
            cus6=str(self.ui.text_sketch_rescale_2_outputcussize2.text())
            cus_list=[cus1,cus2,cus3,cus4,cus5,cus6]
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            '''confirm right content'''
            if quo_no == '':
                message('Please Enter a project before resize',self.ui)
                return
            if not all([f.endswith(".pdf") or f.endswith(".PDF") for f in file_list_ab]):
                message('Please input at least one pdf file in the dialog box',self.ui)
                return
            output_file = Path(self.ui.text_sketch_rescale_1_filesavedir.text())
                                       # self.ui.text_sketch_rescale_2_filename.text())
            if file_exists(output_file):
                ss_name = os.path.join(Path(output_file).parent.absolute(),
                                       date.today().strftime("%Y%m%d") +'-'+ Path(output_file).name)
                if file_exists(ss_name):
                    for i in range(100):
                        new_ss_name = ss_name[:-4] + '(' + str(i+1) + ').pdf'
                        if file_exists(new_ss_name) == False:
                            ss_name = new_ss_name
                            break


                ss_dir = str(self.ui.text_sketch_rescale_1_filesavedir.text())
                os.makedirs(ss_dir, exist_ok=True)
                shutil.copy(output_file, ss_name)

            if len(file_list_ab) == 1:
                input_file = file_list_ab[0]
                # self.ui.table_sketch_rescale_1_filename.setRowCount(0)
                if is_pdf_open(input_file) == True:
                    message('Please close the pdf while rescaling',self.ui)
                    return
            else:
                combine_pdf_dir =output_file
                for file in file_list_ab:
                    if is_pdf_open(file) == True or is_pdf_open(combine_pdf_dir) == True:
                        message('Please close the pdf while rescaling',self.ui)
                        return
                combine_pdf(file_list_ab, combine_pdf_dir)
                input_file = combine_pdf_dir
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            '''get process size change and scale change'''
            input_scale_list = ['50', '75', '100', '150', '200', 'custom']
            input_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24','A0-26', 'custom']
            output_scale_list = ['50', '100', 'custom']
            output_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24','A0-26', 'custom']
            if id_list[0] == 6:
                input_scale = cus_list[0]
            else:
                input_scale = input_scale_list[id_list[0]-1]
            if id_list[2] == 3:
                output_scale = cus_list[3]
            else:
                output_scale = output_scale_list[id_list[2]-1]
            page_size_json_dir = os.path.join(conf["database_dir"], "page_size.json")
            all_page_size = json.load(open(page_size_json_dir))
            if id_list[1] == 8:
                size_input='custom'
                if not (is_float(cus_list[1]) and is_float(cus_list[2])):
                    message('Please enter correct input size width and height',self.ui)
                    return
                input_size_x=convert_mm_to_pixel(float(cus_list[1]))
                input_size_y=convert_mm_to_pixel(float(cus_list[2]))

            else:
                size_input=input_size_list[id_list[1]-1]
                input_size_x = convert_mm_to_pixel(all_page_size[size_input]["width"])
                input_size_y = convert_mm_to_pixel(all_page_size[size_input]["height"])

            if id_list[3] == 8:
                size_output='custom'
                if not (is_float(cus_list[4]) and is_float(cus_list[5])):
                    message('Please enter correct input size width and height',self.ui)
                    return
                output_size_x=convert_mm_to_pixel(float(cus_list[4]))
                output_size_y=convert_mm_to_pixel(float(cus_list[5]))

            else:
                size_output=output_size_list[id_list[3]-1]
                output_size_x = convert_mm_to_pixel(all_page_size[size_output]["width"])
                output_size_y = convert_mm_to_pixel(all_page_size[size_output]["height"])
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            size_list=[input_scale, input_size_x, input_size_y,output_scale, output_size_x, output_size_y]
            output_file = Path(self.ui.text_sketch_rescale_1_filesavedir.text())
                                       # self.ui.text_sketch_rescale_2_filename.text())
            if is_pdf_open(output_file) == True:
                message('Please close the pdf while rescaling',self.ui)
                return
            self.tag_size_dict={'output_scale':output_scale,'output_size':size_output,'cus_list':cus_list[3:]}
            self.progress_dialog = ProgressDialog("Rescaling, open in Bluebeam when done.",self.ui)
            self.process = Process(self.progress_dialog, input_file,size_list,output_file)
            self.process.start()
            self.progress_dialog.exec_()
            # self.ui.table_sketch_rescale_1_filename.setRowCount(0)
            folder_dir=self.ui.text_sketch_rescale_1_filesavedir.text()
            # file_name=self.ui.text_sketch_rescale_2_filename.text()
            self.ui.text_sketch_rescale_1_filedir.setText('')
            self.ui.text_sketch_rescale_1_filesavedir.setText('')
            # self.ui.text_sketch_rescale_2_filename.setText('')
            self.align_file=output_file
            '''set align and color page file name'''
            self.ui.text_sketch_align_filedir.setText(folder_dir)
            # self.ui.text_sketch_align_filename.setText(file_name)
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False
    '''======================================Drawing tab==========================================='''
    '''Set tags in Drawing from Sketch'''
    def drawing_text_change_1(self):
        return
        # try:
        #     if self.ui.text_color_tagnow.text()==self.ui.text_drawing_2.text():
        #         tags = []
        #         # for i in range(len(self.button_group_tag_boxlist)):
        #         #     if self.button_group_tag_boxlist[i].isChecked():
        #         #         tag=self.button_group_tag_textlist[i].text()
        #         #         tags.append(tag)
        #         self.set_draw_table(tags)
        # except:
        #     traceback.print_exc()
    '''Set drawing info in Drawing from Sketch'''
    def drawing_text_change_all(self,name):
        return
        # try:
        #     text_name,set_text_name=self.drawing_text_change_dict[name][0],self.drawing_text_change_dict[name][1]
        #     text=text_name.text()
        #     set_text_name.setText(text)
        # except:
        #     traceback.print_exc()
    '''Set drawing info when sketch change'''
    def drawing_text_change_5(self):
        try:
            page_num = PDFTools_v2.page_count(self.ui.text_drawing_sktechdir.text())
            self.ui.text_drawing_2.setText(str(page_num))
            drawing_name=os.path.join(Path(self.ui.text_drawing_sktechdir.text()).parent.absolute(),self.ui.text_all_3_projectname.text()+'-Mechanical Drawings.pdf')
            self.ui.text_drawing_output.setText(drawing_name)
            self.buttongroup_tag_click2()
        except:
            pass
    '''File change in Drawing Plot file from button'''
    def f_button_drawing_plot_filefrom(self):
        try:
            fileName = self.f_get_filename()
            if fileName!='':
                self.ui.textedit_drawing_plot_filefrom.setText(fileName)
        except:
            traceback.print_exc()
    '''File change in Drawing Plot file from text'''
    def f_text_drawing_plot_filefrom(self):
        try:
            self.ui.textedit_drawing_plot_filefrom.textChanged.disconnect(self.f_text_drawing_plot_filefrom)
        except:
            pass
        try:
            drawing_file_read=self.ui.textedit_drawing_plot_filefrom.toPlainText()
            if drawing_file_read.find("file:///")!=-1:
                drawing_file = drawing_file_read.split("file:///")[1]
                self.ui.textedit_drawing_plot_filefrom.setText(drawing_file)
            else:
                drawing_file=drawing_file_read
            if drawing_file=='':
                return
            try:
                file_path = Path(drawing_file)
                directory = file_path.parent
                pro_name = list(directory.parts)[1]
                loc_project_divide=pro_name.find('-')
                if loc_project_divide != -1:
                    proj_num = pro_name[:loc_project_divide]
                self.f_search_fileadded(proj_num)
                try:
                    self.ui.textedit_drawing_plot_filefrom.textChanged.disconnect(self.f_text_drawing_plot_filefrom)
                except:
                    pass
                self.ui.textedit_drawing_plot_filefrom.setText(drawing_file)
            except:
                traceback.print_exc()
            if self.ui.checkbox_plot_folder.isChecked():
                plot_folder = os.path.join(self.current_folder_address, 'Plot',str(date.today().strftime("%Y%m%d")) + '-plot')
            else:
                plot_folder = os.path.join(self.current_folder_address, 'Plot')
            plot_name=Path(drawing_file).name
            self.ui.text_drawing_plot_folderto.setText(plot_folder)
            self.ui.text_drawing_plot_fileto.setText(plot_name)
            total_page = PDFTools_v2.page_count(drawing_file)
            self.ui.text_drawing_plot_pagerange.setText('1-'+str(total_page))
        except:
            traceback.print_exc()
        finally:
            self.ui.textedit_drawing_plot_filefrom.textChanged.connect(self.f_text_drawing_plot_filefrom)
    '''Clear content in Drawing Plot'''
    def f_button_drawing_plot_fileclear(self):
        self.ui.textedit_drawing_plot_filefrom.setText('')
        self.ui.text_drawing_plot_folderto.setText('')
        self.ui.text_drawing_plot_fileto.setText('')
        self.ui.text_drawing_plot_pagerange.setText('')
    '''Drawing Plot generate link'''
    def f_button_issue_link(self):
        try:
            quotation_number=self.ui.text_all_1_projectnumber.text()
            link_path_ori=link_path
            folder_for_zip=self.ui.text_issue_folder.text()
            zip_name=os.path.join(conf["c_temp_dir"],Path(folder_for_zip).name+'.zip')
            self.progress_dialog = ProgressDialog('Generating link.',self.ui)
            self.link = Link(self.progress_dialog, zip_name,folder_for_zip)
            self.link.start()
            self.progress_dialog.exec_()
            if link_path!=link_path_ori:
                self.ui.text_issue_link.setText(link_path)
                files=list_all_files(folder_for_zip)
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("INSERT INTO bridge.link VALUES (%s, %s, %s, %s, %s)",(link_path, quotation_number, str(datetime.now().strftime("%Y%m%d")), 0, '&&&'.join(files),))
                except:
                    traceback.print_exc()
                finally:
                    mydb.commit()
                    mycursor.close()
                    mydb.close()
                self.f_get_links_from_database()



            else:
                message('Generate link error',self.ui)
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Get link table from database'''
    def f_get_links_from_database(self):
        try:
            table=self.ui.table_plot_link_and_date
            quotation_number=self.ui.text_all_1_projectnumber.text()
            if quotation_number=='':
                return
            link_database=[]
            table.setRowCount(0)
        except:
            traceback.print_exc()
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM bridge.link WHERE quotation_number = %s",(quotation_number,))
            link_database = mycursor.fetchall()
        except:
            traceback.print_exc()
        finally:
            mycursor.close()
            mydb.close()
        try:
            if len(link_database)>0:
                for i in range(len(link_database)):
                    link_info_i=link_database[i]
                    google_link=link_info_i[0]
                    gen_date=link_info_i[2]
                    file_names=link_info_i[4]
                    table.setRowCount(i + 1)
                    table.setItem(i, 0, QTableWidgetItem(str(google_link)))
                    table.setItem(i, 1, QTableWidgetItem(str(gen_date)))
                    table.setItem(i, 2, QTableWidgetItem(str(file_names)))
        except:
            traceback.print_exc()


    '''Drawing Plot analyze link'''
    def f_table_plot_link_and_date(self):
        try:
            table_link=self.ui.table_plot_link_and_date
            table_files=self.ui.table_drawing_linkcontent
            table_files.setRowCount(0)
            current_row = table_link.currentRow()
            if current_row >= 0:
                files_text=table_link.item(current_row, 2).text()
                file_absolute_list = files_text.split("&&&")
                for i in range(len(file_absolute_list)):
                    if i==0:
                        folder_name=Path(file_absolute_list[i]).parent.absolute()
                        self.ui.text_plot_link_folder.setText(str(folder_name))
                    file_name=Path(file_absolute_list[i]).name
                    table_files.setRowCount(i + 1)
                    table_files.setItem(i, 0, QTableWidgetItem(str(file_name)))
                self.ui.frame_link_files.show()
        except:
            traceback.print_exc()
    '''Drawing Plot copy link'''
    def f_button_issue_copylink(self):
        try:
            link=self.ui.text_issue_link.text()
            pyperclip.copy(link)
        except:
            traceback.print_exc()
    '''Drawing link folder set from plot folder'''
    def f_text_drawing_plot_folderto(self):
        folder_name=self.ui.text_drawing_plot_folderto.text()
        self.ui.text_issue_folder.setText(folder_name)
    '''Drawing get logo'''
    def f_button_drawing_chooselogo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        try:
            fileName, _ = QFileDialog.getOpenFileName(self.ui, "QFileDialog.getOpenFileName()",
                                                      self.current_folder_address,
                                                      "PNG Files (*.png *.PNG)", options=options)
            self.ui.text_logo_dir.setText(fileName)
        except:
            traceback.print_exc()
    '''Drawing - Open sketch'''
    def f_button_open_sketch(self):
        try:
            file_dir=self.ui.text_drawing_sktechdir.text()
            open_in_bluebeam(file_dir)
        except:
            traceback.print_exc()
    '''Drawing - Change sketch dir'''
    def f_button_changesketchdir(self):
        try:
            fileName = self.f_get_filename()
            self.ui.text_drawing_sktechdir.setText(fileName)
            if fileName != '':
                directory = Path(fileName).parent
                pro_name = list(directory.parts)[1]
                loc_project_divide = pro_name.find('-')
                if loc_project_divide != -1:
                    proj_num = pro_name[:loc_project_divide]
                self.f_search_fileadded(proj_num)
                self.ui.text_drawing_sktechdir.setText(fileName)
        except:
            traceback.print_exc()

    '''Set drawing sketch dir from Sketch align'''
    def f_text_align_outfile(self):
        try:
            file=self.ui.text_align_outfile.text()
            self.ui.text_drawing_sktechdir.setText(file)
        except:
            traceback.print_exc()
    '''Drawing - Change output file name'''
    def f_button_drawing_changeout(self):
        try:
            file_direc_read = QFileDialog.getExistingDirectory(None, "Choose file", self.current_folder_address)
            new_dir=os.path.join(file_direc_read,self.ui.text_all_3_projectname.text()+'-Mechanical Drawings.pdf')
            self.ui.text_drawing_output.setText(new_dir)
        except:
            traceback.print_exc()
    '''Drawing - Get logo from snipping'''
    def f_getlogo(self):
        try:
            file_dir=os.path.join(self.current_folder_address,'ss',time.strftime("%Y%m%d") + '-logo.png')
            self.ui.text_logo_dir.setText(file_dir)
            self.snipping_thread = SnippingThread()
            self.snipping_thread.start()
        except:
            traceback.print_exc()
    '''Drawing open logo folder'''
    def f_open_logo_folder(self):
        try:
            file_dir=self.ui.text_logo_dir.text()
            folder_dir=Path(file_dir).parent.absolute()
            open_folder(folder_dir)
        except:
            traceback.print_exc()
    '''Drawing table set when description change'''
    def f_combox_drawing_update_des_all(self,rev_num):
        try:
            ver_list=['A','B','C','D','E','F','G']
            if self.description_a_list[rev_num].currentText()=='':
                self.rev_list[rev_num].setText('')
                self.description_b_list[rev_num].setText('')
                self.drw_list[rev_num].setText('')
                self.chk_list[rev_num].setText('')
                self.date_list[rev_num].setText('')
            else:
                if self.rev_list[rev_num].text() == '':
                    self.rev_list[rev_num].setText(ver_list[rev_num])
                    self.description_b_list[rev_num].setText('')
                    self.drw_list[rev_num].setText('DW')
                    self.chk_list[rev_num].setText('FY')
                    self.date_list[rev_num].setText(str(date.today().strftime("%Y%m%d")))
        except:
            traceback.print_exc()
    '''Drawing - Plot'''
    def f_button_drawing_plot(self):
        try:
            file_from=self.ui.textedit_drawing_plot_filefrom.toPlainText()
            folder_to=self.ui.text_drawing_plot_folderto.text()
            file_to = os.path.join(folder_to,self.ui.text_drawing_plot_fileto.text())
            if is_pdf_open(file_from):
                message(f'Please close the pdf at {file_from}',self.ui)
                return
            if is_pdf_open(file_to):
                message(f'Please close the pdf at {file_to}',self.ui)
                return
            if file_exists(file_to):
                overwrite=messagebox('Overwrite',f'{file_to} exists, do you want to overwrite?',self.ui)
                if not overwrite:
                    return
            pages_cus=self.ui.text_drawing_plot_pagerange.text()
            page_list = parse_page_ranges(pages_cus)
            create_directory(folder_to)
            wait_prcocess=f_waitinline('0000')
            if wait_prcocess!='ok now':
                set =messagebox('Waiting confirm', 'Someone is '+wait_prcocess+', do you want to wait?',self.ui)
                if set:
                    shutil.copy(file_from, file_to)
                    remove_pdf_page(file_to, file_to, page_list)
                    file_bytes_ori = get_file_bytes(file_to)
                    file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                    txt_name = os.path.join(file_dir_trans, 'file_names_plot.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    papersize = get_paper_size(file_to, 1)
                    with open(txt_name, 'w') as file:
                        file_names = {'File1': file_to, 'Papersize': papersize}
                        file.write(str(file_names))
                    self.progress_dialog = ProgressDialog('Someone is '+wait_prcocess+', your task is waiting in line.',self.ui)
                    self.waitinline = Waitinline(self.progress_dialog,file_time)
                    self.waitinline.start()
                    self.progress_dialog.exec_()
                else:
                    return
            else:
                shutil.copy(file_from, file_to)
                remove_pdf_page(file_to, file_to, page_list)
                file_bytes_ori = get_file_bytes(file_to)
                file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                txt_name = os.path.join(file_dir_trans, 'file_names_plot.txt')
                os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                papersize = get_paper_size(file_to, 1)
                with open(txt_name, 'w') as file:
                    file_names = {'File1': file_to, 'Papersize': papersize}
                    file.write(str(file_names))
            self.progress_dialog = ProgressDialog('Plotting, open in Bluebeam when done.',self.ui)
            self.plot = Process2PC(self.progress_dialog, ["Plot",file_dir_trans])
            self.plot.start()
            self.progress_dialog.exec_()
            try:
                time.sleep(2)
                file_bytes_new = get_file_bytes(file_to)
                quo_num = self.ui.text_all_1_projectnumber.text()
                pro_num = self.ui.text_all_2_quotationnumber.text()
                pro_name = self.ui.text_all_3_projectname.text()
                pro_info_str = quo_num + '-' + pro_num + '-' + pro_name
                if abs(int(file_bytes_new)-int(file_bytes_ori))>10:
                    flatten_pdf(file_to, file_to)
                    open_in_bluebeam(file_to)
                    message(str(file_to) + '\nPlot successfully', self.ui)
                    # message(pro_info_str+'\nPlot successfully',self.ui)
                else:
                    message(pro_info_str + '\nPlot failed',self.ui)
                    try:
                        cancel_task_txt=conf['cancel_task_txt']
                        with open(cancel_task_txt, 'r',encoding='utf-8') as file:
                            content = file.read()
                        new_content = content + ","+file_time
                        with open(cancel_task_txt, 'w',encoding='utf-8') as file:
                            file.write(new_content)
                    except:
                        traceback.print_exc()
                    try:
                        file_move_name=str(Path(file_to).name)
                        recycle_bin_dir = os.path.join(conf["recycle_bin_dir"], str(datetime.now().strftime("%Y%m%d%H%M%S")),file_move_name)
                        os.makedirs(os.path.dirname(recycle_bin_dir), exist_ok=True)
                        shutil.move(file_to, recycle_bin_dir)
                    except:
                        traceback.print_exc()
            except:
                traceback.print_exc()
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Drawing issue type customize'''
    def f_text_issue_type(self):
        try:
            if str(self.ui.text_issue_type.text()) != '':
                self.ui.checkbox_issue_type.setChecked(True)
            else:
                self.ui.checkbox_issue_type.setChecked(False)
        except:
            traceback.print_exc()
    '''Drawing - Drawing table change when tag change'''
    def buttongroup_tag_click2(self):
        try:
            tags = []
            for i in range(len(self.button_group_tag2_boxlist)):
                if self.button_group_tag2_boxlist[i].isChecked():
                    tag = self.button_group_tag2_textlist[i].text()
                    tags.append(tag)
            self.set_draw_table(tags)
        except:
            traceback.print_exc()
    '''Drawing - Set table'''
    def set_draw_table(self,tags):
        try:
            drawing_num = ['M-000']
            drawing_name = ['COVERSHEET, GENERAL NOTES, LEGEND AND DETAILS']
            for i in range(len(tags)):
                num=100+i
                drawing_num.append('M-' + str(num))
                drawing_name.append(tags[i].upper()+' LAYOUT')
            revision_list=[]
            for i in range(len(drawing_num)):
                revision_list.append('A')
            self.ui.table_drawing.setRowCount(0)
            checkboxes = []
            self.drawing_checkbox_list=[]
            for i in range(len(drawing_num)):
                self.ui.table_drawing.insertRow(i)
                self.ui.table_drawing.setItem(i, 0, QTableWidgetItem(str(drawing_num[i])))
                self.ui.table_drawing.setItem(i, 1, QTableWidgetItem(str(drawing_name[i])))
                self.ui.table_drawing.setItem(i, 2, QTableWidgetItem(str(revision_list[i])))
                checkbox = QCheckBox()
                checkboxes.append(checkbox)
                self.drawing_checkbox_list.append(checkbox)
        except:
            traceback.print_exc()
    '''Drawing - Open plot folder'''
    def f_button_issue_openfolder(self):
        try:
            folder_name=self.ui.text_issue_folder.text()
            open_folder(folder_name)
        except:
            traceback.print_exc()
    '''Drawing - Setup drawing'''
    def f_button_drawing_gettem(self):
        try:
            sketch=self.ui.text_drawing_sktechdir.text()
            pages=get_number_of_page(sketch)
            scale=self.ui.text_drawing_3.text()
            if self.ui.radiobutton_drawing_rest.isChecked():
                drawing_type='restaurant'
            else:
                drawing_type = 'others'
            paper_size=self.ui.text_drawing_4.text()
            if paper_size not in ['A3','A1','A0']:
                message('Undefined paper size',self.ui)
                return
            temp1,temp2,temp3=get_temp_name(scale,drawing_type,paper_size)
            output_file_text=self.ui.text_drawing_output.text()
            output_file=os.path.join(Path(output_file_text).parent.absolute(),Path(output_file_text).name)
            self.copy2ss(output_file)
            shutil.copy(temp2,output_file)
            page_delete(output_file,output_file,pages)
            move_sketch_onoff=self.ui.checkbox_drawing_movesketch.isChecked()
            if move_sketch_onoff:
                self.progress_dialog = ProgressDialog("Moving sketch to paper center.",self.ui)
                self.move_sketch = Process2PC(self.progress_dialog,["Movesketch",sketch,[paper_size]])
                self.move_sketch.start()
                self.progress_dialog.exec_()
            wait_prcocess=f_waitinline('0000')
            if wait_prcocess != 'ok now':
                set = messagebox('Waiting confirm','Someone is ' + wait_prcocess + ', do you want to wait?',self.ui)
                if set:
                    file_time=str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans=os.path.join(conf["trans_dir"],file_time)
                    txt_name=os.path.join(file_dir_trans,'file_names_setupdrawing.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    with open(txt_name, 'w') as file:
                        file_names=sketch+';'+output_file+';'
                        file.write(file_names)
                    self.progress_dialog = ProgressDialog('Someone is '+wait_prcocess+', your task is waiting in line.',self.ui)
                    self.waitinline = Waitinline(self.progress_dialog,file_time)
                    self.waitinline.start()
                    self.progress_dialog.exec_()
                else:
                    return
            else:
                file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                txt_name = os.path.join(file_dir_trans, 'file_names_setupdrawing.txt')
                os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                with open(txt_name, 'w') as file:
                    file_names = sketch + ';' + output_file + ';'
                    file.write(file_names)
            self.progress_dialog = ProgressDialog('Changing sketch to drawing.',self.ui)
            self.snapsketch = Process2PC(self.progress_dialog, ["Snapsketch",file_dir_trans])
            self.snapsketch.start()
            self.progress_dialog.exec_()
            if self.ui.radiobutton_drawing_construction.isChecked():
                issure_des='FOR CONSTRUCTION'
            elif self.ui.checkbox_issue_type.isChecked():
                issure_des = self.ui.text_issue_type.text().upper()
            else:
                issure_des ='FOR APPROVAL'
            content_replace_dict_list = []
            project_number=self.ui.text_all_2_quotationnumber.text().upper()
            if project_number=='':
                project_number=self.ui.text_all_1_projectnumber.text().upper()
            for i in range(pages):
                content_replace_dict = {}
                content_replace_dict['Item_1_rev'] = str(self.ui.table_drawing.item(i + 1, 2).text())
                content_replace_dict['Item_2_rev_des'] = 'ISSUED ' + issure_des
                content_replace_dict['Item_3_drw'] = self.ui.text_drawing_dw.text().upper()
                content_replace_dict['Item_4_chk'] = self.ui.text_drawing_ck.text().upper()
                content_replace_dict['Item_5_date'] = str(date.today().strftime("%Y%m%d"))
                content_replace_dict['Item_6_address'] = self.ui.text_all_3_projectname.text().upper()
                content_replace_dict['Item_7_title'] = str(self.ui.table_drawing.item(i + 1, 1).text()).upper()
                content_replace_dict[
                    'Item_8_scale'] = self.ui.text_drawing_3.text() + '@' + self.ui.text_drawing_4.text().upper()
                content_replace_dict['Item_9_prono'] = project_number
                content_replace_dict['Item_10_drwno'] = str(self.ui.table_drawing.item(i + 1, 0).text()).upper()
                content_replace_dict['Item_11_issue'] = issure_des
                content_replace_dict_list.append(content_replace_dict)
            content_replace_dict_firstpage = {}
            content_replace_dict_firstpage['Item_1_rev'] = str(self.ui.table_drawing.item(0, 2).text())
            content_replace_dict_firstpage['Item_2_rev_des'] = 'ISSUED ' + issure_des
            content_replace_dict_firstpage['Item_3_drw'] = self.ui.text_drawing_dw.text().upper()
            content_replace_dict_firstpage['Item_4_chk'] = self.ui.text_drawing_ck.text().upper()
            content_replace_dict_firstpage['Item_5_date'] = str(date.today().strftime("%Y%m%d"))
            content_replace_dict_firstpage['Item_6_address'] = self.ui.text_all_3_projectname.text().upper()
            content_replace_dict_firstpage['Item_7_title'] = str(self.ui.table_drawing.item(0, 1).text()).upper()
            content_replace_dict_firstpage['Item_8_scale'] = 'N.T.S@' + self.ui.text_drawing_4.text().upper()
            content_replace_dict_firstpage['Item_9_prono'] = project_number
            content_replace_dict_firstpage['Item_10_drwno'] = str(self.ui.table_drawing.item(0, 0).text()).upper()
            content_replace_dict_firstpage['Item_11_issue'] = issure_des
            self.progress_dialog = ProgressDialog('Setting frame information.',self.ui)
            self.changeframe = Changeframe(self.progress_dialog, temp1,output_file,temp3,content_replace_dict_list,content_replace_dict_firstpage)
            self.changeframe.start()
            self.progress_dialog.exec_()
            if self.ui.text_logo_dir.text()!='':
                self.progress_dialog = ProgressDialog('Setting client logo, open in Bluebeam when done.',self.ui)
                self.pastelogo = Pastelogo(self.progress_dialog, self.ui.text_logo_dir.text(),output_file,paper_size)
                self.pastelogo.start()
                self.progress_dialog.exec_()
            open_in_bluebeam(output_file)
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Drawing - Update drawing frame'''
    def f_button_drawing_update(self):
        try:
            pdf_file = self.ui.text_update_filepath.text()
            for i in range(len(self.rev_list)):
                if self.rev_list[len(self.rev_list)-i-1].text()!='':
                    latest_version_num=len(self.rev_list)-i-1
                    break
            drw_by=self.drw_list[latest_version_num].text()
            chk_by=self.chk_list[latest_version_num].text()
            issue_type=self.description_a_list[latest_version_num].currentText()[6:]+self.description_b_list[latest_version_num].text()
            pro_no=self.ui.text_drawing_update_prono.text()
            pro_name=self.ui.text_drawing_update_project.toPlainText()
            scale_a1='N.T.S@'
            scale_a2 =self.ui.text_drawing_update_scale.text().split('@')[0]+'@'
            scale_b=self.ui.text_drawing_update_scale.text().split('@')[1]
            new_rev_list_all=[]
            for i in range(latest_version_num+1):
                rev=self.rev_list[i].text()
                des=self.description_a_list[i].currentText()+self.description_b_list[i].text()
                drw=self.drw_list[i].text()
                chk=self.chk_list[i].text()
                dat=self.date_list[i].text()
                new_rev_list_all.append([rev,des,drw,chk,dat])
            if is_float(self.ui.text_drawing_update_coverpage.text()):
                last_cover_page=int(self.ui.text_drawing_update_coverpage.text())
            else:
                last_cover_page =int(self.ui.text_drawing_update_coverpage.text().split('-')[1])
            if is_float(self.ui.text_drawing_update_drawingpage.text()):
                last_drawing_page=int(self.ui.text_drawing_update_drawingpage.text())
            else:
                last_drawing_page =int(self.ui.text_drawing_update_drawingpage.text().split('-')[1])
            self.progress_dialog = ProgressDialog('Updating frame information, open in Bluebeam when done.',self.ui)
            self.updateframe = UpdateFrame(self.progress_dialog, pdf_file,last_drawing_page,self.drawingframe_context_list,chk_by,drw_by,issue_type,pro_no,pro_name,new_rev_list_all,scale_a1,scale_b,scale_a2,last_cover_page)
            self.updateframe.start()
            self.progress_dialog.exec_()
            open_in_bluebeam(pdf_file)
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Drawing - Get drawing frame info'''
    def f_button_changeupdatefile(self):
        try:
            fileName = self.f_get_filename()
            self.ui.text_update_filepath.setText(fileName)
            for content in self.rev_list+self.description_b_list+self.drw_list+self.chk_list+self.date_list:
                content.setText('')
            for content in self.description_a_list:
                content.setCurrentIndex(0)
            total_page=PDFTools_v2.page_count(fileName)
            self.ui.text_drawing_update_totalpage.setText(str(total_page))
            paper_size = get_paper_size(fileName, 0)
            self.ui.text_drawing_update_papersize.setText(str(paper_size))
            drawingframe_context_list=[]
            cover_page_list=[]
            drawing_page_list=[]
            other_page_list=[]
            first_drawing_page=[]
            for i in range(int(total_page)):
                structured_context=version_setter(fileName, i+1, paper_size)
                drawingframe_context_list.append(structured_context)
                scale = get_frame_content(structured_context, '',"SCALE",0)
                if scale[:5]=='N.T.S':
                    cover_page_list.append(i+1)
                elif scale.find('@')!=-1:
                    drawing_page_list.append(i+1)
                    first_drawing_page.append(i+1)
                else:
                    other_page_list.append(i+1)
            try:
                if cover_page_list[0]==cover_page_list[-1]:
                    cover_page_range=str(cover_page_list[0])
                else:
                    cover_page_range =str(cover_page_list[0])+'-'+str(cover_page_list[-1])
            except:
                cover_page_range=''
            self.ui.text_drawing_update_coverpage.setText(cover_page_range)
            try:
                if drawing_page_list[0]==drawing_page_list[-1]:
                    drawing_page_range=str(drawing_page_list[0])
                else:
                    drawing_page_range =str(drawing_page_list[0])+'-'+str(drawing_page_list[-1])
            except:
                drawing_page_range=''
            self.ui.text_drawing_update_drawingpage.setText(drawing_page_range)
            try:
                if other_page_list[0]==other_page_list[-1]:
                    other_page_range=str(other_page_list[0])
                else:
                    other_page_range =str(other_page_list[0])+'-'+str(other_page_list[-1])
            except:
                other_page_range=''
            self.ui.text_other_update_drawingpage.setText(other_page_range)
            if first_drawing_page!=[]:
                if first_drawing_page[0]<int(self.ui.text_drawing_update_totalpage.text()):
                    drawing_context=drawingframe_context_list[first_drawing_page[0]-1]
                    self.drawingframe_context_list=drawingframe_context_list
                    scale = get_frame_content(drawing_context,'', "SCALE",0)
                    project = get_frame_content(drawing_context,'', "PROJECT",0)
                    pro_no = get_frame_content(drawing_context,'', "PROJECT No.",0)
                    issue_type = get_frame_content(drawing_context,'', "TYPE OF ISSUE",0)
                    if project!=self.ui.text_all_3_projectname.text().upper():
                        self.ui.text_drawing_update_project.setText(self.ui.text_all_3_projectname.text().upper())
                        self.ui.text_drawing_update_project.setStyleSheet("""
                            QTextEdit {background-color: #ffffff;color: #ff0000;font: bold 18px "Calibri";}""")
                    else:
                        self.ui.text_drawing_update_project.setText(self.ui.text_all_3_projectname.text().upper())
                    if pro_no!=self.ui.text_all_2_quotationnumber.text():
                        if pro_no!=self.ui.text_all_1_projectnumber.text():
                            if self.ui.text_all_2_quotationnumber.text()!='':
                                self.ui.text_drawing_update_prono.setText(self.ui.text_all_2_quotationnumber.text())
                            else:
                                self.ui.text_drawing_update_prono.setText(self.ui.text_all_1_projectnumber.text())
                            self.ui.text_drawing_update_prono.setStyleSheet("""QLineEdit {background-color: #ffffff;color: #ff0000;font: bold 18px "Calibri";}""")
                        else:
                            self.ui.text_drawing_update_prono.setText(self.ui.text_all_1_projectnumber.text())
                    else:
                        self.ui.text_drawing_update_prono.setText(self.ui.text_all_2_quotationnumber.text())
                    self.ui.text_drawing_update_scale.setText(scale)
                    self.ui.text_drawing_update_issuetype.setText(issue_type)

                    version = drawing_context['VERSION']
                    for i in range(len(version)):
                        REV = get_frame_content(drawing_context, 'VERSION', "REV", i)
                        DESCRIPTION = get_frame_content(drawing_context, 'VERSION', "REVISION DESCRIPTION", i)
                        DRW = get_frame_content(drawing_context, 'VERSION', "DRW", i)
                        CHK = get_frame_content(drawing_context, 'VERSION', "CHK", i)
                        DATE = get_frame_content(drawing_context, 'VERSION', "DATE Y.M.D", i)
                        self.rev_list[i].setText(REV)
                        self.drw_list[i].setText(DRW)
                        self.chk_list[i].setText(CHK)
                        self.date_list[i].setText(DATE)
                        if DESCRIPTION=='ISSUED FOR CONSTRUCTION':
                            self.description_a_list[i].setCurrentIndex(1)
                            self.description_b_list[i].setText('')
                        elif DESCRIPTION=='ISSUED FOR APPROVAL':
                            self.description_a_list[i].setCurrentIndex(2)
                            self.description_b_list[i].setText('')
                        elif DESCRIPTION=='ISSUED FOR TENDER':
                            self.description_a_list[i].setCurrentIndex(3)
                            self.description_b_list[i].setText('')
                        else:
                            self.description_a_list[i].setCurrentIndex(4)
                            try:
                                self.description_b_list[i].setText(DESCRIPTION.split(' ')[2])
                            except:
                                self.description_b_list[i].setText('/')
                    pic_name=os.path.join(conf['c_temp_dir'],'client_logo.png')
                    getlogo(fileName,pic_name,paper_size)
                    pixmap = QPixmap(pic_name)
                    self.ui.label_drawing_update_logo.setPixmap(pixmap)
                    self.ui.label_drawing_update_logo.setScaledContents(True)
            else:
                message('Cannot find drawing page.',self.ui)
                return
        except:
            traceback.print_exc()
    '''======================================Plot tab=============================================='''
    def f_text_issue_folder(self):
        try:
            table=self.ui.table_drawing_foldercontent
            folder_dir=self.ui.text_issue_folder.text()
            file_list = os.listdir(folder_dir)
            table.setRowCount(0)
            for file in file_list:
                row_position = table.rowCount()
                table.insertRow(row_position)
                table.setItem(row_position, 0, QTableWidgetItem(file))
        except:
            traceback.print_exc()

    '''======================================Overlay tab==========================================='''
    '''Service type change in Overlay'''
    def f_text_overlay_servicetype(self):
        try:
            if self.ui.text_overlay_servicetype.text()!='':
                self.button_overlay_group.button(8).setChecked(True)
                service_name=self.ui.text_overlay_servicetype.text()
                self.service_name = service_name
                self.ui.text_overlay_outfile.setText(os.path.join(self.ui.text_overlay_folder1.text(),self.ui.text_overlay_folder2.text(),str(date.today().strftime("%Y%m%d"))+'-Combined '+service_name+'.pdf'))
            else:
                self.button_overlay_group.button(8).setChecked(False)
        except:
            traceback.print_exc()
    '''File change in Overaly markup from'''
    def f_button_overlay_markupfrom_change(self):
        try:
            fileName = self.f_get_filename()
            if fileName!='':
                self.ui.text_overlay_markupfrom_file.setText(fileName)
            page_num=get_number_of_page(fileName)
            self.ui.text_pagenum_markupfrom.setText(str(page_num))
            paper_size=get_paper_size(fileName, 0)
            self.ui.text_pagesize_markupfrom.setText(paper_size)
            try:
                self.preview_pdf()
            except:
                pass
        except:
            traceback.print_exc()
    '''File change in Overaly markup to'''
    def f_button_overlay_markupto_change(self):
        try:
            fileName = self.f_get_filename()
            if fileName!='':
                self.ui.text_overlay_markupto_file.setText(fileName)
            page_num=get_number_of_page(fileName)
            self.ui.text_pagenum_markupto.setText(str(page_num))
            paper_size=get_paper_size(fileName, 0)
            self.ui.text_pagesize_markupto.setText(paper_size)
            try:
                self.preview_pdf()
            except:
                pass
        except:
            traceback.print_exc()
    '''Overlay copy markup frame set'''
    def preview_pdf(self):
        try:
            file1=self.ui.text_overlay_markupfrom_file.text()
            file2=self.ui.text_overlay_markupto_file.text()
            for child in self.frame_copymk_from.children():
                if isinstance(child, QWidget):
                    self.frame_copymk_from.layout().removeWidget(child)
                    child.deleteLater()
            for child in self.frame_copymk_to.children():
                if isinstance(child, QWidget):
                    self.frame_copymk_to.layout().removeWidget(child)
                    child.deleteLater()
            pages1=PDFTools_v2.page_count(file1)
            pages2 = PDFTools_v2.page_count(file2)
            checkbox_style="""QCheckBox {background-color: rgb(165, 42, 42); color: rgb(255, 255, 255); border-style: outset; border-width: 2px; border-radius: 10px; border-color: rgb(80, 80, 80); font: bold 15px; padding: 6px; width: 40px;height: 18px;}"""
            self.copymkfrom_preview_pagecheck=[]
            for i in range(pages1):
                pic_name= os.path.join(conf['c_temp_dir'], 'mrk_from'+str(i)+'.png')
                getpage(file1, pic_name, i)
                pixmap = QPixmap(pic_name)
                checkbox = QCheckBox('Page '+str(i+1))
                checkbox.setStyleSheet(checkbox_style)
                self.copymkfrom_preview_pagecheck.append(checkbox)
                label1=QLabel('label')
                label1.setFixedSize(300, 250)
                row_layout = QHBoxLayout()
                row_layout.addWidget(checkbox)
                row_layout.addWidget(label1)
                self.layout_copymk_from.addLayout(row_layout, i, 0)
                checkbox.setChecked(True)
                label1.setPixmap(pixmap)
                label1.setScaledContents(True)

            self.copymkto_preview_pagecheck = []
            for i in range(pages2):
                pic_name = os.path.join(conf['c_temp_dir'], 'mrk_to' + str(i) + '.png')
                getpage(file2, pic_name, i)
                pixmap = QPixmap(pic_name)
                checkbox = QCheckBox('Page ' + str(i + 1))
                checkbox.setStyleSheet(checkbox_style)
                self.copymkto_preview_pagecheck.append(checkbox)
                label2 = QLabel('label')
                label2.setFixedSize(300, 250)
                row_layout2 = QHBoxLayout()
                row_layout2.addWidget(checkbox)
                row_layout2.addWidget(label2)
                self.layout_copymk_to.addLayout(row_layout2, i, 0)
                checkbox.setChecked(True)
                label2.setPixmap(pixmap)
                label2.setScaledContents(True)
        except:
            traceback.print_exc()
    '''Grayscale checkbox show or not when overlay is selected'''
    def f_checkbox_overlay_update(self):
        try:
            pass
        except:
            traceback.print_exc()
    '''Overlay - Grayscale'''
    def f_overlay_gray(self,gray_file):
        try:
            self.progress_dialog = ProgressDialog("Grayscaling, open in Bluebeam when done.",self.ui)
            self.gray = Gray(self.progress_dialog, gray_file)
            self.gray.start()
            self.progress_dialog.exec_()
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Overlay - Output file change'''
    def f_text_overlay_outfile_change(self):
        try:
            try:
                service_name=self.service_name
                overlay_file_name=service_name+ ' on Mech overlay.pdf'
            except:
                service_name=''
                overlay_file_name ='Overlay.pdf'
            self.ui.text_overlay_outfile_2.setText(os.path.join(Path(self.ui.text_overlay_outfile.text()).parent.absolute(),overlay_file_name))
            self.ui.text_overlay_outfile_3.setText(os.path.join(Path(self.ui.text_overlay_outfile.text()).parent.absolute(), 'Mech with updated background.pdf'))

            layer_name=self.ui.text_overlay_layer.text()
            if layer_name.find('-')!=-1:
                if service_name!='':
                    new_layer_name=service_name+'-'+layer_name.split('-')[1]
                else:
                    new_layer_name=layer_name.split('-')[1]
            else:
                if service_name!='':
                    new_layer_name=service_name+'-'+layer_name
                else:
                    new_layer_name=layer_name
            self.ui.text_overlay_layer.setText(new_layer_name)

        except:
            traceback.print_exc()
    '''Overlay service type clicked'''
    def f_buttongroup_overlay_clicked(self,button):
        try:
            button_id = self.button_overlay_group.id(button)
            service_name_list=['Architecture','Structure','Stormwater','Hydraulic','Electrical','Fire','Background','?']
            if button_id<8:
                service_name=service_name_list[button_id-1]
                if button_id<7:
                    self.ui.checkbox_overlay_overlay.setChecked(True)
                    self.ui.checkbox_overlay_update.setChecked(False)
                else:
                    self.ui.checkbox_overlay_overlay.setChecked(True)
                    self.ui.checkbox_overlay_update.setChecked(True)
            else:
                service_name=self.ui.text_overlay_servicetype.text()
            self.service_name=service_name
            self.ui.text_overlay_outfile.setText(os.path.join(self.ui.text_overlay_folder1.text(),self.ui.text_overlay_folder2.text(),str(date.today().strftime("%Y%m%d"))+'-Combined '+service_name+'.pdf'))
        except:
            traceback.print_exc()
    '''Overlay  - Copymarkup'''
    def f_button_overlay_copymarkup(self):
        try:
            file_markupfrom=self.ui.text_overlay_markupfrom_file.text()
            file_markupto = self.ui.text_overlay_markupto_file.text()
            if is_pdf_open(file_markupfrom) == True or is_pdf_open(file_markupto) == True:
                message('Please close the first pdf while combining',self.ui)
                return
            if self.ui.text_pagesize_markupfrom.text()!=self.ui.text_pagesize_markupto.text():
                message('Paper sizes are no the same',self.ui)
                return
            copyornot_from_list=[]
            from_true=0
            to_true=0
            for checkbox in self.copymkfrom_preview_pagecheck:
                if checkbox.isChecked():
                    copyornot_from_list.append(True)
                    from_true+=1
                else:
                    copyornot_from_list.append(False)
            copyornot_to_list=[]
            for checkbox in self.copymkto_preview_pagecheck:
                if checkbox.isChecked():
                    copyornot_to_list.append(True)
                    to_true += 1
                else:
                    copyornot_to_list.append(False)
            if from_true!=to_true:
                message('Pages are not the same',self.ui)
                return
            file_bytes_ori=get_file_bytes(file_markupto)
            wait_prcocess=f_waitinline('0000')
            if wait_prcocess!='ok now':
                set=messagebox('Waiting confirm','Someone is ' + wait_prcocess + ', do you want to wait?',self.ui)
                if set:
                    file_time=str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans=os.path.join(conf["trans_dir"],file_time)
                    txt_name=os.path.join(file_dir_trans,'file_names_copymarkup.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    with open(txt_name, 'w') as file:
                        file_names={'File1': file_markupfrom, 'File2': file_markupto, 'copyornot': [copyornot_from_list,copyornot_to_list]}
                        file.write(str(file_names))
                    self.progress_dialog = ProgressDialog('Someone is '+wait_prcocess+', your task is waiting in line.',self.ui)
                    self.waitinline = Waitinline(self.progress_dialog,file_time)
                    self.waitinline.start()
                    self.progress_dialog.exec_()
                else:
                    return
            else:
                file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                txt_name = os.path.join(file_dir_trans, 'file_names_copymarkup.txt')
                os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                with open(txt_name, 'w') as file:
                    file_names = {'File1': file_markupfrom, 'File2': file_markupto,
                                  'copyornot': [copyornot_from_list, copyornot_to_list]}
                    file.write(str(file_names))
            self.progress_dialog = ProgressDialog('Copying markup, open in Bluebeam when done.',self.ui)
            self.copymarkup = Process2PC(self.progress_dialog, ["Copymarkup",file_dir_trans])
            self.copymarkup.start()
            self.progress_dialog.exec_()
            try:
                time.sleep(2)
                file_bytes_new = get_file_bytes(file_markupto)
                if abs(int(file_bytes_new)-int(file_bytes_ori))>10:
                    open_in_bluebeam(file_markupto)
                else:
                    message('Copy markup failed',self.ui)
            except:
                message('Copy markup failed',self.ui)
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Overlay - Service - Overlay'''
    def f_button_overlay_service1_overlay(self):
        try:
            self.f_button_overlay_align()
            input_file1 = os.path.join(self.ui.text_overlay_mech_outfolder.text(),
                                       self.ui.text_overlay_mech_filenameadd.text())
            input_file2 = self.ui.text_overlay_outfile.text()
            layername=self.ui.text_overlay_layer.text()
            pages1 = PDFTools_v2.page_count(input_file1)
            pages2 = PDFTools_v2.page_count(input_file2)
            if int(pages1) - int(pages2) != 0:
                message('Page numbers are not the same',self.ui)
                return
            if layername=='':
                message('Please add layer name',self.ui)
                return
            color_list = {'Architecture': '#FF0000', 'Structure': '#00FFFF', 'Stormwater': '#008000',
                          'Hydraulic': '#008000', 'Electrical': '#FFAA00', 'Fire': '#AA00FF','Background':None}
            if self.service_name == '':
                message('Choose overlay type first',self.ui)
                return
            try:
                newcolor = color_list[self.service_name]
            except:
                newcolor = None
            if self.ui.radiobutton_servicetype8.isChecked():
                if self.ui.radiobutton_cus_1.isChecked():
                    newcolor = '#0000FF'
                elif self.ui.radiobutton_cus_2.isChecked():
                    newcolor = '#808000'
                elif self.ui.radiobutton_cus_3.isChecked():
                    newcolor = '#FF6600'

            if self.ui.checkbox_overlay_overlay.isChecked()==False and self.ui.checkbox_overlay_update.isChecked()==False:
                message('Choose at least one option',self.ui)
                return

            if self.ui.checkbox_overlay_overlay.isChecked():
                outputfile = self.ui.text_overlay_outfile_2.text()
                outputfile2 = self.ui.text_overlay_outfile_3.text()
                wait_prcocess=f_waitinline('0000')
                print(wait_prcocess)
                if wait_prcocess!='ok now':
                    set = messagebox('Waiting confirm','Someone is ' + wait_prcocess + ', do you want to wait?',self.ui)
                    if set:
                        file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                        file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                        txt_name = os.path.join(file_dir_trans, 'file_names_overlay.txt')
                        os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                        with open(txt_name, 'w') as file:
                            file_contetnt = {'File1': input_file1, 'File2': input_file2, 'File3': outputfile,
                                             'File4': '', 'Color': newcolor, 'Name': layername}
                            file.write(str(file_contetnt))
                        self.progress_dialog = ProgressDialog('Someone is '+wait_prcocess+', your task is waiting in line.',self.ui)
                        self.waitinline = Waitinline(self.progress_dialog,file_time)
                        self.waitinline.start()
                        self.progress_dialog.exec_()
                    else:
                        return
                else:
                    file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                    txt_name = os.path.join(file_dir_trans, 'file_names_overlay.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    with open(txt_name, 'w') as file:
                        file_contetnt = {'File1': input_file1, 'File2': input_file2, 'File3': outputfile,
                                         'File4': '', 'Color': newcolor, 'Name': layername}
                        file.write(str(file_contetnt))

                self.progress_dialog = ProgressDialog('Overlaying, open in Bluebeam when done.',self.ui)
                self.overlay_pc = Process2PC(self.progress_dialog, ["Overlay_PC",file_dir_trans])
                self.overlay_pc.start()
                self.progress_dialog.exec_()
                self.ui.text_overlay_mech_filenameadd.setText(self.ui.text_overlay_outfile_2.text())
            if self.ui.checkbox_overlay_update.isChecked():
                outputfile2 = self.ui.text_overlay_outfile_3.text()
                wait_prcocess = f_waitinline('0000')
                print(wait_prcocess)
                if wait_prcocess != 'ok now':
                    set=messagebox('Waiting confirm','Someone is ' + wait_prcocess + ', do you want to wait?',self.ui)
                    if set:
                        file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                        file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                        txt_name = os.path.join(file_dir_trans, 'file_names_overlay.txt')
                        os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                        with open(txt_name, 'w') as file:
                            file_contetnt = {'File1': input_file1, 'File2': input_file2, 'File3': outputfile2,
                                             'File4': '','Color': None, 'Name': layername}
                            file.write(str(file_contetnt))
                        self.progress_dialog = ProgressDialog(
                            'Someone is ' + wait_prcocess + ', your task is waiting in line.',self.ui)
                        self.waitinline = Waitinline(self.progress_dialog, file_time)
                        self.waitinline.start()
                        self.progress_dialog.exec_()
                    else:
                        return

                else:
                    file_time = str(datetime.now().strftime("%Y%m%d%H%M%S"))
                    file_dir_trans = os.path.join(conf["trans_dir"], file_time)
                    txt_name = os.path.join(file_dir_trans, 'file_names_overlay.txt')
                    os.makedirs(os.path.dirname(txt_name), exist_ok=True)
                    with open(txt_name, 'w') as file:
                        file_contetnt = {'File1': input_file1, 'File2': input_file2, 'File3': outputfile2,
                                         'File4': '','Color': None, 'Name': layername}
                        file.write(str(file_contetnt))
                self.progress_dialog = ProgressDialog('Updating background, open in Bluebeam when done.',self.ui)
                self.overlay_pc = Process2PC(self.progress_dialog, ["Overlay_PC",file_dir_trans])
                self.overlay_pc.start()
                self.progress_dialog.exec_()
            try:
                open_in_bluebeam(outputfile)
            except:
                pass
            try:
                open_in_bluebeam(outputfile2)
            except:
                pass
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Overlay - Set new folder'''
    def f_button_overlay_set(self):
        try:
            fn1=self.ui.text_overlay_folder1.text()
            fn2=self.ui.text_overlay_folder2.text()
            folder_name = os.path.join(fn1, fn2)
            if os.path.exists(folder_name):
                open_folder(folder_name)
            else:
                create_directory(folder_name)
                open_folder(folder_name)
        except:
            traceback.print_exc()
    '''Overlay - Align'''
    def f_button_overlay_align(self):
        try:
            align_name1=self.ui.text_overlay_outfile.text()
            mech_file = os.path.join(self.ui.text_overlay_mech_outfolder.text(),self.ui.text_overlay_mech_filenameadd.text())
            tmp_file1=r'C:\Copilot_template\mech_tiqu_1.pdf'
            if file_exists(align_name1) == False:
                message('Align file does not exist',self.ui)
                return
            combine_list = [align_name1]
            if is_pdf_open(align_name1) == True or is_pdf_open(mech_file) == True:
                message('Please close the pdf while aligning',self.ui)
                return
            extract_first_page(mech_file, tmp_file1)
            combine_pdf([tmp_file1,align_name1],align_name1)
            self.f_sketch_align_align_1(align_name1)
            outputname=align_name1
            combine_list2=[]
            for file in combine_list:
                combine_list2.append(os.path.join(r'C:\Copilot_template',Path(file).name))
            outputname2 = os.path.join(r'C:\Copilot_template', Path(align_name1).name)
            if is_pdf_open(outputname) == True or is_pdf_open(outputname2) == True:
                message('Please close the pdf while aligning',self.ui)
            shutil.copy(outputname2,outputname)
            remove_first_page(align_name1, align_name1)
        except Exception as e:
            traceback.print_exc()
            raise e
    '''Overlay - Combine'''
    def f_overlay_mech_combine(self):
        try:
            combinefile_list =self.get_table_files("Overlay mech")
            current_pdf_parent_directory = str(self.ui.text_overlay_mech_foldername.text())
            output_file_dir = str(self.ui.text_overlay_mech_outfolder.text())
            name_extension=str(self.ui.text_overlay_mech_filenameadd.text())
            combine_pdf_dir=os.path.join(output_file_dir, name_extension)
            combinefile_list_ab=[]
            for i in range(len(combinefile_list)):
                combinefile_list_ab.append(os.path.join(current_pdf_parent_directory, combinefile_list[i]))
            for file in combinefile_list_ab:
                if is_pdf_open(file)==True:
                    message('Please close the pdf while combining',self.ui)
                    return
            if is_pdf_open(combine_pdf_dir) == True:
                message('Please close the pdf while combining',self.ui)
                return
            self.progress_dialog = ProgressDialog("Combining service files, open in Bluebeam when done.",self.ui)
            self.combine_pro = Combine(self.progress_dialog,combinefile_list_ab, combine_pdf_dir)
            self.combine_pro.start()
            self.progress_dialog.exec_()
            open_in_bluebeam(combine_pdf_dir)
        except:
            traceback.print_exc()
        finally:
            global processing_or_not
            processing_or_not = False
    '''Overlay - Rescale service (to do 1)'''
    def f_button_overlay_rescale(self):
        try:
            '''get file list and dir'''
            file_list=self.get_table_files("Overlay service")
            file_list_ab=[]
            for i in range(len(file_list)):
                file_list_ab.append(os.path.join(self.overlay_service1_parentdir, file_list[i]))
            '''get quo_no and group selections'''
            quo_no = str(self.ui.text_all_1_projectnumber.text())
            id1=self.button_group_overlay_service1_os.checkedId()
            id2 = self.button_group_overlay_service1_orisize.checkedId()
            id3 = self.button_group_overlay_service1_outs.checkedId()
            id4 = self.button_group_overlay_service1_outsize.checkedId()
            id_list=[id1,id2,id3,id4]
            cus1=str(self.ui.text_overlay_service1_os.text())
            cus2=str(self.ui.text_overlay_service_orisize1.text())
            cus3=str(self.ui.text_overlay_service_orisize2.text())
            cus4=str(self.ui.text_overlay_service1_outs.text())
            cus5=str(self.ui.text_overlay_service_outsize1.text())
            cus6=str(self.ui.text_overlay_service_outsize2.text())
            cus_list=[cus1,cus2,cus3,cus4,cus5,cus6]
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            '''confirm right content'''
            if quo_no == '':
                message('Please Enter a project before resize',self.ui)
                return
            if not all([f.endswith(".pdf") or f.endswith(".PDF") for f in file_list_ab]):
                message('Please input at least one pdf file in the dialog box',self.ui)
                return
            output_file = self.ui.text_overlay_outfile.text()
            if len(file_list_ab) == 1:
                input_file = file_list_ab[0]
                if is_pdf_open(input_file) == True:
                    message('Please close the pdf while rescaling',self.ui)
                    return
            else:
                combine_pdf_dir =output_file
                for file in file_list_ab:
                    if is_pdf_open(file) == True:
                        message('Please close the pdf while rescaling',self.ui)
                        return
                if is_pdf_open(combine_pdf_dir) == True:
                    message('Please close the pdf while rescaling',self.ui)
                    return
                combine_pdf(file_list_ab, combine_pdf_dir)
                flatten_pdf(combine_pdf_dir,combine_pdf_dir)
                input_file = combine_pdf_dir
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            '''get process size change and scale change'''
            input_scale_list = ['50', '75', '100', '150', '200', 'custom']
            input_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24','A0-26', 'custom']
            output_scale_list = ['50', '100', 'custom']
            output_size_list = ['A4', 'A3', 'A2', 'A1', 'A0', 'A0-24','A0-26', 'custom']
            if id_list[0] == 6:
                input_scale = cus_list[0]
            else:
                input_scale = input_scale_list[id_list[0]-1]
            if id_list[2] == 3:
                output_scale = cus_list[3]
            else:
                output_scale = output_scale_list[id_list[2]-1]
            page_size_json_dir = os.path.join(conf["database_dir"], "page_size.json")
            all_page_size = json.load(open(page_size_json_dir))
            if id_list[1] == 8:
                if not (is_float(cus_list[1]) and is_float(cus_list[2])):
                    message('Please enter correct input size width and height',self.ui)
                    return
                input_size_x=convert_mm_to_pixel(float(cus_list[1]))
                input_size_y=convert_mm_to_pixel(float(cus_list[2]))
            else:
                size_input=input_size_list[id_list[1]-1]
                input_size_x = convert_mm_to_pixel(all_page_size[size_input]["width"])
                input_size_y = convert_mm_to_pixel(all_page_size[size_input]["height"])

            if id_list[3] == 8:
                if not (is_float(cus_list[4]) and is_float(cus_list[5])):
                    message('Please enter correct input size width and height',self.ui)
                    return
                output_size_x=convert_mm_to_pixel(float(cus_list[4]))
                output_size_y=convert_mm_to_pixel(float(cus_list[5]))

            else:
                size_output=output_size_list[id_list[3]-1]
                output_size_x = convert_mm_to_pixel(all_page_size[size_output]["width"])
                output_size_y = convert_mm_to_pixel(all_page_size[size_output]["height"])
        except Exception as e:
            traceback.print_exc()
            raise e
        try:
            size_list=[input_scale, input_size_x, input_size_y,output_scale, output_size_x, output_size_y]
            output_file = self.ui.text_overlay_outfile.text()
            if is_pdf_open(output_file) == True:
                message('Please close the pdf while rescaling',self.ui)
                return
            self.progress_dialog = ProgressDialog("Rescaling, open in Bluebeam when done.",self.ui)
            self.process = Process(self.progress_dialog, input_file,size_list,output_file)
            self.process.start()
            self.progress_dialog.exec_()
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            global processing_or_not
            processing_or_not = False

    '''=======================Kitchen======================='''
    def f_calculate_v_and_f(self,flow_rate,W,H,material):
        try:
            velocity=flow_rate/1000/W/H
            if material=='Masonry':
                e=1
            else:
                e=0.15
            Dh=4*1000*W*H/2/(W+H)
            Re=66.4*Dh*velocity
            Fd=0.25/(math.log10(e/3.7/Dh+5.74/(Re**0.9)))**2
            Fr=Fd*1.204*velocity*velocity*1000/2/Dh
            return [round(velocity,2),round(Fr,2)]
        except:
            traceback.print_exc()
            return False


    def f_text_kitchen_duct_flow(self):
        try:
            flow_rate_str=self.ui.text_kitchen_duct_flow.text()
            W_str=self.ui.text_kitchen_duct_size1.text()
            H_str=self.ui.text_kitchen_duct_size2.text()
            material=self.ui.combobox_kitchen_duct_material.currentText()
            if is_float(flow_rate_str) and is_float(W_str) and is_float(H_str):
                flow_rate=float(flow_rate_str)
                W=float(W_str)
                H=float(H_str)
                if flow_rate*W*H>0:
                    result=self.f_calculate_v_and_f(flow_rate,W,H,material)
                    if result==False:
                        self.ui.text_kitchen_duct_velocity.setText('')
                        self.ui.text_kitchen_duct_friction.setText('')
                    else:
                        velocity=result[0]
                        Fr=result[1]
                        self.ui.text_kitchen_duct_velocity.setText(str(velocity))
                        self.ui.text_kitchen_duct_friction.setText(str(Fr))
            else:
                self.ui.text_kitchen_duct_velocity.setText('')
                self.ui.text_kitchen_duct_friction.setText('')
        except:
            traceback.print_exc()


    def f_button_kitchen_duct_calculate(self):
        try:
            flow_rate_str=self.ui.text_kitchen_duct_flow.text()
            W_str=self.ui.text_kitchen_duct_size1.text()
            material=self.ui.combobox_kitchen_duct_material.currentText()
            if is_float(flow_rate_str) and is_float(W_str):
                flow_rate=float(flow_rate_str)
                W=float(W_str)
                if flow_rate*W>0:
                    for i in range(20):
                        H=0.1+0.05*i
                        result=self.f_calculate_v_and_f(flow_rate,W,H,material)
                        if result!=False:
                            velocity=result[0]
                            Fr=result[1]
                            if velocity<=7 and Fr<=1:
                                self.ui.text_kitchen_duct_size2.setText(str(round(H,2)))
                                self.ui.text_kitchen_duct_velocity.setText(str(velocity))
                                self.ui.text_kitchen_duct_friction.setText(str(Fr))
                                break
        except:
            traceback.print_exc()

    def f_text_kitchen_duct_flex_flow(self):
        try:
            flow_str=self.ui.text_kitchen_duct_flex_flow.text()
            dia=''
            if is_float(flow_str):
                flow=float(flow_str)
                if 0<flow<=45:
                    dia=150
                elif 45<flow<=80:
                    dia=200
                elif 80<flow<=150:
                    dia=250
                elif 150<flow<=220:
                    dia=300
                elif 220<flow<=330:
                    dia=350
                elif 330<flow<=425:
                    dia=400
                elif 425<flow<=540:
                    dia=450
                elif flow>540:
                    dia=500
            self.ui.text_kitchen_duct_flex_dia.setText(str(dia))
        except:
            traceback.print_exc()

    def f_button_kitchen_duct_save(self):
        try:
            flow_rate_str=self.ui.text_kitchen_duct_flow.text()
            W_str=self.ui.text_kitchen_duct_size1.text()
            H_str=self.ui.text_kitchen_duct_size2.text()
            if is_float(flow_rate_str) and is_float(W_str) and is_float(H_str):
                flow_rate = float(flow_rate_str)
                W = float(W_str)
                H = float(H_str)
                if flow_rate * W * H > 0:
                    table = self.ui.table_kitchen_duct
                    row_position = table.rowCount()
                    table.insertRow(row_position)
                    table.setItem(row_position, 0, QTableWidgetItem(flow_rate_str))
                    table.setItem(row_position, 1, QTableWidgetItem(W_str))
                    table.setItem(row_position, 2, QTableWidgetItem(H_str))
                    self.f_set_table_style(table)
        except:
            traceback.print_exc()


    '''=======================Timesheet======================='''
    def time_to_minutes(self,time_str):
        try:
            print(time_str)
            if time_str.find('m')!=-1:
                match = re.match(r'(\d+)h\s*(\d+)m', time_str)
                if match:
                    hours = int(match.group(1))
                    minutes = int(match.group(2))
                    return hours * 60 + minutes
                else:
                    match_minutes = re.match(r'(\d+)m', time_str)
                    if match_minutes:
                        return int(match_minutes.group(1))
                    else:
                        return 30
            else:
                if is_float(time_str):
                    return int(float(time_str))
                else:
                    return 30
        except:
            traceback.print_exc()
            return 30



    def f_connect_time(self):
        try:
            for i in range(5):
                time_daily_list=self.time_daily_dict[i]
                if time_daily_list!=[]:
                    for j in range(len(time_daily_list)):
                        est_text=time_daily_list[j][1]
                        act_text=time_daily_list[j][2]
                        # est_text.textChanged.connect(partial(self.f_calculate_pro_time,day=i+1))
                        # act_text.textChanged.connect(partial(self.f_calculate_pro_time,day=i+1))

        except:
            traceback.print_exc()


    def f_calculate_sum_time_all(self):
        try:
            self.f_calculate_pro_time(day=1)
            self.f_calculate_pro_time(day=2)
            self.f_calculate_pro_time(day=3)
            self.f_calculate_pro_time(day=4)
            self.f_calculate_pro_time(day=5)
            self.f_calculate_sum_time(day=1)
            self.f_calculate_sum_time(day=2)
            self.f_calculate_sum_time(day=3)
            self.f_calculate_sum_time(day=4)
            self.f_calculate_sum_time(day=5)
        except:
            traceback.print_exc()

    def f_calculate_sum_time(self,day):
        try:
            i=day-1
            total_est=0
            total_act=0
            est_general_text=self.ui.findChild(QLineEdit,'text_timesheet_general_est_'+str(i+1))
            act_general_text = self.ui.findChild(QLineEdit, 'text_timesheet_general_act_' + str(i + 1))
            est_other_text=self.ui.findChild(QLineEdit,'text_timesheet_other_est_'+str(i+1))
            act_other_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_act_' + str(i + 1))
            est_project_text=self.ui.findChild(QLineEdit,'text_timesheet_project_est_'+str(i+1))
            act_project_text = self.ui.findChild(QLineEdit, 'text_timesheet_project_act_' + str(i + 1))
            if is_float(est_general_text.text()):
                total_est+=int(est_general_text.text())
            if is_float(est_other_text.text()):
                total_est+=int(est_other_text.text())
            if is_float(est_project_text.text()):
                total_est+=int(est_project_text.text())
            if is_float(act_general_text.text()):
                total_act+=int(act_general_text.text())
            if is_float(act_other_text.text()):
                total_act+=int(act_other_text.text())
            if is_float(act_project_text.text()):
                total_act+=int(act_project_text.text())
            est_total_text = self.ui.findChild(QLineEdit, 'text_timesheet_total_est_' + str(i + 1))
            act_total_text = self.ui.findChild(QLineEdit, 'text_timesheet_total_act_' + str(i + 1))
            est_total_text.setText(str(total_est))
            act_total_text.setText(str(total_act))
        except:
            traceback.print_exc()

    def f_calculate_pro_time(self,day):
        try:
            time_daily_list=self.time_daily_dict[day-1]
            total_project_est=0
            total_project_act=0
            if time_daily_list!=[]:
                for time_daily_list_i in time_daily_list:
                    est_text=time_daily_list_i[1]
                    act_text=time_daily_list_i[2]
                    if est_text.text()!='':
                        total_project_est+=int(est_text.text())
                    if act_text.text()!='':
                        total_project_act+=int(act_text.text())
            est_pro_text = self.ui.findChild(QLineEdit, 'text_timesheet_project_est_' + str(day))
            act_pro_text = self.ui.findChild(QLineEdit, 'text_timesheet_project_act_' + str(day))
            est_pro_text.setText(str(total_project_est))
            act_pro_text.setText(str(total_project_act))
        except:
            traceback.print_exc()

    def f_button_timesheet_syncback(self):
        try:
            email=self.time_sync_info_list[0]
            asana_id_to_sync=self.time_sync_info_list[1]
            for i in range(5):
                date_today=self.time_sync_info_list[2][i]
                est_time_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_est_' + str(i + 1))
                act_time_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_act_' + str(i + 1))
                description_text = self.ui.findChild(QTextEdit, 'text_timesheet_des_' + str(i + 1))
                if is_float(est_time_text.text()):
                    est_time=int(est_time_text.text())
                else:
                    est_time =0
                if is_float(act_time_text.text()):
                    act_time=int(act_time_text.text())
                else:
                    act_time =0
                description=description_text.toPlainText()
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("DELETE FROM bridge.general_time_with_description WHERE user_email = %s AND date=%s",(email,date_today,))
                    mycursor.execute("INSERT INTO bridge.general_time_with_description VALUES (%s, %s, %s, %s, %s)",(date_today,email,est_time,act_time,description))
                    mydb.commit()
                except:
                    traceback.print_exc()
                finally:
                    mycursor.close()
                    mydb.close()

                try:
                    timesheet_list=self.time_daily_dict[i]
                    for task_time_list in timesheet_list:
                        asana_url=task_time_list[0]
                        asana_id=asana_url.split('/')[-1]
                        est_time=task_time_list[1].text()
                        act_time=task_time_list[2].text()
                        if not is_float(est_time):
                            est_time=0
                        update_asana_estimated_time(asana_id, int(est_time))
                        if not is_float(act_time):
                            act_time=0
                        update_asana_actual_time(asana_id, int(act_time))
                except:
                    traceback.print_exc()
                    message('Sync to database and Asana unsuccessfully',self.ui)
                    return

            message('Sync to database and Asana successfully',self.ui)
        except:
            traceback.print_exc()

    def get_weekday_list(self,selected_date):
        try:
            day_of_week = selected_date.dayOfWeek()
            monday = selected_date.addDays(-day_of_week + 1)
            tuesday = monday.addDays(1)
            wednesday = monday.addDays(2)
            thursday = monday.addDays(3)
            friday = monday.addDays(4)
            return [monday.toString('yyyy-MM-dd'),tuesday.toString('yyyy-MM-dd'),wednesday.toString('yyyy-MM-dd'),
                    thursday.toString('yyyy-MM-dd'),friday.toString('yyyy-MM-dd')]

        except:
            traceback.print_exc()


    def f_button_timesheet_sync(self):
        try:
            table=self.ui.table_timesheet_user
            if table.rowCount() == 0:
                message('Please login',self.ui)
                return
            elif table.rowCount()==1:
                row=0
            else:
                selected_items = table.selectedItems()
                if selected_items:
                    row = selected_items[0].row()
                else:
                    message('Please choose name',self.ui)
                    return
            email=table.item(row,0).text()
            asana_id_to_sync=table.item(row,2).text()
            qdate = self.ui.calender_timesheet.selectedDate()
            selected_date = date(qdate.year(), qdate.month(), qdate.day())
            weekday_list = self.get_weekday_list(QDate(qdate.year(), qdate.month(), qdate.day()))
            self.time_sync_info_list=[email,asana_id_to_sync,weekday_list]
            frame_list=[self.Frame_Mon,self.Frame_Tue,self.Frame_Wed,self.Frame_Thu,self.Frame_Fri]
            layout_list=[self.layout_Frame_Mon,self.layout_Frame_Tue,self.layout_Frame_Wed,self.layout_Frame_Thu,self.layout_Frame_Fri]
            self.time_daily_dict={0:[],1:[],2:[],3:[],4:[]}


            for i in range(5):
                for child in frame_list[i].children():
                    if isinstance(child, QWidget):
                        frame_list[i].layout().removeWidget(child)
                        child.deleteLater()
                for j in reversed(range(layout_list[i].count())):
                    item = layout_list[i].itemAt(j)
                    if isinstance(item, QSpacerItem):
                        layout_list[i].takeAt(j)

                est_time_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_est_' + str(i + 1))
                act_time_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_act_' + str(i + 1))
                description_text = self.ui.findChild(QTextEdit, 'text_timesheet_des_' + str(i + 1))
                est_time_text.setText('')
                act_time_text.setText('')
                description_text.setText('')

                daily_general_est=0
                daily_general_act=0
                data = get_asana_tasks_by_user_by_date(email, datetime.strptime(weekday_list[i], "%Y-%m-%d"))
                if data!=None:
                    sorted_data = dict(sorted(data.items(), key=lambda x: int(x[0].split("/")[-1])))
                    for permalink_url, item in sorted_data.items():
                        task_name=item['name']
                        pro_name=item['parent']
                        act_time_get=item['actual_time']
                        asana_url=permalink_url
                        if pro_name is not None:
                            if act_time_get==None:
                                act_time=0
                            else:
                                act_time=self.time_to_minutes(str(act_time_get))
                            est_time_get=item['estimated_time']
                            if est_time_get==None:
                                est_time=30
                            else:
                                est_time = self.time_to_minutes(str(est_time_get))
                            if pro_name.startswith('P:'):
                                grid_layout = QGridLayout()
                                lineedit1_1=QLineEdit("Task:")
                                lineedit1_1.setMaximumWidth(30)
                                lineedit1_2=QLineEdit(task_name)
                                lineedit1_2.setAlignment(Qt.AlignLeft)
                                lineedit2_1=QLineEdit("Pro:")
                                lineedit2_1.setMaximumWidth(30)
                                lineedit2_2=QLineEdit(pro_name)
                                lineedit2_2.setAlignment(Qt.AlignLeft)
                                lineedit3_1=QLineEdit("Est:")
                                lineedit3_1.setMaximumWidth(30)
                                lineedit3_2 = QLineEdit(str(est_time))
                                lineedit3_3 = QLineEdit("Act:")
                                lineedit3_3.setMaximumWidth(30)
                                lineedit3_4 = QLineEdit(str(act_time))
                                grid_layout.addWidget(lineedit1_1, 0, 0)
                                grid_layout.addWidget(lineedit1_2, 0, 1, 1, 3)
                                grid_layout.addWidget(lineedit2_1, 1, 0)
                                grid_layout.addWidget(lineedit2_2, 1, 1, 1, 3)
                                grid_layout.addWidget(lineedit3_1, 2, 0)
                                grid_layout.addWidget(lineedit3_2, 2, 1)
                                grid_layout.addWidget(lineedit3_3, 2, 2)
                                grid_layout.addWidget(lineedit3_4, 2, 3)
                                grid_container = QWidget()
                                grid_container.setLayout(grid_layout)
                                grid_container.setFixedHeight(80)
                                grid_container.setStyleSheet("border: 1px solid white; ")
                                for widget in grid_container.findChildren(QWidget):
                                    widget.setStyleSheet("border: none;")
                                if est_time_get==None:
                                    lineedit3_2.setStyleSheet("background-color: rgb(240, 128, 128);")
                                else:
                                    lineedit3_2.setStyleSheet("background-color: white;")
                                if act_time_get == None:
                                    lineedit3_4.setStyleSheet("background-color: rgb(240, 128, 128);")
                                else:
                                    lineedit3_4.setStyleSheet("background-color: white;")

                                layout_list[i].addWidget(grid_container)

                                task_time_list=[asana_url,lineedit3_2,lineedit3_4]
                                self.time_daily_dict[i].append(task_time_list)
                            else:
                                daily_general_est+=est_time
                                daily_general_act+=act_time
                    self.ui.findChild(QLineEdit, 'text_timesheet_general_est_' + str(i + 1)).setText(str(daily_general_est))
                    self.ui.findChild(QLineEdit,'text_timesheet_general_act_'+str(i+1)).setText(str(daily_general_act))
                    spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
                    layout_list[i].addItem(spacer)



                date_today=weekday_list[i]
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("SELECT * FROM bridge.general_time_with_description WHERE user_email = %s AND date=%s",
                                     (email, date_today,))
                    general_time_database = mycursor.fetchall()
                    if len(general_time_database) == 1:
                        est_time=general_time_database[0][2]
                        act_time=general_time_database[0][3]
                        description=general_time_database[0][4]
                        est_time_text=self.ui.findChild(QLineEdit,'text_timesheet_other_est_'+str(i+1))
                        act_time_text = self.ui.findChild(QLineEdit, 'text_timesheet_other_act_' + str(i + 1))
                        description_text = self.ui.findChild(QTextEdit, 'text_timesheet_des_' + str(i + 1))
                        est_time_text.setText(str(est_time))
                        act_time_text.setText(str(act_time))
                        description_text.setText(str(description))
                except:
                    traceback.print_exc()
                finally:
                    mycursor.close()
                    mydb.close()
            self.f_connect_time()
            self.f_calculate_pro_time(day=1)
            self.f_calculate_pro_time(day=2)
            self.f_calculate_pro_time(day=3)
            self.f_calculate_pro_time(day=4)
            self.f_calculate_pro_time(day=5)
            self.f_calculate_sum_time(day=1)
            self.f_calculate_sum_time(day=2)
            self.f_calculate_sum_time(day=3)
            self.f_calculate_sum_time(day=4)
            self.f_calculate_sum_time(day=5)
            message('Sync successfully',self.ui)
        except:
            traceback.print_exc()
            message('Sync failed',self.ui)



if __name__ == '__main__':
    global link_path
    link_path=''
    # global processing_or_not
    # processing_or_not=False
    multiprocessing.freeze_support()
    app = QApplication(sys.argv)
    ui = uic.loadUi(r"T:\00-Template-Do Not Modify\00-Bridge template\ui\bluebeam_copilot.ui")
    layout_copilot_pro_top = ui.findChild(QGridLayout, 'gridLayout_12')
    tab_copilot = ui.findChild(QTabWidget, 'tabWidget')

    ui2=uic.loadUi(r"T:\00-Template-Do Not Modify\00-Bridge template\ui\Copilot-Bridge.ui")
    frame_bridge_top_1 = ui2.findChild(QFrame, 'frame_pro_stage')
    frame_bridge_top_2 = ui2.findChild(QFrame, 'frame_pro_func')
    frame_bridge_top_3 = ui2.findChild(QFrame, 'frame_pro_invoice')

    tab_bridge_1 = ui2.findChild(QWidget, 'tab_project_info')
    tab_bridge_2 = ui2.findChild(QWidget, 'tab_fee')
    tab_bridge_3 = ui2.findChild(QWidget, 'tab_finan')

    row = layout_copilot_pro_top.rowCount() - 1
    layout_copilot_pro_top.addWidget(frame_bridge_top_1, row, layout_copilot_pro_top.columnCount())
    layout_copilot_pro_top.addWidget(frame_bridge_top_2, row, layout_copilot_pro_top.columnCount())
    layout_copilot_pro_top.addWidget(frame_bridge_top_3, row, layout_copilot_pro_top.columnCount())
    tab_copilot.addTab(tab_bridge_1, "Project Info")
    tab_copilot.addTab(tab_bridge_2, "Fee Details")
    tab_copilot.addTab(tab_bridge_3, "Financial Panel")

    stats = Stats(ui)
    stats2=Stats_bridge(ui2,stats)
    # stats.set_quotation_number("000000AA")
    stats.load_project("000000AA")


    app.aboutToQuit.connect(stats.on_close)
    stats.ui.show()
    ui.showMaximized()
    # app.exec_()
    sys.exit(app.exec_())

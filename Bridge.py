from PyQt5.QtWidgets import (QApplication, QTableWidget, QTableWidgetItem,QHeaderView,QButtonGroup, QCheckBox, QMessageBox,QVBoxLayout,QLabel,
                             QProgressBar,QDialog,QRadioButton,QPushButton,QFileDialog,QDialogButtonBox,QTabWidget, QFrame,QLineEdit,QComboBox)
from PyQt5.QtCore import Qt,QThread, pyqtSignal, QLoggingCategory
from PyQt5.QtGui import QPixmap,QColor, QFont
import shutil
import os
from pathlib import Path
import threading
from datetime import date,datetime, timedelta
import traceback
import subprocess
from PyQt5.QtNetwork import QNetworkAccessManager
import time
from functools import partial
from utility.asana_function import (create_asana_project, update_asana_project, update_asana_email,
                                    get_asana_sub_tasks, update_asana_sub_task, get_asana_email, create_asana_invoice,
                                    update_asana_invoice, create_asana_bill, update_asana_bill, update_asana_project_tags)
from utility.email_function import convert_time_format,f_email_fee,f_email_invoice
import mysql.connector
from conf.conf_bridge import CONFIGURATION as conf
import psutil
from utility.utility_bridge import (open_folder, open_link_with_edge, file_exists, is_float, create_directory,
                                    get_first_name, check_or_create_folder, delete_project_database, open_pdf_thread,
                                    char2num, change_quotation_number, export_excel_thread, f_write_invoice_excel, f_write_minor_fee_excel, f_write_major_fee_excel,
                                    f_write_installation_fee_excel, is_valid_email, convert_date, f_get_client_xero_id, check_excel_open)
from utility.xero_function import (login_xero2, refresh_token, create_xero_invoice, create_xero_bill, get_xero_invoice_status,
                                   update_xero_invoice, update_xero_bill, get_xero_bill_status, create_xero_contact, update_xero_contact, update_xero_invoice_contact)
import re
from win32com import client as win32client
import Levenshtein
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
    def __init__(self,info):
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
        for i in range(100):
            if not self.running:
                break
            time.sleep(self.sumtime/100)
            self.progress.emit(i)
        while self.running:
            pass
    def stop(self):
        self.running = False
'''Project - Email fee proposal - Waiting'''
class Wait_email_fee(QThread):
    sinout1=pyqtSignal(str)
    def __init__(self,self_main,parent=None):
        super(Wait_email_fee,self).__init__(parent)
        self.self_main=self_main
    def run(self):
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            quotation_number = self.self_main.ui.text_proinfo_1_quonum.text()
            print(quotation_number)
            mycursor.execute("SELECT * FROM bridge.emails WHERE quotation_number = %s", (quotation_number,))
            fee_revision_old = mycursor.fetchall()[0][6]
            mycursor.close()
            mydb.close()
            print(fee_revision_old)
            for i in range(60):
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.emails WHERE quotation_number = %s", (quotation_number,))
                fee_revision = mycursor.fetchall()[0][6]
                mycursor.close()
                mydb.close()
                print(fee_revision,i)
                if fee_revision!=fee_revision_old:
                    self.sinout1.emit(quotation_number)
                    return
                time.sleep(5)
        except:
            traceback.print_exc()

'''Project - Sofware logout'''
class Software_logout(QThread):
    sinout3=pyqtSignal(list)
    def __init__(self,self_main,parent=None):
        super(Software_logout,self).__init__(parent)
        self.self_main=self_main
    def run(self):
        try:
            self.stop_thread=False
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            user_email=self.self_main.parent.ui.text_user_email.text()
            action_time = str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            quo_num=self.self_main.ui.text_proinfo_1_quonum.text()
            pro_num=self.self_main.ui.text_proinfo_1_pronum.text()
            pro_name=self.self_main.ui.text_proinfo_1_proname.text()
            pro_info_list=[quo_num,pro_num,pro_name]
            mycursor.execute("SELECT * FROM bridge.last_update WHERE quotation_number = %s", (quo_num,))
            invoices_database = mycursor.fetchall()
            if len(invoices_database)>0:
                mycursor.execute("UPDATE bridge.last_update SET user_email = %s, last_update_time=%s WHERE quotation_number=%s",
                                 (user_email, action_time,quo_num,))
            else:
                mycursor.execute("INSERT INTO bridge.last_update VALUES (%s, %s, %s)",
                                 (quo_num,user_email, action_time,))
            mydb.commit()
        except:
            traceback.print_exc()
        finally:
            mycursor.close()
            mydb.close()
        try:
            for i in range(1800):
                time.sleep(1)
                if self.stop_thread:
                    break
                if i==1799:
                    self.sinout3.emit(pro_info_list)
        except:
            traceback.print_exc()
    def stop_thread_func(self):
        self.stop_thread = True

'''Project - Sofware logout'''

class Refresh_table(QThread):
    sinout4 = pyqtSignal(str)
    def __init__(self, self_main, parent=None):
        super(Refresh_table, self).__init__(parent)
        self.self_main = self_main
    def run(self):
        try:
            self.self_main.parent.generate_search_bar(order=1)
            self.sinout4.emit('ok')
        except:
            traceback.print_exc()

'''Finan - Email invoice - Waiting'''
class Wait_email_inv(QThread):
    sinout2=pyqtSignal(str)
    def __init__(self,self_main,parent=None):
        super(Wait_email_inv,self).__init__(parent)
        self.self_main=self_main
    def run(self):
        try:
            quo_num=self.self_main.ui.text_proinfo_1_quonum.text()
            color_dict = {"#d2d2d2": "Backlog", "#ff0000": "Sent", "#008000": "Paid", "#800080": "Voided", }
            inv_track_list=[]
            for i in range(8):
                invoice_number_text = self.self_main.parent.ui.findChild(QLineEdit, "text_finan_1_inv" + str(i + 1))
                if invoice_number_text.text()!='' and invoice_number_text.text().find('-')==-1:
                    palette = invoice_number_text.palette()
                    background_color = palette.color(palette.Window).name()
                    state = color_dict[background_color]
                    if state=='Backlog':
                        inv_track_list.append(invoice_number_text.text())
            if inv_track_list==[]:
                return
        except:
            traceback.print_exc()
        try:
            for i in range(60):
                for invoice in inv_track_list:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("SELECT * FROM bridge.invoices WHERE invoice_number = %s", (invoice,))
                    inv_database=mycursor.fetchall()
                    mycursor.close()
                    mydb.close()
                    print(i,inv_database)
                    if len(inv_database) == 1:
                        inv_database = inv_database[0]
                        state_new=inv_database[2]
                        if state_new=='Sent':
                            self.sinout2.emit(quo_num)
                            return
                time.sleep(5)
        except:
            traceback.print_exc()
'''Project - Design certificate dialog'''
class CustomDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Design Certificate Selection")
        self.setGeometry(150, 150, 250, 150)
        layout = QVBoxLayout(self)
        self.radio1 = QRadioButton("Design Certificate")
        self.radio2 = QRadioButton("Design Compliance Certificate")
        self.radio1.setChecked(True)
        layout.addWidget(self.radio1)
        layout.addWidget(self.radio2)
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        if parent:
            screen = parent.screen()
            screen_geometry = screen.availableGeometry()
            message_geometry = self.frameGeometry()
            message_geometry.moveCenter(screen_geometry.center())
            self.move(message_geometry.topLeft())

    def get_selected_option(self):
        if self.radio1.isChecked():
            return 'Design'
        elif self.radio2.isChecked():
            return 'Compliance'
        return "None"

'''=================================Main function===================================='''
class Stats_bridge():
    def __init__(self,ui,parent):
        super().__init__()
        self.ui=ui
        self.parent=parent
        self.search_bar_update = True
        '''ui dict'''
        service_list_all = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis', 'var']
        item_name_list_all = ["mech_1", "mech_2", "mech_3", "cfd_1", "cfd_2", "cfd_3", "ele_1", "ele_2", "ele_3",
                              "hyd_1", "hyd_2", "hyd_3","fire_1", "fire_2", "fire_3", "install_1", "install_2", "install_3", "mis_1", "mis_2",
                              "mis_3","mechrev_1", "mechrev_2", "mechrev_3", "var_1", "var_2", "var_3"]
        self.ui_name_dict = {}
        for service in service_list_all:
            for i in range(4):
                button_name = 'button_finan_2_' + service + '_upload' + str(i + 1)
                self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
            for i in range(3):
                button_name = 'button_finan_2_' + service + '_upload' + str(i + 5)
                self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
        for i in range(8):
            button_name = 'button_finan_4_preview' + str(i + 1)
            self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
            button_name = 'button_finan_4_remitfull' + str(i + 1)
            self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
            button_name = "button_finan_4_email" + str(i + 1)
            self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
            for j in range(3):
                button_name = 'button_finan_4_remit' + str(i + 1) + '_' + str(j + 1)
                self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
        for service in service_list_all[:-1]:
            button_name = "button_fee_3_savesql_" + service
            self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
            for i in range(3):
                button_name = "button_fee_" + service + '_up_' + str(i + 1)
                self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
                button_name = "button_fee_" + service + '_down_' + str(i + 1)
                self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
                button_name = "button_fee_" + service + '_add_' + str(i + 1)
                self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)
        for i in range(4):
            button_name = "button_fee_stage_" + str(i + 1)
            self.ui_name_dict[button_name] = self.parent.ui.findChild(QPushButton, button_name)

        for i in range(8):
            radiobutton_name = "radbuttton_proinfo_1_projecttype_" + str(i + 1)
            self.ui_name_dict[radiobutton_name] = self.parent.ui.findChild(QRadioButton, radiobutton_name)
        for i in range(12):
            radiobutton_name = "radbutton_proinfo_2_clientcontacttype_" + str(i + 1)
            self.ui_name_dict[radiobutton_name] = self.parent.ui.findChild(QRadioButton, radiobutton_name)
            radiobutton_name = "radbutton_proinfo_2_contactcontacttype_" + str(i + 1)
            self.ui_name_dict[radiobutton_name] = self.parent.ui.findChild(QRadioButton, radiobutton_name)
        for service in service_list_all:
            for i in range(4):
                for j in range(8):
                    radiobutton_name = "radbutton_finan_1_" + service + '_' + str(i + 1) + '_' + str(j + 1)
                    self.ui_name_dict[radiobutton_name] = self.parent.ui.findChild(QRadioButton, radiobutton_name)

        for service in service_list_all[:-1]:
            frame_name = "frame_fee_scope_" + service
            self.ui_name_dict[frame_name] = self.parent.ui.findChild(QFrame, frame_name)
        for service in service_list_all:
            frame_name = "frame_fee_pro_" + service
            self.ui_name_dict[frame_name] = self.parent.ui.findChild(QFrame, frame_name)
            frame_name = "frame_finan_1_" + service
            self.ui_name_dict[frame_name] = self.parent.ui.findChild(QFrame, frame_name)
            frame_name = "frame_finan_2_" + service
            self.ui_name_dict[frame_name] = self.parent.ui.findChild(QFrame, frame_name)
            frame_name = "frame_finan_3_" + service
            self.ui_name_dict[frame_name] = self.parent.ui.findChild(QFrame, frame_name)

        for item_name in item_name_list_all:
            text_name = "text_finan_2_" + item_name + "_1"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + item_name + "_2"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + item_name + "_3"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + item_name + "_4"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + item_name + "_5"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        for service in service_list_all:
            text_name = "text_finan_1_" + service + "_fee0_1"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_" + service + "_fee0_2"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_1"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_2"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_3"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_4"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_5"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_6"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_0_7"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_4_1"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_2_" + service + "_4_2"
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            for i in range(4):
                text_name = "text_finan_2_" + service + "_0_" + str(i + 2)
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            for i in range(3):
                text_name = "text_finan_2_" + service + '_' + str(i + 1) + '_2'
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
                text_name = "text_finan_2_" + service + '_' + str(i + 1) + '_1'
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            for i in range(4):
                text_name = "text_finan_1_" + service + '_doc' + str(i + 1)
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
                text_name = "text_finan_1_" + service + '_fee' + str(i + 1) + '_1'
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
                text_name = "text_finan_1_" + service + '_fee' + str(i + 1) + '_2'
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_3_" + service
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_3_" + service + '_in'
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        for i in range(8):
            text_name = "text_finan_1_invclient_" + str(i + 1)
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_total" + str(i + 1) + '_1'
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_total" + str(i + 1) + '_2'
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_total" + str(i + 1) + '_3'
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_total" + str(i + 1) + '_4'
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_total" + str(i + 1) + '_5'
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_inv" + str(i + 1)
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            text_name = "text_finan_1_total1_" + str(i + 1)
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)
            for j in range(3):
                text_name = "text_finan_4_remit_" + str(i + 1) + '_'+str(j+1)
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        ui_text_list = ['name', 'company', 'adress', 'abn', 'phone', 'email']
        type_list = ['client', 'contact']
        for text_item_name in ui_text_list:
            for text_type in type_list:
                text_name = "text_proinfo_2_" + text_type + text_item_name
                self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        for i in range(3):
            text_name="text_fee_2_period" + str(i+1)
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        for i in range(4):
            text_name = "text_stage_table_name" + str(i + 1)
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        for i in range(3):
            text_name="text_finan_5_note" + str(i + 1)
            self.ui_name_dict[text_name] = self.parent.ui.findChild(QLineEdit, text_name)

        for i in range(4):
            table_name = "table_fee_stage_" + str(i + 1)
            self.ui_name_dict[table_name] = self.parent.ui.findChild(QTableWidget, table_name)
        for service in service_list_all[:-1]:
            for j in range(3):
                table_name = "table_fee_3_" + service + "_" + str(j + 1)
                self.ui_name_dict[table_name] = self.parent.ui.findChild(QTableWidget, table_name)
        for service in service_list_all:
            table_name = "table_fee_4_" + service
            self.ui_name_dict[table_name] = self.parent.ui.findChild(QTableWidget, table_name)

        for i in range(9):
            checkbox_name = "checkbox_proinfo_1_servicettype_" + str(i + 1)
            self.ui_name_dict[checkbox_name] = self.parent.ui.findChild(QCheckBox, checkbox_name)
        for item_name in item_name_list_all:
            checkbox_name = "checkbox_finan_2_" + item_name
            self.ui_name_dict[checkbox_name] = self.parent.ui.findChild(QCheckBox, checkbox_name)
        for service in service_list_all:
            checkbox_name = "checkbox_finan_2_" + service + "_0"
            self.ui_name_dict[checkbox_name] = self.parent.ui.findChild(QCheckBox, checkbox_name)
        for i in range(4):
            checkbox_name = "checkbox_fee_stage_" + str(i + 1)
            self.ui_name_dict[checkbox_name] = self.parent.ui.findChild(QCheckBox, checkbox_name)

        for item_name in item_name_list_all:
            combobox_name="combobox_finan_2_" + item_name
            self.ui_name_dict[combobox_name]=self.parent.ui.findChild(QComboBox, combobox_name)


        '''Last state'''
        self.last_state=''
        self.search_times=0
        '''Tab widget position set'''
        self.ui.tabWidget.setTabPosition(QTabWidget.South)
        '''Network manager'''
        self.manager = QNetworkAccessManager(self.ui)
        '''Lock function'''
        self.lock = threading.Lock()
        '''Press button set'''
        self.button_to_function_1 = {
            self.ui.button_pro_search: self.f_button_pro_search,
            self.ui.button_pro_func_openfolder:partial(self.f_profolder_open, folder_name="Project"),
            self.ui.button_pro_func_backup:partial(self.f_profolder_open, folder_name="Backup"),
            self.ui.button_pro_func_database:partial(self.f_profolder_open, folder_name="Database"),
            self.ui.button_fee_1_today:partial(self.f_gettodaydate, text_name="Service"),
            self.ui.button_fee_3_install_today:partial(self.f_gettodaydate, text_name="Install"),
            self.ui.button_pro_func_openasana:self.f_button_pro_func_openasana,
            self.ui.button_pro_func_renameproject:self.f_button_pro_func_renameproject,
            self.ui.button_fee_lock:partial(self.f_button_lock, lock_name="Fee"),
            self.ui.button_finan_1_lock:partial(self.f_button_lock, lock_name="Financial"),
            self.ui.button_proinfo_deletepro:self.f_button_proinfo_deletepro,
            self.ui.button_pro_state_setup:self.f_button_pro_state_setup,
            self.ui.button_pro_func_updateasana:partial(self.f_button_pro_func_updateasana_button,update_sub=False),
            self.ui.button_pro_state_emailfee:self.f_button_pro_state_emailfee,
            self.ui.button_pro_state_chasefee:self.f_button_pro_state_chasefee,
            self.ui.button_pro_func_designcert:self.f_button_pro_func_designcert,
            self.ui.button_pro_state_genfee:self.f_button_pro_state_genfee,
            self.ui.button_pro_clear:partial(self.f_clear2default,save_or_not=True),
            self.ui.button_fee_3_install_preview:self.f_button_fee_3_install_preview,
            self.ui.button_fee_3_install_email:self.f_button_fee_3_install_email,
            self.ui.button_finan_5_upload:self.f_button_finan_5_upload,
            self.ui.button_finan_5_confirm:self.f_button_finan_5_confirm,
            self.ui.button_finan_1_gen:self.f_button_finan_1_gen,
            self.ui.button_finan_1_del: self.f_button_finan_1_del,
            self.ui.button_pro_func_loginxero:self.f_button_pro_func_loginxero,
            self.ui.button_pro_func_refreshxero:self.f_button_pro_func_refreshxero,
            self.ui.button_pro_func_updatexero:self.f_button_pro_func_updatexero,
            self.ui.button_proinfo_search_unlock1:partial(self.f_button_proinfo_search_unlock,contact_type='client'),
            self.ui.button_proinfo_search_unlock2: partial(self.f_button_proinfo_search_unlock, contact_type='contact'),
            self.ui.button_proinfo_search_update1:partial(self.f_button_proinfo_search_update,contact_type='client'),
            self.ui.button_proinfo_search_update2: partial(self.f_button_proinfo_search_update, contact_type='contact'),
            self.ui.button_proinfo_search_create1: partial(self.f_button_proinfo_search_create,contact_type='client'),
            self.ui.button_proinfo_search_create2:  partial(self.f_button_proinfo_search_create, contact_type='contact'),
            self.ui.toolbutton_fee_revision_add:self.f_toolbutton_fee_revision_add,
            self.ui.button_pro_checksimilar:self.f_button_pro_checksimilar,
            self.ui.button_pro_func_syncxero:self.f_button_pro_func_syncxero,
        }

        self.button_to_function_2={}

        self.f_button_proinfo_search_unlock(contact_type='client')
        self.f_button_proinfo_search_unlock(contact_type='contact')


        for service in service_list_all:
            for i in range(4):
                button_name='button_finan_2_'+service+'_upload' + str(i+1)
                button=self.ui_name_dict[button_name]
                self.button_to_function_2[button]=partial(self.f_button_upload_subconsultant_invoice, service_name=service,id=str(i+1))
            for i in range(3):
                button_name="button_finan_2_"+ service + '_upload' + str(i + 5)
                button=self.ui_name_dict[button_name]
                self.button_to_function_2[button]=partial(self.f_finan_bill_upload, service=service, id=i + 1)

        self.button_to_function_3={}
        for i in range(8):
            button_name='button_finan_4_preview'+str(i+1)
            button=self.ui_name_dict[button_name]
            self.button_to_function_3[button]=partial(self.f_invoice_preview,id=str(i+1))

            button_name='button_finan_4_remitfull'+str(i+1)
            button=self.ui_name_dict[button_name]
            self.button_to_function_3[button]=partial(self.f_invoice_remit,part='Full',id=str(i+1))

            button_name='button_finan_4_remit'+str(i+1)+'_1'
            button=self.ui_name_dict[button_name]
            self.button_to_function_3[button]=partial(self.f_invoice_remit,part='Part1',id=str(i+1))

            button_name='button_finan_4_remit'+str(i+1)+'_2'
            button = self.ui_name_dict[button_name]
            self.button_to_function_3[button]=partial(self.f_invoice_remit,part='Part2',id=str(i+1))

            button_name='button_finan_4_remit'+str(i+1)+'_3'
            button = self.ui_name_dict[button_name]
            self.button_to_function_3[button]=partial(self.f_invoice_remit,part='Part3',id=str(i+1))

            button_name="button_finan_4_email" + str(i + 1)
            button=self.ui_name_dict[button_name]
            self.button_to_function_3[button]=partial(self.f_finan_email, id=i + 1)

        self.button_to_function_4 = {}
        for service in service_list_all[:-1]:
            button_name="button_fee_3_savesql_"+ service
            button = self.ui_name_dict[button_name]
            self.button_to_function_4[button]=partial(self.f_save_scope2default,service_short=service)
            for i in range(3):
                button_name="button_fee_" + service + '_up_' + str(i + 1)
                button=self.ui_name_dict[button_name]
                self.button_to_function_4[button] = partial(self.f_fee_scope_table_up, service=service, id=str(i + 1))

                button_name="button_fee_" + service + '_down_' + str(i + 1)
                button=self.ui_name_dict[button_name]
                self.button_to_function_4[button] = partial(self.f_fee_scope_table_down, service=service, id=str(i + 1))

                button_name = "button_fee_" + service + '_add_' + str(i + 1)
                button=self.ui_name_dict[button_name]
                self.button_to_function_4[button] = partial(self.f_fee_scope_table_add, service=service, id=str(i + 1))

        self.button_to_function_5 = {}
        for i in range(4):
            button_name="button_fee_stage_"+ str(i + 1)
            button=self.ui_name_dict[button_name]
            self.button_to_function_5[button]=partial(self.f_save_stage2default,stage=i)

        '''Button group 1'''
        self.radiobuttongroup_proinfo_1_proposaltype=QButtonGroup(self.ui)
        radiobuttongroup_proinfo_1_proposaltype_list=[self.ui.radbuttton_proinfo_1_proposaltype_1,self.ui.radbuttton_proinfo_1_proposaltype_2]
        self.radiobuttongroup_proinfo_1_projecttype=QButtonGroup(self.ui)
        radiobuttongroup_proinfo_1_projecttype_list=[self.ui_name_dict["radbuttton_proinfo_1_projecttype_" + str(i+1)] for i in range(8)]
        self.checkboxgroup_proinfo_1_servicettype=QButtonGroup(self.ui)
        self.checkboxgroup_proinfo_1_servicettype.setExclusive(False)
        checkboxgroup_proinfo_1_servicettype_list=[self.ui_name_dict["checkbox_proinfo_1_servicettype_" + str(i+1)] for i in range(9)]
        self.radbuttongroup_proinfo_2_clienttype=QButtonGroup(self.ui)
        radbuttongroup_proinfo_2_clienttype_list=[self.ui.radbutton_proinfo_2_clienttype_1,self.ui.radbutton_proinfo_2_clienttype_2]
        self.radbuttongroup_proinfo_2_clientcontacttype=QButtonGroup(self.ui)
        radbuttongroup_proinfo_2_clientcontacttype_list=[self.ui_name_dict["radbutton_proinfo_2_clientcontacttype_" + str(i + 1)] for i in range(12)]
        self.radbuttongroup_proinfo_2_contactcontacttype=QButtonGroup(self.ui)
        radbuttongroup_proinfo_2_contactcontacttype_list=[self.ui_name_dict["radbutton_proinfo_2_contactcontacttype_" + str(i + 1)] for i in range(12)]

        self.button_group_dict = {self.radiobuttongroup_proinfo_1_proposaltype: (radiobuttongroup_proinfo_1_proposaltype_list, self.f_proinfo_1_proposaltype_change),
                                  self.radiobuttongroup_proinfo_1_projecttype:(radiobuttongroup_proinfo_1_projecttype_list,''),
                                  self.checkboxgroup_proinfo_1_servicettype:(checkboxgroup_proinfo_1_servicettype_list,self.f_checkboxgroup_proinfo_1_servicettype),
                                  self.radbuttongroup_proinfo_2_clienttype:(radbuttongroup_proinfo_2_clienttype_list,''),
                                  self.radbuttongroup_proinfo_2_clientcontacttype:(radbuttongroup_proinfo_2_clientcontacttype_list,''),
                                  self.radbuttongroup_proinfo_2_contactcontacttype:(radbuttongroup_proinfo_2_contactcontacttype_list,''),}

        '''Button group 2'''
        self.radiobutton_groups={}
        self.radiobutton_groups_list=[]
        self.radiobutton_groups_dict={}
        for service in service_list_all:
            for i in range(4):
                button_group=QButtonGroup(self.ui)
                button_list=[self.ui_name_dict["radbutton_finan_1_" +service+'_'+str(i+1) + '_'+str(j + 1)] for j in range(8)]
                self.radiobutton_groups[button_group]=(button_list,self.f_radbutton_finan_change)
                self.radiobutton_groups_list.append(button_group)
                self.radiobutton_groups_dict[service+'_'+str(i+1)]=button_group

        '''Frame set'''
        for service in service_list_all[:-1]:
            frame_scope=self.ui_name_dict["frame_fee_scope_" + service]
            frame_scope.setVisible(False)
            frame_feedetail=self.ui_name_dict["frame_fee_pro_" + service]
            frame_feedetail.setVisible(False)
            frame_invoice=self.ui_name_dict["frame_finan_1_" + service]
            frame_invoice.setVisible(False)
            frame_bill=self.ui_name_dict["frame_finan_2_" + service]
            frame_bill.setVisible(False)
            frame_profit=self.ui_name_dict["frame_finan_3_" + service]
            frame_profit.setVisible(False)


        '''Text change connect'''
        self.text_to_function_1 = {self.ui.text_fee_5_minor_area:partial(self.f_fee_calculation_table_change, table_name="Minor"),
                                 self.ui.text_fee_5_major_apt: partial(self.f_fee_calculation_table_change,table_name="Major"),
                                 self.ui.text_fee_5_minor_price:partial(self.f_fee_calculation_table_change, table_name="Minor"),
                                 self.ui.text_fee_5_major_price: partial(self.f_fee_calculation_table_change,table_name="Major"),
                                 self.ui.text_proinfo_2_clientname:partial(self.f_client_search,client_type="Client"),
                                 self.ui.text_proinfo_2_contactname:partial(self.f_client_search,client_type="Main Contact"),
                                 self.ui.text_fee_1_revision:self.f_set_four_buttons,
                                 self.ui.text_proinfo_1_pronum:self.f_text_proinfo_1_pronum_change
                                 }

        self.text_to_function_2={}
        for item_name in item_name_list_all:
            text_name="text_finan_2_"+item_name+"_4"
            self.text_to_function_2[self.ui_name_dict[text_name]]=partial(self.f_text_finan_2_fee_change,item_name=item_name)
            text_name="text_finan_2_" + item_name + "_1"
            self.text_to_function_2[self.ui_name_dict[text_name]] = partial(self.f_text_finan_2_num_change, item_name=item_name)

        self.text_to_function_3 = {}
        for service in service_list_all:
            self.text_to_function_3[self.ui_name_dict["text_finan_1_" + service + "_fee0_1"]] = partial(self.f_text_finan_3_profit_calculation, item_name=service)
            self.text_to_function_3[self.ui_name_dict["text_finan_2_" + service + "_0_6"]] = partial(self.f_text_finan_2_price_sum_change, item_name=service)
            for i in range(4):
                self.text_to_function_3[self.ui_name_dict["text_finan_2_" + service + "_0_"+str(i+2)]] = partial(self.f_text_finan_2_price_change, item_name=service)
            for i in range(3):
                self.text_to_function_3[self.ui_name_dict["text_finan_2_" + service + '_' + str(i + 1) + '_2']]=partial(self.f_text_bill_client, service=service, text_id=i+1)

        self.text_to_function_4 = {}
        for i in range(8):
            self.text_to_function_4[self.ui_name_dict["text_finan_1_invclient_" + str(i+1)]]=partial(self.f_text_finan_client, text_id=i+1)
            self.text_to_function_4[self.ui_name_dict["text_finan_1_total" + str(i+1)+'_3']]=self.f_calculate_sum_payment
            self.text_to_function_4[self.ui_name_dict["text_finan_1_total" + str(i + 1)+'_1']]=partial(self.f_text_gen_quo_inv_num, text_id=i+1)
            self.text_to_function_4[self.ui_name_dict["text_finan_1_total" + str(i + 1) + '_4']] = self.f_set_payment_date
            self.text_to_function_4[self.ui_name_dict["text_finan_1_inv" + str(i+1)]]=partial(self.f_text_finan_inv_num, index=i+1)


        self.text_to_function_5={}
        ui_text_list = ['name', 'company', 'adress', 'abn', 'phone', 'email']
        type_list = ['client', 'contact']
        text_list = []
        for text_item_name in ui_text_list:
            for text_type in type_list:
                content_ui = self.ui_name_dict["text_proinfo_2_" + text_type + text_item_name]
                text_list.append(content_ui)
        text_list.append(self.ui.text_proinfo_1_proname)
        text_list.append(self.ui.text_proinfo_1_shopname)
        for text in text_list:
            self.text_to_function_5[text]=partial(self.f_text_proinfo_split, text_name=text)

        '''Table size set'''
        row_height_all = 20

        table_list=[self.ui.table_client_search,self.ui.table_contact_search,self.ui.table_finan_inv_client_search,self.ui.table_finan_bill_client_search]
        for table in table_list:
            font = QFont()
            font.setPointSize(conf['Table font'])
            table.setFont(font)
            table.setColumnHidden(0, True)
            for row in range(table.rowCount()):
                table.setRowHeight(row, row_height_all)
            header = table.horizontalHeader()
            for column in range(1, table.columnCount()):
                header.setSectionResizeMode(column, QHeaderView.Stretch)

        table_list=[self.ui.table_fee_5_cal_area,self.ui.table_fee_5_cal_apt,self.ui.table_fee_5_cal_carpark]
        for table in table_list:
            font = QFont()
            font.setPointSize(conf['Table font'])
            table.setFont(font)
            for row in range(table.rowCount()):
                table.setRowHeight(row, row_height_all)
                for column in range(table.columnCount()):
                    item = table.item(row, column) or QTableWidgetItem()
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
            header = table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Stretch)

        table_list = [self.ui.table_proinfo_3_1_area,self.ui.table_proinfo_4_drawing]
        for table in table_list:
            font = QFont()
            font.setPointSize(conf['Table font'])
            table.setFont(font)
            for row in range(table.rowCount()):
                table.setRowHeight(row, row_height_all)
                for column in range(table.columnCount()):
                    item = table.item(row, column) or QTableWidgetItem()
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
            header = table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Interactive)
            for row in [0,2]:
                table.setColumnWidth(row, 120)
            header.setSectionResizeMode(1, QHeaderView.Stretch)

        table_list=[self.ui.table_proinfo_3_2_carpark,self.ui.table_proinfo_3_2_area]
        for table in table_list:
            font = QFont()
            font.setPointSize(conf['Table font'])
            table.setFont(font)
            for row in range(table.rowCount()):
                table.setRowHeight(row, row_height_all)
                for column in range(table.columnCount()):
                    item = table.item(row, column) or QTableWidgetItem()
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
            header = table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Interactive)
            column_width_dict={0:120,2:45,3:45,4:45,5:45,6:45,7:45,8:45}
            for column,width in column_width_dict.items():
                table.setColumnWidth(column, width)
            header.setSectionResizeMode(1, QHeaderView.Stretch)

        for i in range(4):
            table=self.ui_name_dict["table_fee_stage_"+str(i+1)]
            font = QFont()
            font.setPointSize(conf['Table font'])
            table.setFont(font)
            for row in range(table.rowCount()):
                table.setRowHeight(row, row_height_all)
                for column in range(table.columnCount()):
                    item = table.item(row, column) or QTableWidgetItem()
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
            table.setColumnWidth(0, 40)
            table.horizontalHeader().setStretchLastSection(True)

        for service in service_list_all[:-1]:
            for j in range(3):
                table=self.ui_name_dict["table_fee_3_"+service+"_" + str(j+1)]
                font = QFont()
                font.setPointSize(conf['Table font'])
                table.setFont(font)
                for row in range(table.rowCount()):
                    table.setRowHeight(row, row_height_all)
                table.setColumnWidth(0, 40)
                table.horizontalHeader().setStretchLastSection(True)
                table.setSelectionBehavior(table.SelectRows)

        for service in service_list_all:
            table=self.ui_name_dict["table_fee_4_" + service]
            font = QFont()
            font.setPointSize(conf['Table font'])
            table.setFont(font)
            for row in range(table.rowCount()):
                table.setRowHeight(row, row_height_all)
                for column in range(table.columnCount()):
                    item = table.item(row, column) or QTableWidgetItem()
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

        table=self.ui.table_proinfo_history
        font = QFont()
        font.setPointSize(conf['Table font'])
        table.setFont(font)
        table.setColumnWidth(0, 100)
        table.setColumnWidth(1, 150)
        table.horizontalHeader().setStretchLastSection(True)

        '''Table hide'''
        self.ui.table_client_search.hide()
        self.ui.table_contact_search.hide()
        self.ui.table_finan_inv_client_search.hide()
        self.ui.table_finan_bill_client_search.hide()
        '''Table set'''
        self.table_connection_dict={}
        self.table_connection_dict[self.ui.table_proinfo_3_1_area]=self.f_table_proinfo_3_1_area_total
        self.table_proinfo_3_2_carpark_func=partial(self.f_table_proinfo_3_2_carpark_area_total, table_type="Car park")
        self.table_connection_dict[self.ui.table_proinfo_3_2_carpark] =self.table_proinfo_3_2_carpark_func
        self.table_proinfo_3_2_area_func=partial(self.f_table_proinfo_3_2_carpark_area_total, table_type="Apt")
        self.table_connection_dict[self.ui.table_proinfo_3_2_area] =self.table_proinfo_3_2_area_func
        self.table_fee_4_cal_dict={}
        for service in service_list_all:
            func=partial(self.f_fee_4_total, table_type=service)
            table=self.ui_name_dict["table_fee_4_" + service]
            self.table_fee_4_cal_dict[service]=(table,func)
            self.table_connection_dict[table] =func
        self.table_connection_dict[self.ui.table_fee_5_cal_carpark] =self.f_table_fee_5_cal_carpark_cal

        self.ui.table_client_search.itemClicked.connect(partial(self.f_search_table_click, contact_type='Client'))
        self.ui.table_contact_search.itemClicked.connect(partial(self.f_search_table_click, contact_type='Main Contact'))
        self.ui.table_finan_inv_client_search.itemClicked.connect(self.f_search_table_click2)
        self.ui.table_finan_bill_client_search.itemClicked.connect(self.f_search_table_click3)

        '''Label picture set'''
        pic_name = os.path.join(conf["database_dir_copilot_bridge"], 'carpark.png')
        pixmap = QPixmap(pic_name)
        self.ui.label_fee_5.setPixmap(pixmap)
        self.ui.label_fee_5.setScaledContents(True)
        '''Combobox set'''
        for i in range(self.ui.combobox_pro_1_state.count()):
            item = self.ui.combobox_pro_1_state.model().item(i)
            item.setBackground(QColor('gray') if i < 4 else QColor('white'))
            if i < 4:
                item.setFlags(item.flags() & ~Qt.ItemIsSelectable)
        '''Checkbox change set'''
        self.checkbox_to_function= {}
        for item_name in item_name_list_all:
            self.checkbox_to_function[self.ui_name_dict["checkbox_finan_2_" + item_name]] = partial(self.f_text_finan_2_fee_change, item_name=item_name)

        self.checkbox_to_function_1 = {}
        for service in service_list_all:
            self.checkbox_to_function_1[self.ui_name_dict["checkbox_finan_2_" + service+"_0"]] = partial(self.f_text_finan_2_price_sum_change, item_name=service)

        self.checkbox_to_function_2 = {}
        for i in range(4):
            self.checkbox_to_function_2[self.ui_name_dict["checkbox_fee_stage_" + str(i+1)]]=partial(self.f_stage_total_check,index=i+1)

        self.connect_all()

        '''Stage checkbox set'''
        for i in range(4):
            table = self.ui_name_dict["table_fee_stage_" + str(i + 1)]
            for row in range(table.rowCount()):
                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(checkbox_item.flags() | Qt.ItemIsUserCheckable)
                checkbox_item.setCheckState(Qt.Unchecked)
                table.setItem(row, 0, checkbox_item)

        '''State change set'''
        self.ui.combobox_pro_1_state.currentIndexChanged.connect(self.f_pro_state_change)

        '''Thread definition'''
        self.thread_1=Wait_email_fee(self)
        self.thread_1.sinout1.connect(self.f_email_fee_get)
        self.thread_2=Wait_email_inv(self)
        self.thread_2.sinout2.connect(self.f_email_inv_get)
        self.thread_3=Software_logout(self)
        self.thread_3.sinout3.connect(self.f_software_logout)
        self.thread_4=Refresh_table(self)
        self.thread_4.sinout4.connect(self.f_refresh_finish)


        self.datanow ={"Client id":None,"Main contact id":None,"Xero client id":None,"Xero main contact id":None,"Client invoice id":[None for _ in range(8)]}
    '''=============================Connection set===================================='''
    '''Button connection'''
    def connect_buttons(self):
        try:
            button_to_function_list=[self.button_to_function_1,self.button_to_function_2,self.button_to_function_3,
                                     self.button_to_function_4,self.button_to_function_5]
            for button_to_function in button_to_function_list:
                for button, function in button_to_function.items():
                    button.clicked.connect(function)
        except:
            traceback.print_exc()
    '''Checkbox connection'''
    def checkbox_connection(self):
        try:
            for checkbox, function in self.checkbox_to_function.items():
                checkbox.stateChanged.connect(function)
            for checkbox, function in self.checkbox_to_function_1.items():
                checkbox.stateChanged.connect(function)
            for checkbox, function in self.checkbox_to_function_2.items():
                checkbox.stateChanged.connect(function)
        except:
            traceback.print_exc()
    '''Text connection'''
    def connect_texts(self):
        try:
            text_to_function_list=[self.text_to_function_1,self.text_to_function_2,self.text_to_function_3,
                                   self.text_to_function_4,self.text_to_function_5]
            for text_to_function in text_to_function_list:
                for text, function in text_to_function.items():
                    text.textChanged.connect(function)
        except:
            traceback.print_exc()

    '''Button group connection'''
    def connect_button_groups(self):
        try:
            for button_group in self.button_group_dict.keys():
                button_group_list=self.button_group_dict[button_group][0]
                function=self.button_group_dict[button_group][1]
                for i in range(len(button_group_list)):
                    button_group.addButton(button_group_list[i], i+1)
                if function!='':
                    button_group.buttonClicked.connect(function)
            for button_group in self.radiobutton_groups.keys():
                button_group_list=self.radiobutton_groups[button_group][0]
                function =self.radiobutton_groups[button_group][1]
                for i in range(len(button_group_list)):
                    button_group.addButton(button_group_list[i], i+1)
                if function!='':
                    button_group.buttonClicked.connect(function)
        except:
            traceback.print_exc()

    '''Table connection'''
    def table_connection(self):
        try:
            for table, func in self.table_connection_dict.items():
                table.itemChanged.connect(func)
        except:
            traceback.print_exc()

    '''Button disconnection'''
    def disconnect_buttons(self):
        try:
            button_to_function_list=[self.button_to_function_1,self.button_to_function_2,self.button_to_function_3,
                                     self.button_to_function_4,self.button_to_function_5]
            for button_to_function in button_to_function_list:
                for button, function in button_to_function.items():
                    button.clicked.disconnect(function)
        except:
            traceback.print_exc()
    '''Checkbox disconnection'''
    def checkbox_disconnection(self):
        try:
            for checkbox, function in self.checkbox_to_function.items():
                checkbox.stateChanged.disconnect(function)
            for checkbox, function in self.checkbox_to_function_1.items():
                checkbox.stateChanged.disconnect(function)
            for checkbox, function in self.checkbox_to_function_2.items():
                checkbox.stateChanged.disconnect(function)
        except:
            traceback.print_exc()
    '''Text disconnection'''
    def disconnect_texts(self):
        try:
            text_to_function_list=[self.text_to_function_1,self.text_to_function_2,self.text_to_function_3,
                                   self.text_to_function_4,self.text_to_function_5]
            for text_to_function in text_to_function_list:
                for text, function in text_to_function.items():
                    text.textChanged.disconnect(function)
        except:
            traceback.print_exc()

    '''Button group disconnection'''
    def disconnect_button_groups(self):
        try:
            for button_group in self.button_group_dict.keys():
                button_group_list=self.button_group_dict[button_group][0]
                function=self.button_group_dict[button_group][1]
                for i in range(len(button_group_list)):
                    button_group.addButton(button_group_list[i], i+1)
                if function!='':
                    button_group.buttonClicked.disconnect(function)
            for button_group in self.radiobutton_groups.keys():
                button_group_list=self.radiobutton_groups[button_group][0]
                function =self.radiobutton_groups[button_group][1]
                for i in range(len(button_group_list)):
                    button_group.addButton(button_group_list[i], i+1)
                if function!='':
                    button_group.buttonClicked.disconnect(function)
        except:
            traceback.print_exc()

    def table_disconnection(self):
        try:
            for table, func in self.table_connection_dict.items():
                table.itemChanged.connect(func)
        except:
            traceback.print_exc()


    '''Connect all'''
    def connect_all(self):
        try:
            self.connect_buttons()
        except:
            traceback.print_exc()
        try:
            self.checkbox_connection()
        except:
            traceback.print_exc()
        try:
            self.connect_texts()
        except:
            traceback.print_exc()
        try:
            self.connect_button_groups()
        except:
            traceback.print_exc()
        try:
            self.table_connection()
        except:
            traceback.print_exc()

    '''Disconnect all'''
    def disconnect_all(self):
        try:
            self.disconnect_buttons()
        except:
            traceback.print_exc()
        try:
            self.checkbox_disconnection()
        except:
            traceback.print_exc()
        try:
            self.disconnect_texts()
        except:
            traceback.print_exc()
        try:
            self.disconnect_button_groups()
        except:
            traceback.print_exc()
        try:
            self.table_disconnection()
        except:
            traceback.print_exc()

    '''Calculate all'''
    def calculate_all(self):
        item_name_list=["mech_1","mech_2","mech_3","cfd_1","cfd_2","cfd_3","ele_1","ele_2","ele_3","hyd_1","hyd_2","hyd_3",
                        "fire_1","fire_2","fire_3","install_1","install_2","install_3","mis_1","mis_2","mis_3",
                        "mechrev_1","mechrev_2","mechrev_3","var_1","var_2","var_3"]
        service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis', 'var']
        try:
            '''Checkbox'''
            for item in item_name_list:
                self.f_text_finan_2_fee_change(item)
            for service in service_list:
                self.f_text_finan_2_price_sum_change(service)
            '''Text'''
            for service in service_list:
                self.f_text_finan_2_price_change(service)
                self.f_text_finan_3_profit_calculation(service)
            self.f_calculate_sum_payment()
            self.f_fee_calculation_table_change("Minor")
            self.f_fee_calculation_table_change("Major")
            '''Button group'''
            self.f_radbutton_finan_change()
            '''Table'''
            self.f_table_proinfo_3_2_carpark_area_total("Car park")
            self.f_table_proinfo_3_2_carpark_area_total("Apt")
            for service in service_list:
                self.f_fee_4_total(service)
            self.f_table_fee_5_cal_carpark_cal()
        except:
            traceback.print_exc()
    '''Frame show or not connection'''
    def toggleVisibility_all(self,frame_name,toolbutton_name,true_text,false_text):
        try:
            frame_name.setVisible(not frame_name.isVisible())
            if frame_name.isVisible()==True:
                toolbutton_name.setText(true_text)
            else:
                toolbutton_name.setText(false_text)
        except:
            pass

    '''=============================General function================================='''
    '''Set color of table'''
    def f_set_table_style(self,table):
        try:
            for row in range(table.rowCount()):
                for column in range(table.columnCount()):
                    if table.item(row, column)==None:
                        item=QTableWidgetItem()
                    else:
                        item = table.item(row, column)
                    item.setBackground(QColor(240, 240, 240) if row % 2 == 0 else QColor(255, 255, 255))
                    table.setItem(row, column, item)
        except:
            traceback.print_exc()
    '''Get table item by row and column'''
    def f_get_table_item(self,table,row,column):
        try:
            item=table.item(row,column)
            if item is None:
                return ''
            else:
                return item.text()

        except:
            traceback.print_exc()
    '''Change text to None'''
    def f_change2none(self,content):
        try:
            if content=='':
                return None
            else:
                return content
        except:
            traceback.print_exc()
    '''Change table content to None'''
    def f_tableitem2none(self,item):
        try:
            if item:
                if str(item.text())!='':
                    return item.text()
            return None
        except:
            traceback.print_exc()
            return None
    '''Str list contains non-unicode letter'''
    def contains_non_unicode_characters(self,str_list):
        contains=False
        for s in str_list:
            if s!=None:
                if bool(re.search(r'[^\x00-\x7F]', s)):
                    contains=True
        return contains
    '''Clear text list'''
    def clear_text_list(self,text_list):
        try:
            for text in text_list:
                text.setText('')
        except:
            traceback.print_exc()
    '''Clear button group list'''
    def clear_buttongroup(self,buttongroup_list):
        try:
            for buttongroup in buttongroup_list:
                for button in buttongroup.buttons():
                    button.setChecked(False)
        except:
            traceback.print_exc()
    '''Clear radiobutton group list'''
    def clear_radiobutton_group(self,radiobuttongroup_list):
        try:
            for buttongroup in radiobuttongroup_list:
                radio1 = QRadioButton("")
                buttongroup.addButton(radio1)
                radio1.setChecked(True)
                buttongroup.removeButton(radio1)
        except:
            traceback.print_exc()
    '''Add user history'''
    def f_add_user_history(self,action_content):
        try:
            table=self.ui.table_proinfo_history
            table_row_now = table.rowCount()
            table.insertRow(table_row_now)
            user_name=self.user_email
            action_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            table.setItem(table_row_now, 0, QTableWidgetItem(user_name))
            table.setItem(table_row_now, 1, QTableWidgetItem(str(action_time)))
            table.setItem(table_row_now, 2, QTableWidgetItem(str(action_content)))
            self.f_set_table_style(table)
        except:
            traceback.print_exc()
    '''Get proposal type'''
    def get_proposal_type(self):
        try:
            proposal_id = self.radiobuttongroup_proinfo_1_proposaltype.checkedId()
            if proposal_id == -1:
                message('Please choose proposal type',self.parent.ui)
                return
            proposal_type_list = ['Minor Project', 'Major Project']
            proposal_type = proposal_type_list[proposal_id - 1]
            return proposal_type
        except:
            traceback.print_exc()
    '''Get service type'''
    def f_get_pro_service_type(self,function_type):
        try:
            service_types_list = ['Mechanical Service', 'CFD Service', 'Electrical Service',
                                  'Hydraulic Service', 'Fire Service', 'Kitchen Ventilation',
                                  'Mechanical Review', 'Miscellaneous', 'Installation']
            service_type_list = []
            checkboxgroup_proinfo_1_servicettype_list = [self.ui_name_dict["checkbox_proinfo_1_servicettype_" + str(i + 1)] for i in range(9)]
            for i in range(len(checkboxgroup_proinfo_1_servicettype_list)):
                if checkboxgroup_proinfo_1_servicettype_list[i].isChecked():
                    service_type_list.append(service_types_list[i])
            if function_type=='Invoice':
                service_type_list.append('Variation')
                if 'Kitchen Ventilation' in service_type_list:
                    service_type_list.remove('Kitchen Ventilation')
            elif function_type=='Fee proposal':
                if 'Kitchen Ventilation' in service_type_list:
                    service_type_list.remove('Kitchen Ventilation')
                if 'Installation' in service_type_list:
                    service_type_list.remove('Installation')
                if 'Mechanical Service' in service_type_list and 'Mechanical Review' in service_type_list:
                    service_type_list.remove('Mechanical Review')
            elif function_type=='Frame_with_var':
                service_type_list.append('Variation')
                if 'Mechanical Service' in service_type_list and 'Kitchen Ventilation' in service_type_list:
                    service_type_list.remove('Kitchen Ventilation')
                elif 'Kitchen Ventilation' in service_type_list:
                    service_type_list.remove('Kitchen Ventilation')
                    service_type_list.append('Mechanical Service')
            elif function_type=='Frame_without_var':
                if 'Mechanical Service' in service_type_list and 'Kitchen Ventilation' in service_type_list:
                    service_type_list.remove('Kitchen Ventilation')
                elif 'Kitchen Ventilation' in service_type_list:
                    service_type_list.remove('Kitchen Ventilation')
                    service_type_list.append('Mechanical Service')
            elif function_type=='Set up':
                pass
            return service_type_list
        except:
            traceback.print_exc()
    '''Get project type'''
    def f_get_project_type(self):
        try:
            pro_type_id = self.radiobuttongroup_proinfo_1_projecttype.checkedId()
            if pro_type_id == -1:
                message("Please choose project type",self.parent.ui)
                return
            pro_type_list = ['Restaurant', 'Office', 'Commercial', 'Group House', 'Apartment',
                             'Mixed-use Complex', 'School', 'Others']
            project_type = pro_type_list[pro_type_id - 1]
            return project_type
        except:
            traceback.print_exc()
    def f_get_project_type2(self):
        try:
            pro_type_id = self.radiobuttongroup_proinfo_1_projecttype.checkedId()
            if pro_type_id == -1:
                return None
            pro_type_list = ['Restaurant', 'Office', 'Commercial', 'Group House', 'Apartment',
                             'Mixed-use Complex', 'School', 'Others']
            project_type = pro_type_list[pro_type_id - 1]
            return project_type
        except:
            traceback.print_exc()
    '''Get contact type'''
    def get_contact_type(self):
        try:
            contact_id = self.radbuttongroup_proinfo_2_clienttype.checkedId()
            if contact_id == -1:
                message("Please choose contact type",self.parent.ui)
                return
            contact_type_list = ['Client', 'Main Contact']
            contact_type = contact_type_list[contact_id - 1]
            return contact_type
        except:
            traceback.print_exc()
    '''Get service name for pdf'''
    def f_get_service_name_for_pdf(self):
        try:
            service_types_list = ['Mechanical', 'CFD', 'Electrical', 'Hydraulic',
                                  'Fire', 'Mechanical', 'Mechanical Review', 'Miscellaneous', 'Installation']
            service_type_list = []
            checkboxgroup_proinfo_1_servicettype_list = [self.ui_name_dict["checkbox_proinfo_1_servicettype_" + str(i + 1)] for i in [0, 1, 2, 3, 4, 6, 7]]
            for i in range(len(checkboxgroup_proinfo_1_servicettype_list)):
                if checkboxgroup_proinfo_1_servicettype_list[i].isChecked():
                    service_type_list.append(service_types_list[i])
            if len(service_type_list) == 1:
                service_name = service_type_list[0]
            elif len(service_type_list) == 2:
                service_name = f'{service_type_list[0]} and {service_type_list[1]}'
            elif len(service_type_list) == 3:
                service_name = f'{service_type_list[0]}, {service_type_list[1]} and {service_type_list[2]}'
            else:
                service_name = "Multi"
            return service_name
        except:
            traceback.print_exc()
    '''Get overdue fee'''
    def f_get_overdue_fee(self):
        overdue_fee=0
        try:
            for i in range(8):
                inv_text=self.ui_name_dict["text_finan_1_inv" + str(i+1)]
                palette = inv_text.palette()
                background_color = palette.color(palette.Window).name()
                if background_color=="#ff0000":
                    fee=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_1'].text()
                    payment=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_3'].text()
                    if is_float(fee) and is_float(payment):
                        overdue_fee+=float(fee)-float(payment)
        except:
            traceback.print_exc()
        finally:
            return overdue_fee
    '''Get invoice dict'''
    def f_get_invoice_dict(self,inv_id):
        invoice={"Invoice status":None,"name":None,"Payment Date":None,"Payment InGST":None,"Net":None}
        try:
            color_dict = {"#d2d2d2": "Backlog", "#ff0000": "Sent", "#008000": "Paid", "#800080": "Voided", "#aaaaaa":"Backlog"}
            inv_text=self.ui_name_dict["text_finan_1_inv" + str(inv_id)]
            invoice_num=inv_text.text()
            palette = inv_text.palette()
            background_color = palette.color(palette.Window).name()
            invoice["Invoice status"]=color_dict[background_color]
            invoice["name"] = 'INV ' + invoice_num
            invoice["Net"] = self.f_change2none(self.ui_name_dict["text_finan_1_total" + str(inv_id)+'_1'].text())
            invoice["Payment InGST"]=self.f_change2none(self.ui_name_dict["text_finan_1_total" + str(inv_id)+'_3'].text())
            invoice["Payment Date"]=self.f_change2none(self.ui_name_dict["text_finan_1_total" + str(inv_id)+'_4'].text())
        except:
            traceback.print_exc()
        finally:
            return invoice
    '''Get invoice update list'''
    def f_get_inv_update_list(self):
        inv_update_list=[]
        try:
            for i in range(8):
                inv_text_name=self.ui_name_dict["text_finan_1_inv" + str(i + 1)]
                inv_fee_name=self.ui_name_dict["text_finan_1_total" + str(i + 1)+'_1']
                try:
                    if inv_text_name.text()!='' and float(inv_fee_name.text())>0:
                        inv_update_list.append(i)
                except:
                    pass
        except:
            traceback.print_exc()
        finally:
            return inv_update_list
    '''Get bill dict'''
    def f_get_bill_dict(self,service_short, bill_subid,bill_id):
        bill_dict={"name":None,"Bill status":None,"From":None,"Bill in date":None,"Paid date":None,"Type":None,"Amount Excl GST":None,"Amount Incl GST":None,"HeadsUp":None}
        color_dict = {"#aaaaaa": "Draft","#d2d2d2": "Draft", "#ff0000": "Awaiting Approval", "#008000": "Paid","#800080": "Voided", "#ffa500": "Awaiting Payment"}
        try:
            item_name=service_short + '_' + str(bill_subid)
            text_name_number =self.ui_name_dict["text_finan_2_" + item_name + '_1'].text()
            bill_dict["name"] = "BIL " + self.ui.text_proinfo_1_pronum.text() + text_name_number
            bill_dict["Bill in date"]=self.datanow["Bill in date"][bill_id]
            bill_dict["Paid date"]=self.datanow["Bill paid date"][bill_id]
            description =self.ui_name_dict["text_finan_2_" + item_name + '_2'].text()
            bill_dict["From"] = description
            doc_name = self.ui_name_dict["text_finan_2_" + item_name + '_3'].text()
            bill_dict["HeadsUp"] = doc_name
            fee = self.ui_name_dict["text_finan_2_" + item_name + '_4'].text()
            bill_dict["Amount Excl GST"] = fee
            fee_total_text = self.ui_name_dict["text_finan_2_" + item_name + '_5']
            bill_dict["Amount Incl GST"] = fee_total_text.text()
            palette = fee_total_text.palette()
            background_color = palette.color(palette.Window).name()
            bill_dict["Bill status"] = color_dict[background_color]
            bill_type = self.ui_name_dict["combobox_finan_2_" + item_name].currentText()
            bill_dict["Type"] = bill_type
        except:
            traceback.print_exc()
        finally:
            return bill_dict
    '''Get bill total dict'''
    def f_get_bill_total_dict(self):
        bill_total_dict={}
        try:
            service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis', 'var']
            for service_short in service_list:
                for i in range(3):
                    bill_num=self.ui_name_dict["text_finan_2_"+ service_short + '_' +str(i + 1)+'_1'].text()
                    if bill_num!='':
                        bill_total_dict[bill_num]=(service_short,i+1)
        except:
            traceback.print_exc()
        finally:
            return bill_total_dict
    '''File open or not'''
    def is_file_open(self,file_path):
        for process in psutil.process_iter(['pid', 'name']):
            try:
                for item in process.open_files():
                    if item.path == file_path:
                        return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False
    '''Folder files open or not'''
    def check_folder_files_open(self,folder_path):
        new_folder_path=folder_path+'_temp'
        try:
            os.rename(folder_path, new_folder_path)
            os.rename(new_folder_path, folder_path)
            return True
        except:
            return False
    '''=============================Thread sinout fuctions=========================='''
    '''Fee proposal email get'''
    def f_email_fee_get(self,quo_num):
        try:
            quo_num_now=self.ui.text_proinfo_1_quonum.text()
            if quo_num==quo_num_now:
                pro_num=self.ui.text_proinfo_1_pronum.text()
                pro_name=self.ui.text_proinfo_1_proname.text()
                pro_info=quo_num+'-'+pro_num+'-'+pro_name
                color_tuple = ("green", "green", "green","yellow")
                self.f_set_pro_buttons_color(color_tuple)
                self.ui.combobox_pro_1_state.setCurrentText("Chase Fee Acceptance")
                update_asana=messagebox('Update Asana',"Do you want to update Asana?", self.parent.ui)
                if update_asana:
                    update_asana_result=self.f_button_pro_func_updateasana(False)
                    if update_asana_result == 'Success':
                        message(f'Project {pro_info} \nAsana Updated Successfully', self.parent.ui)
                    elif update_asana_result == 'Fail':
                        message(f'Project {pro_info} \nUpdate Asana failed, please contact admin', self.parent.ui)
                    else:
                        message(update_asana_result, self.parent.ui)
                self.f_add_user_history('Sent an email to Client')
            else:
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    user_email = self.user_email
                    action_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    action_content='Sent an email to Client'
                    mycursor.execute("INSERT INTO bridge.user_history VALUES (%s, %s, %s, %s)",(quo_num, user_email, action_time, action_content))
                    mycursor.execute("UPDATE bridge.status_colour SET gen_color = %s, email_color = %s, chase_color = %s WHERE quotation_number=%s", (2,2,1,quo_num))
                    mycursor.execute("UPDATE bridge.projects SET project_states = %s WHERE quotation_number=%s", ('Chase Fee Acceptance',quo_num))
                    mydb.commit()
                except:
                    traceback.print_exc()
                finally:
                    mycursor.close()
                    mydb.close()
        except:
            traceback.print_exc()

    '''Software logout'''
    def f_software_logout(self,pro_info_list):
        try:
            self.f_clear2default('True_must')
            self.parent.ui.tabWidget.setCurrentIndex(9)
            quo_num=pro_info_list[0]
            pro_num=pro_info_list[1]
            pro_name=pro_info_list[2]
            message(f'Project {quo_num}-{pro_num}-{pro_name} is logged out',self.parent.ui)
        except:
            traceback.print_exc()
    '''Refresh finish'''
    def f_refresh_finish(self,info):
        try:
            print(info)
            self.search_bar_update=True
        except:
            traceback.print_exc()

    '''Invoice email get'''
    def f_email_inv_get(self,quo_num):
        quotation_number = self.ui.text_proinfo_1_quonum.text()
        if quo_num!=quotation_number:
            return
        invoice_state_color_dict = {"Backlog": conf['Text gray style'],"Sent":conf['Text red style'],"Paid":conf['Text green style'],"Voided":conf['Text purple style'],}
        invoice_state_color_dict2 = {"Backlog": conf['Text gray_white style'], "Sent": conf['Text red_white style'],"Paid": conf['Text green_white style'], "Voided": conf['Text purple_white style'], }
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",
                                           database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM bridge.invoices WHERE quotation_number = %s", (quotation_number,))
            invoices_database = mycursor.fetchall()
        except:
            traceback.print_exc()
            return
        finally:
            mycursor.close()
            mydb.close()
        try:
            if len(invoices_database) > 0:
                for i in range(len(invoices_database)):
                    invoice_number = invoices_database[i][0]
                    state = invoices_database[i][2]
                    for j in range(8):
                        inv_text =self.ui_name_dict["text_finan_1_inv" + str(j + 1)]
                        finan_text_list=[self.ui_name_dict["text_finan_1_total" + str(j + 1)+'_'+str(k+1)] for k in range(5)]
                        if inv_text.text() == invoice_number:
                            style = invoice_state_color_dict[state]
                            style2 = invoice_state_color_dict2[state]
                            inv_text.setStyleSheet(style)
                            for finan_text in finan_text_list:
                                finan_text.setStyleSheet(style2)
        except:
            traceback.print_exc()

    '''=============================Project buttons================================='''
    '''0 - Close'''
    def on_close(self):
        self.f_clear2default(True)
    '''0.0 - Save to database'''
    def f_save2database(self):
        try:
            quotation_number=self.ui.text_proinfo_1_quonum.text()
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            '''bridge.projects'''
            project_number=self.f_change2none(self.ui.text_proinfo_1_pronum.text())
            project_name=self.f_change2none(self.ui.text_proinfo_1_proname.text())
            if project_name==None:
                message('Project name error',self.parent.ui)
            project_states=self.ui.combobox_pro_1_state.currentText()
            address_to=self.get_contact_type()
            shop_name=self.f_change2none(self.ui.text_proinfo_1_shopname.text())
            if self.radiobuttongroup_proinfo_1_proposaltype.checkedId() == 1:
                proposal_type='Minor'
            elif self.radiobuttongroup_proinfo_1_proposaltype.checkedId() == 2:
                proposal_type = 'Major'
            else:
                proposal_type = None
            project_type=self.f_get_project_type2()
            service_include_list=[0 for _ in range(9)]
            checkboxgroup_proinfo_1_servicettype_list = [self.ui_name_dict["checkbox_proinfo_1_servicettype_" + str(i + 1)] for i in range(9)]
            for i in range(len(checkboxgroup_proinfo_1_servicettype_list)):
                checkbox=checkboxgroup_proinfo_1_servicettype_list[i]
                if checkbox.isChecked():
                    service_include_list[i]=1
            client_id=self.datanow["Client id"]
            main_contact_id=self.datanow["Main contact id"]
            project_notes=None
            if proposal_type=='Minor':
                asana_apt=self.f_change2none(self.ui.text_proinfo_3_1_area.text())
                asana_basement=self.f_change2none(self.ui.text_proinfo_3_1_carspot.text())
                asana_feature=self.f_change2none(self.ui.text_proinfo_3_1_note.text())
            elif proposal_type=='Major':
                asana_apt = self.f_change2none(self.ui.text_proinfo_3_2_area.text())
                asana_basement=self.f_change2none(self.ui.text_proinfo_3_2_carspot.text())
                asana_feature=self.f_change2none(self.ui.text_proinfo_3_2_note.text())
                project_notes=self.f_change2none(self.ui.text_proinfo_3_2_pronote.toPlainText())
            else:
                asana_apt=None
                asana_basement=None
                asana_feature=None
            asana_id=self.f_change2none(self.datanow["Asana id"])
            asana_url=self.f_change2none(self.datanow["Asana url"])
            current_folder_address=self.f_change2none(self.datanow["Current folder"])
            email=self.f_change2none(self.ui.text_proinfo_email.toPlainText())
            fee_accepted_date=self.f_change2none(self.datanow["Fee accepted date"])
            fee_amount=self.f_change2none(self.ui.text_finan_1_total0_1.text())
            paid_fee=self.f_change2none(self.ui.text_finan_1_total0_3.text())
            overdue_fee=self.f_get_overdue_fee()
            mycursor.execute("UPDATE bridge.projects SET project_number = %s, project_name = %s, "
                             "project_states = %s, address_to = %s, shop_name = %s, proposal_type = %s, project_type = %s,"
                             "mechanical_service = %s, CFD_service = %s, electrical_service = %s, hydraulic_service = %s,"
                             "fire_service = %s, kitchen_ventilation = %s, mechanical_review = %s, miscellaneous = %s, "
                             "installation = %s, client_id = %s, main_contact_id = %s, asana_apt = %s, asana_basement = %s, asana_feature = %s,"
                             "project_notes = %s, asana_id = %s, asana_url = %s, current_folder_address = %s, email = %s,"
                             "fee_accepted_date = %s, fee_amount=%s, overdue_fee=%s, paid_fee=%s"
                             " WHERE quotation_number=%s",
                             (project_number,project_name,
                              project_states,address_to,shop_name,proposal_type,project_type,
                              service_include_list[0],service_include_list[1],service_include_list[2],service_include_list[3],
                              service_include_list[4],service_include_list[5],service_include_list[6],service_include_list[7],
                              service_include_list[8],client_id,main_contact_id,asana_apt,asana_basement,asana_feature,
                              project_notes,asana_id,asana_url,current_folder_address,email,
                              fee_accepted_date, fee_amount, overdue_fee, paid_fee, quotation_number,))
            '''bridge.levels'''
            if proposal_type == 'Minor':
                table=self.ui.table_proinfo_3_1_area
                for row in range(table.rowCount()-1):
                    levels=self.f_tableitem2none(table.item(row, 0))
                    spaces=self.f_tableitem2none(table.item(row, 1))
                    area=self.f_tableitem2none(table.item(row, 2))
                    mycursor.execute("UPDATE bridge.levels SET levels = %s, spaces = %s, area = %s"
                                     " WHERE quotation_number=%s AND row_index=%s",
                                     (levels, spaces, area, quotation_number,row,))
            '''bridge.rooms'''
            if proposal_type == 'Major':
                table=self.ui.table_proinfo_3_2_carpark
                for row in range(table.rowCount()-1):
                    levels=self.f_tableitem2none(table.item(row, 0))
                    room_type=self.f_tableitem2none(table.item(row, 1))
                    A_car_spots=self.f_tableitem2none(table.item(row, 2))
                    A_other_spots=self.f_tableitem2none(table.item(row, 3))
                    B_car_spots=self.f_tableitem2none(table.item(row, 4))
                    B_other_spots=self.f_tableitem2none(table.item(row, 5))
                    C_car_spots=self.f_tableitem2none(table.item(row, 6))
                    C_other_spots=self.f_tableitem2none(table.item(row, 7))
                    mycursor.execute("UPDATE bridge.rooms SET levels = %s, room_type = %s, A_car_spots = %s,"
                                     "A_other_spots = %s, B_car_spots = %s, B_other_spots = %s, C_car_spots = %s, C_other_spots = %s"
                                     " WHERE quotation_number=%s AND row_index=%s",
                                     (levels, room_type, A_car_spots, A_other_spots, B_car_spots, B_other_spots,
                                      C_car_spots, C_other_spots, quotation_number,row,))
            '''bridge.apartments'''
            if proposal_type == 'Major':
                table=self.ui.table_proinfo_3_2_area
                for row in range(table.rowCount()-1):
                    levels=self.f_tableitem2none(table.item(row, 0))
                    room_type=self.f_tableitem2none(table.item(row, 1))
                    A_car_spots=self.f_tableitem2none(table.item(row, 2))
                    A_other_spots=self.f_tableitem2none(table.item(row, 3))
                    B_car_spots=self.f_tableitem2none(table.item(row, 4))
                    B_other_spots=self.f_tableitem2none(table.item(row, 5))
                    C_car_spots=self.f_tableitem2none(table.item(row, 6))
                    C_other_spots=self.f_tableitem2none(table.item(row, 7))
                    mycursor.execute("UPDATE bridge.apartments SET levels = %s, room_type = %s, A_car_spots = %s,"
                                     "A_other_spots = %s, B_car_spots = %s, B_other_spots = %s, C_car_spots = %s, C_other_spots = %s"
                                     " WHERE quotation_number=%s AND row_index=%s",
                                     (levels, room_type, A_car_spots, A_other_spots, B_car_spots, B_other_spots,
                                      C_car_spots, C_other_spots, quotation_number,row,))
            '''bridge.drawings'''
            table=self.ui.table_proinfo_4_drawing
            for row in range(table.rowCount()):
                drawing_number=self.f_tableitem2none(table.item(row, 0))
                drawing_name=self.f_tableitem2none(table.item(row, 1))
                drawing_revision=self.f_tableitem2none(table.item(row, 2))
                mycursor.execute("UPDATE bridge.drawings SET drawing_number = %s, drawing_name = %s, drawing_revision = %s"
                                 " WHERE quotation_number=%s AND row_index=%s",
                                 (drawing_number, drawing_name, drawing_revision, quotation_number, row,))
            '''bridge.scope'''
            service_list=self.f_get_pro_service_type('Frame_without_var')
            for service_full_name in service_list:
                service=conf['Service short dict'][service_full_name]
                mycursor.execute("DELETE FROM bridge.scope"
                    " WHERE quotation_number=%s AND minor_or_major=%s AND service=%s",
                    (quotation_number, proposal_type,service_full_name,))
                table_list=[self.ui_name_dict["table_fee_3_" + service+'_'+str(i+1)] for i in range(3)]
                type_list=['Extent','Clarifications','Deliverables']
                for i in range(len(table_list)):
                    table=table_list[i]
                    type=type_list[i]
                    for row in range(table.rowCount()):
                        checkbox = table.item(row, 0)
                        include=0
                        if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                            if checkbox.checkState() == Qt.Checked:
                                include=1
                        content = self.f_tableitem2none(table.item(row, 1))
                        mycursor.execute("INSERT INTO bridge.scope VALUES (%s, %s, %s, %s, %s, %s, %s)",
                                 (quotation_number, proposal_type, service_full_name, type, include, content, row))
            '''bridge.major_stage'''
            if proposal_type=='Major':
                mycursor.execute("DELETE FROM bridge.major_stage"
                                 " WHERE quotation_number=%s",
                                 (quotation_number,))
                table_list =[self.ui_name_dict["table_fee_stage_" + str(i + 1)] for i in range(4)]
                for i in range(len(table_list)):
                    table=table_list[i]
                    stage=i+1
                    for row in range(table.rowCount()):
                        checkbox = table.item(row, 0)
                        include = 0
                        if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                            if checkbox.checkState() == Qt.Checked:
                                include = 1
                        content = self.f_tableitem2none(table.item(row, 1))
                        mycursor.execute("INSERT INTO bridge.major_stage VALUES (%s, %s, %s, %s, %s)",
                                         (quotation_number, include, stage, content, row))
            '''bridge.fee'''
            installation_date=self.f_change2none(self.ui.text_fee_3_install_date.text())
            if installation_date!=None:
                installation_date=str(datetime.strptime(installation_date, "%d-%b-%Y").strftime("%Y-%m-%d"))
            installation_revision=self.f_change2none(self.ui.text_fee_3_install_rev.text())
            installation_program=self.f_change2none(self.ui.text_fee_3_install_program.toPlainText())

            date_obj =datetime.strptime(self.ui.text_fee_1_date.text(), "%d-%b-%Y")
            formatted_date=date_obj.strftime("%Y-%m-%d")

            fee_date=self.f_change2none(formatted_date)
            fee_revision=self.f_change2none(self.ui.text_fee_1_revision.text())
            fee_proposal_stage=self.f_change2none(self.ui.text_fee_2_period1.text())
            pre_design_stage=self.f_change2none(self.ui.text_fee_2_period2.text())
            documentation_stage=self.f_change2none(self.ui.text_fee_2_period3.text())
            stage_1_name=self.f_change2none(self.ui.text_stage_table_name1.text())
            if self.ui.checkbox_fee_stage_1.isChecked():
                stage_1_include=1
            else:
                stage_1_include=0
            stage_2_name=self.f_change2none(self.ui.text_stage_table_name2.text())
            if self.ui.checkbox_fee_stage_2.isChecked():
                stage_2_include = 1
            else:
                stage_2_include = 0
            stage_3_name=self.f_change2none(self.ui.text_stage_table_name3.text())
            if self.ui.checkbox_fee_stage_3.isChecked():
                stage_3_include = 1
            else:
                stage_3_include = 0
            stage_4_name=self.f_change2none(self.ui.text_stage_table_name4.text())
            if self.ui.checkbox_fee_stage_4.isChecked():
                stage_4_include = 1
            else:
                stage_4_include = 0
            area_price=self.f_change2none(self.ui.text_fee_5_minor_price.text())
            apt_price=self.f_change2none(self.ui.text_fee_5_major_price.text())
            total_area=self.f_change2none(self.ui.text_fee_5_minor_area.text())
            total_apt=self.f_change2none(self.ui.text_fee_5_major_apt.text())
            mycursor.execute("UPDATE bridge.fee SET installation_date = %s, installation_revision = %s, installation_program = %s,"
                             "fee_date = %s,fee_revision = %s, fee_proposal_stage = %s, pre_design_stage = %s, documentation_stage = %s, stage_1_name = %s,"
                             "stage_1_include = %s, stage_2_name = %s, stage_2_include = %s, stage_3_name = %s, stage_3_include = %s,"
                             "stage_4_name = %s, stage_4_include = %s, area_price = %s, apt_price = %s, total_area=%s,total_apt=%s"
                             " WHERE quotation_number=%s",
                             (installation_date, installation_revision, installation_program,
                              fee_date,fee_revision, fee_proposal_stage, pre_design_stage, documentation_stage, stage_1_name,
                              stage_1_include, stage_2_name, stage_2_include, stage_3_name, stage_3_include,
                              stage_4_name, stage_4_include, area_price, apt_price,total_area,total_apt,quotation_number,))
            '''bridge.fee_item'''
            service_list=self.f_get_pro_service_type('Frame_with_var')
            for service in service_list:
                service_short=conf['Service short dict'][service]
                table=self.ui_name_dict["table_fee_4_" + service_short]
                for row in range(table.rowCount()-1):
                    content = self.f_tableitem2none(table.item(row, 0))
                    amount = self.f_tableitem2none(table.item(row, 1))
                    column_index=None
                    for i in range(8):
                        radiobutton=self.ui_name_dict["radbutton_finan_1_" + service_short+'_'+str(row+1)+'_'+str(i+1)]
                        if radiobutton.isChecked():
                            column_index=i
                    mycursor.execute("UPDATE bridge.fee_item SET content = %s, amount = %s, column_index = %s"
                        " WHERE quotation_number=%s AND service=%s AND row_index=%s",
                        (content,amount,column_index,quotation_number,service,row))
            '''bridge.remittances'''
            for i in range(8):
                remit_type = None
                button_full=self.ui_name_dict["button_finan_4_remitfull" + str(i + 1)]
                button_part_list=[self.ui_name_dict["button_finan_4_remit" + str(i + 1)+'_'+str(j+1)] for j in range(3)]
                palette = button_full.palette()
                background_color = palette.color(palette.Window).name()
                if background_color=='#c8c8c8':
                    remit_type='Full'
                for button in button_part_list:
                    palette = button.palette()
                    background_color = palette.color(palette.Window).name()
                    if background_color == '#c8c8c8':
                        remit_type = 'Part'
                remit1=self.f_change2none(self.ui_name_dict["text_finan_4_remit_" + str(i + 1)+'_1'].text())
                remit2 = self.f_change2none(self.ui_name_dict["text_finan_4_remit_" + str(i + 1)+'_2'].text())
                remit3 = self.f_change2none(self.ui_name_dict["text_finan_4_remit_" + str(i + 1)+'_3'].text())
                mycursor.execute("UPDATE bridge.remittances SET remit_type = %s, remit1 = %s, remit2 = %s, remit3 = %s"
                                 " WHERE quotation_number=%s AND i=%s",
                                 (remit_type,remit1,remit2,remit3,quotation_number,i))


                '''bridge.remittance_upload'''
                button_list=[self.ui_name_dict["button_finan_4_preview" + str(i+1)],self.ui_name_dict["button_finan_4_email" + str(i+1)],button_full]+button_part_list
                button_onoff_list=[]
                for button in button_list:
                    palette = button.palette()
                    background_color = palette.color(palette.Window).name()
                    if background_color == '#c8c8c8':
                        button_onoff_list.append(1)
                    else:
                        button_onoff_list.append(0)
                mycursor.execute("UPDATE bridge.remittance_upload SET preview=%s, email=%s, remit_full = %s, remit_1 = %s, remit_2 = %s, remit_3 = %s"
                        " WHERE quotation_number=%s AND i=%s",
                        (button_onoff_list[0], button_onoff_list[1], button_onoff_list[2],button_onoff_list[3], button_onoff_list[4],button_onoff_list[5], quotation_number, i))

            '''bridge.user_history'''
            table=self.ui.table_proinfo_history
            mycursor.execute("DELETE FROM bridge.user_history"
                " WHERE quotation_number=%s",(quotation_number,))
            for row in range(table.rowCount()):
                user_email=self.f_tableitem2none(table.item(row, 0))
                action_time=self.f_tableitem2none(table.item(row, 1))
                action_content=self.f_tableitem2none(table.item(row, 2))
                try:
                    mycursor.execute("INSERT INTO bridge.user_history VALUES (%s, %s, %s, %s)",
                                     (quotation_number,user_email, action_time, action_content))
                except:
                    pass
            '''bridge.fee_acceptance'''
            version1_note=self.f_change2none(self.ui.text_finan_5_note1.text())
            version2_note = self.f_change2none(self.ui.text_finan_5_note2.text())
            version3_note = self.f_change2none(self.ui.text_finan_5_note3.text())
            mycursor.execute("UPDATE bridge.fee_acceptance SET version1_note = %s, version2_note = %s, version3_note = %s"
                             " WHERE quotation_number=%s",
                             (version1_note, version2_note, version3_note, quotation_number))
            '''bridge.acceptance_upload'''
            button_list=[self.ui.button_finan_5_upload,self.ui.button_finan_5_confirm]
            button_onoff_list=[]
            for button in button_list:
                palette = button.palette()
                background_color = palette.color(palette.Window).name()
                if background_color == '#c8c8c8':
                    button_onoff_list.append(1)
                else:
                    button_onoff_list.append(0)

            mycursor.execute("UPDATE bridge.acceptance_upload SET v1_upload=%s, v1_confirm=%s"
                " WHERE quotation_number=%s",(button_onoff_list[0], button_onoff_list[1], quotation_number))

            '''bridge.supplier_fee'''
            service_list = self.f_get_pro_service_type('Frame_with_var')
            for service_full in service_list:
                service_short = conf['Service short dict'][service_full]
                all_clients=self.f_change2none(self.ui_name_dict["text_finan_2_" + service_short+'_0_1'].text())
                checkbox=self.ui_name_dict["checkbox_finan_2_" + service_short+'_0']
                no_gst = 1 if checkbox.isChecked() else 0
                origin=self.f_change2none(self.ui_name_dict["text_finan_2_" + service_short+'_0_2'].text())
                version1=self.f_change2none(self.ui_name_dict["text_finan_2_" + service_short+'_0_3'].text())
                version2 = self.f_change2none(self.ui_name_dict["text_finan_2_" + service_short+'_0_4'].text())
                version3 = self.f_change2none(self.ui_name_dict["text_finan_2_" + service_short+'_0_5'].text())
                mycursor.execute("UPDATE bridge.supplier_fee SET all_clients = %s, no_gst = %s, origin = %s, version1 = %s, version2 = %s, version3 = %s"
                    " WHERE quotation_number=%s AND service=%s",
                    (all_clients,no_gst,origin,version1,version2,version3, quotation_number, service_full))
            '''bridge.invoices'''
            mycursor.execute("DELETE FROM bridge.invoices"
                " WHERE quotation_number=%s",(quotation_number,))
            color_dict = {"#d2d2d2": "Backlog", "#ff0000": "Sent", "#008000": "Paid", "#800080": "Voided", "#aaaaaa": "Backlog"}
            for i in range(8):
                invoice_number_text=self.ui_name_dict["text_finan_1_inv" + str(i+1)]
                invoice_number=self.f_change2none(invoice_number_text.text())
                if invoice_number!=None:
                    palette = invoice_number_text.palette()
                    background_color = palette.color(palette.Window).name()
                    state = color_dict[background_color]
                    asana_id=self.datanow["Asana invoice id"][i]
                    xero_id=self.datanow["Xero invoice id"][i]
                    client_id=self.datanow["Client invoice id"][i]
                    payment_date_text=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_4']
                    payment_amount_text=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_3']
                    fee_amount_text=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_1']
                    payment_date=self.f_change2none(payment_date_text.text())
                    payment_amount=self.f_change2none(payment_amount_text.text())
                    fee_amount=self.f_change2none(fee_amount_text.text())
                    mycursor.execute("INSERT INTO bridge.invoices VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                                     (invoice_number,quotation_number,state,asana_id,xero_id,client_id,payment_date,payment_amount,fee_amount,i))
            '''bridge.bills'''
            mycursor.execute("DELETE FROM bridge.bills"
                             " WHERE quotation_number=%s", (quotation_number,))
            color_dict = {"#aaaaaa": "Draft", "#ff0000": "Awaiting Approval", "#008000": "Paid",
                          "#800080": "Voided", "#ffa500": "Awaiting Payment", "#d2d2d2": "Draft"}
            service_list=self.f_get_pro_service_type('Frame_with_var')
            for service in service_list:
                service_short=conf['Service short dict'][service]
                for i in range(3):
                    bill_letter=self.ui_name_dict["text_finan_2_" +service_short+ '_' +str(i+1)+'_1'].text()
                    if bill_letter!='':
                        bill_number=self.ui.text_proinfo_1_pronum.text()+bill_letter
                        fee_total_text =self.ui_name_dict["text_finan_2_" + service_short + '_' + str(i + 1) + '_5']
                        palette = fee_total_text.palette()
                        background_color = palette.color(palette.Window).name()
                        state = color_dict[background_color]
                        bill_type =self.ui_name_dict["combobox_finan_2_" + service_short + '_' + str(i + 1)].currentText()
                        fee_amount_text=self.ui_name_dict["text_finan_2_" + service_short + '_' + str(i + 1) + '_4']
                        fee_amount=self.f_change2none(fee_amount_text.text())
                        checkbox=self.ui_name_dict["checkbox_finan_2_" + service_short + '_' + str(i + 1)]
                        if checkbox.isChecked():
                            no_gst=1
                        else:
                            no_gst=0
                        heads_up_text =self.ui_name_dict["text_finan_2_" + service_short + '_' + str(i + 1) + '_3']
                        heads_up = self.f_change2none(heads_up_text.text())
                        bill_id=char2num(bill_letter)
                        asana_id=self.datanow["Asana bill id"][bill_id]
                        xero_id=self.datanow["Xero bill"][bill_id]
                        client_id=self.datanow["Client bill id"][bill_id]
                        bill_in_date=self.datanow["Bill in date"][bill_id]
                        paid_date=self.datanow["Bill paid date"][bill_id]
                        mycursor.execute("INSERT INTO bridge.bills VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                                         (bill_number,quotation_number,state,asana_id,xero_id,client_id,bill_in_date,paid_date,bill_type,service,fee_amount,no_gst,heads_up,i))
                        mydb.commit()
            '''bridge.bill_upload'''
            for service_full, service_short in conf['Service short dict'].items():
                for i in range(3):
                    button=self.ui_name_dict["button_finan_2_" +service_short + '_upload'+ str(i + 5)]
                    palette = button.palette()
                    background_color = palette.color(palette.Window).name()
                    if background_color == '#c8c8c8':
                        button_onoff=1
                    else:
                        button_onoff=0
                    mycursor.execute("UPDATE bridge.bill_upload SET upload=%s"
                        " WHERE quotation_number=%s AND service=%s AND i=%s",
                        (button_onoff,quotation_number,service_full,i))
            '''bridge.supplier_upload'''
            for service_full, service_short in conf['Service short dict'].items():
                button_upload_list=[]
                for i in range(4):
                    button=self.ui_name_dict["button_finan_2_" +service_short + '_upload'+ str(i + 1)]
                    palette = button.palette()
                    background_color = palette.color(palette.Window).name()
                    if background_color == '#c8c8c8':
                        button_upload_list.append(1)
                    else:
                        button_upload_list.append(0)
                mycursor.execute("UPDATE bridge.supplier_upload SET original_upload=%s,v1_upload=%s,v2_upload=%s,v3_upload=%s"
                                 " WHERE quotation_number=%s AND service=%s",
                                 (button_upload_list[0], button_upload_list[1], button_upload_list[2], button_upload_list[3],quotation_number, service_full))
            '''bridge.status_colour'''
            button_list = [self.ui.button_pro_state_genfee,self.ui.button_pro_state_emailfee, self.ui.button_pro_state_chasefee]
            color_dict = {"#a52a2a": 0, "#ffa500": 1,"#008000": 2}
            color_list2save=[]
            for button in button_list:
                palette = button.palette()
                background_color = palette.color(palette.Window).name()
                color_list2save.append(color_dict[background_color])
            mycursor.execute("UPDATE bridge.status_colour SET gen_color=%s,email_color=%s,chase_color=%s"
                " WHERE quotation_number=%s",
                (color_list2save[0], color_list2save[1], color_list2save[2], quotation_number))

            mydb.commit()
            return True
        except:
            traceback.print_exc()
            return False
        finally:
            mycursor.close()
            mydb.close()
    '''0.1 - Clear all'''
    def f_clear2default(self,save_or_not):
        try:
            self.thread_3.stop_thread_func()
        except:
            traceback.print_exc()
        try:
            self.ui.combobox_pro_1_state.currentIndexChanged.disconnect(self.f_pro_state_change)
        except:
            pass
        try:
            if self.ui.text_proinfo_1_quonum.text()!='' and save_or_not==True:
                save = messagebox('Save', 'Do you want to save project information to database?', self.parent.ui)
                if save:
                    save_right=self.f_save2database()
                    if not save_right:
                        return
            if self.ui.text_proinfo_1_quonum.text()!='' and save_or_not=='True_must':
                save_right=self.f_save2database()
                if not save_right:
                    return
            self.datanow ={"Client id":None,"Main contact id":None,"Xero client id":None,"Xero main contact id":None,}
            quotation2logout=self.ui.text_proinfo_1_quonum.text()
            '''Project table'''
            text_list=[self.ui.text_proinfo_1_quonum,self.ui.text_proinfo_1_pronum,self.ui.text_proinfo_1_proname,
                       self.ui.text_proinfo_1_shopname]
            self.clear_text_list(text_list)
            buttongroup_list=[self.checkboxgroup_proinfo_1_servicettype]
            self.clear_buttongroup(buttongroup_list)
            radiobuttongroup_list=[self.radiobuttongroup_proinfo_1_proposaltype,self.radiobuttongroup_proinfo_1_projecttype,self.radbuttongroup_proinfo_2_clienttype]
            self.clear_radiobutton_group(radiobuttongroup_list)
            self.ui.combobox_pro_1_state.setCurrentText('Set Up')
            '''Client and main contact table'''
            text_list=[]
            ui_text_list = ['name', 'company', 'adress', 'abn', 'phone', 'email']
            type_list=['client','contact']
            for text_item_name in ui_text_list:
                for text_type in type_list:
                    content_ui = self.ui_name_dict["text_proinfo_2_"+ text_type + text_item_name]
                    text_list.append(content_ui)
            self.clear_text_list(text_list)
            radiobuttongroup_list=[self.radbuttongroup_proinfo_2_clientcontacttype,self.radbuttongroup_proinfo_2_contactcontacttype]
            self.clear_radiobutton_group(radiobuttongroup_list)
            '''Building features - Minor'''
            self.ui.table_proinfo_3_1_area.clearContents()
            self.ui.table_proinfo_3_1_area.setItem(0, 0, QTableWidgetItem('Tenancy Level'))
            text_list = [self.ui.text_proinfo_3_1_area,self.ui.text_proinfo_3_1_carspot,self.ui.text_proinfo_3_1_note]
            self.clear_text_list(text_list)
            '''Building features - Major'''
            self.ui.table_proinfo_3_2_carpark.clearContents()
            self.ui.table_proinfo_3_2_area.clearContents()
            text_list = [self.ui.text_proinfo_3_2_area,self.ui.text_proinfo_3_2_carspot,self.ui.text_proinfo_3_2_note,self.ui.text_proinfo_3_2_pronote]
            self.clear_text_list(text_list)
            '''Drawing number'''
            self.ui.table_proinfo_4_drawing.clearContents()
            '''Email'''
            self.ui.text_proinfo_email.setText('')
            '''User history'''
            self.ui.table_proinfo_history.setRowCount(0)
            '''Reference'''
            text_list = [self.ui.text_fee_1_date,self.ui.text_fee_1_revision]
            self.clear_text_list(text_list)
            '''Stage'''
            stage_name_list=['Design Application','Design Development','Construction Documentation','Construction Phase Service']
            for i in range(4):
                text_name=self.ui_name_dict["text_stage_table_name" + str(i + 1)]
                text_name.setText(stage_name_list[i])
                checkbox=self.ui_name_dict["checkbox_fee_stage_" + str(i + 1)]
                checkbox.setChecked(False)
                table = self.ui_name_dict["table_fee_stage_" + str(i + 1)]
                table.clearContents()
                for row in range(table.rowCount()):
                    checkbox_item = QTableWidgetItem()
                    checkbox_item.setFlags(checkbox_item.flags() | Qt.ItemIsUserCheckable)
                    checkbox_item.setCheckState(Qt.Unchecked)
                    table.setItem(row, 0, checkbox_item)
            '''Scope of work'''
            service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis']
            for service in service_list:
                for i in range(3):
                    table=self.ui_name_dict["table_fee_3_" + service+'_'+str(i + 1)]
                    table.setRowCount(0)
            text_list = [self.ui.text_fee_3_install_date,self.ui.text_fee_3_install_rev,self.ui.text_fee_3_install_program]
            self.clear_text_list(text_list)
            '''Fee table set'''
            service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis','var']
            for service in service_list:
                table=self.ui_name_dict["table_fee_4_" + service]
                table.clearContents()
            text_list = [self.ui.text_fee_4_ex,self.ui.text_fee_4_in]
            self.clear_text_list(text_list)
            '''Calculation'''
            text_list = [self.ui.text_fee_5_minor_area,self.ui.text_fee_5_minor_price,self.ui.text_fee_5_major_apt,self.ui.text_fee_5_major_price]
            self.clear_text_list(text_list)
            self.ui.table_fee_5_cal_carpark.clearContents()
            '''Invoice set'''
            for i in range(8):
                text=self.ui_name_dict["text_finan_1_inv" + str(i + 1)]
                finan_text_list = [self.ui_name_dict["text_finan_1_total" + str(i + 1) + '_' + str(k + 1)] for k in range(5)]
                text.setText('')
                text.setStyleSheet(conf['Text original style'])
                for finan_text in finan_text_list:
                    finan_text.setText('')
                    finan_text.setStyleSheet(conf['Text original_white style'])
                text_client=self.ui_name_dict["text_finan_1_invclient_" + str(i + 1)]
                text_client.setText('')
            service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'install', 'mis', 'mechrev', 'var']
            for service in service_list:
                for i in range(4):
                    text_list = [self.ui_name_dict["text_finan_1_" + service + '_doc' + str(i + 1)],
                                 self.ui_name_dict["text_finan_1_" + service + '_fee' + str(i + 1) + '_1'],
                                 self.ui_name_dict["text_finan_1_" + service + '_fee' + str(i + 1) + '_2']]
                    self.clear_text_list(text_list)

                    if i==0 or i==2:
                        self.ui_name_dict["text_finan_1_" + service + '_doc' + str(i + 1)].setStyleSheet("""QLineEdit {font: 12pt "Calibri";color: rgb(0, 0, 0);border-style: ridge;
                                            border-color: rgb(100, 100, 100);border-width: 1px;background-color: rgb(240, 240, 240);}""")
                    else:
                        self.ui_name_dict["text_finan_1_" + service + '_doc' + str(i + 1)].setStyleSheet("""QLineEdit {font: 12pt "Calibri";color: rgb(0, 0, 0);border-style: ridge;
                                                                    border-color: rgb(100, 100, 100);border-width: 1px;background-color: rgb(220, 220, 220);}""")


                text_list=[self.ui_name_dict["text_finan_1_" + service + "_fee0_1"],self.ui_name_dict["text_finan_1_" + service + "_fee0_2"]]
                self.clear_text_list(text_list)
            self.clear_radiobutton_group(self.radiobutton_groups_list)

            text_list=[self.ui_name_dict["text_finan_1_total" +str(i + 1)+'_4'] for i in range(8)]
            self.clear_text_list(text_list)
            text_list = [self.ui.text_finan_1_total0_1,self.ui.text_finan_1_total0_2]
            for i in range(8):
                for j in range(3):
                    text_list.append(self.ui_name_dict["text_finan_1_total" +str(i + 1)+'_'+str(j+1)])
            for text in text_list:
                text.setText('0')
            '''Remittance set'''
            text_list=[self.ui_name_dict["text_finan_4_remit_" + str(i + 1) + '_' + str(j + 1)] for i in range(8) for j in range(3)]
            self.clear_text_list(text_list)
            button_list1=[self.ui_name_dict["button_finan_4_remitfull" + str(i + 1)] for i in range(8)]
            button_list2=[self.ui_name_dict["button_finan_4_remit" + str(i + 1)+'_'+str(j+1)] for i in range(8) for j in range(3)]
            button_list3=[self.ui_name_dict["button_finan_4_email" + str(i + 1)] for i in range(8)]
            button_list4=[self.ui_name_dict["button_finan_4_preview" + str(i + 1)] for i in range(8)]
            for button in button_list1+button_list2+button_list3+button_list4:
                button.setStyleSheet(conf["Red button style"])
            '''Bill table set'''
            service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'install', 'mis', 'mechrev', 'var']
            text_name_dict = {"mech": 'Mechanical',"cfd": 'CFD',"ele": 'Electrical',"hyd": 'Hydraulic',"fire": 'Fire',
                              "install": 'Installation',"mis": 'mech',"mechrev": 'Mechanical',"var":'Others'}
            for service in service_list:
                text_list=[self.ui_name_dict["text_finan_2_" +service+'_'+str(i)+'_'+str(j+1)] for i in range(4) for j in range(5)]
                self.clear_text_list(text_list)
                total_list=['0_6','0_7','4_1','4_2']
                text_list=[self.ui_name_dict["text_finan_2_" + service + '_' + total_list[i]] for i in range(4)]
                self.clear_text_list(text_list)
                checkbox_list=[self.ui_name_dict["checkbox_finan_2_" +service+'_'+str(i)] for i in range(4)]
                for checkbox in checkbox_list:
                    checkbox.setChecked(False)
                service_combo=text_name_dict[service]
                combo_list=[self.ui_name_dict["combobox_finan_2_" +service+'_'+str(i + 1)] for i in range(3)]
                for combobox in combo_list:
                    combobox.setCurrentText(service_combo)
                text_list=[self.ui_name_dict["text_finan_2_" + service + '_' + str(i+1) + '_5'] for i in range(3)]
                for text in text_list:
                    text.setStyleSheet("font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(170, 170, 170);")
            text_list = [self.ui.text_finan_2_total1,self.ui.text_finan_2_total2]
            self.clear_text_list(text_list)

            for service in service_list:
                button_list=[self.ui_name_dict["button_finan_2_"+service+"_upload" + str(i + 1)] for i in range(7)]
                for button in button_list:
                    button.setStyleSheet(conf["Red button style"])
            '''Fee acceptance set'''
            text_list=[self.ui_name_dict["text_finan_5_note" + str(i + 1)] for i in range(3)]
            self.clear_text_list(text_list)
            button_list=[self.ui.button_finan_5_upload,self.ui.button_finan_5_confirm]
            for button in button_list:
                button.setStyleSheet(conf["Red button style"])
            '''Profit set'''
            service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'install', 'mis', 'mechrev', 'var']
            text_list=[self.ui_name_dict["text_finan_3_" + service_list[i]] for i in range(len(service_list))]
            for text in text_list:
                text.setText('0')
            text_list=[self.ui_name_dict["text_finan_3_" + service_list[i]+'_in'] for i in range(len(service_list))]
            for text in text_list:
                text.setText('0')

            self.ui.text_finan_3_total1.setText('0')
            self.ui.text_finan_3_total2.setText('0')
            '''Lock fee and finan'''
            if self.ui.button_finan_1_lock.text()=='Unlock':
                self.f_button_lock(lock_name = "Financial")
            if self.ui.button_fee_lock.text()=='Unlock':
                self.f_button_lock(lock_name = "Fee")
            if self.ui.button_proinfo_search_unlock1.text() == 'Unlock':
                self.f_button_proinfo_search_unlock(contact_type='client')
            if self.ui.button_proinfo_search_unlock2.text() == 'Unlock':
                self.f_button_proinfo_search_unlock(contact_type='contact')

            '''4 frames'''
            service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'mechrev', 'mis','install']
            frame_scope_list=[self.ui_name_dict["frame_fee_scope_" + service_list[i]] for i in range(len(service_list))]
            frame_feedetail_list=[self.ui_name_dict["frame_fee_pro_" + service_list[i]] for i in range(len(service_list))]
            frame_invoice_list =[self.ui_name_dict["frame_finan_1_" + service_list[i]] for i in range(len(service_list))]
            frame_bill_list =[self.ui_name_dict["frame_finan_2_" + service_list[i]] for i in range(len(service_list))]
            frame_profit_list =[self.ui_name_dict["frame_finan_3_" + service_list[i]] for i in range(len(service_list))]
            for frame in frame_scope_list+frame_feedetail_list+frame_invoice_list+frame_bill_list+frame_profit_list:
                frame.setVisible(False)
            self.f_set_four_buttons()
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("DELETE FROM bridge.last_update WHERE quotation_number=%s", (quotation2logout,))
                mydb.commit()
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
        except:
            traceback.print_exc()
        finally:
            self.ui.combobox_pro_1_state.currentIndexChanged.connect(self.f_pro_state_change)
        try:
            self.search_times+=1
            print(self.search_times)
            if self.search_times == 5:
                self.search_bar_update=False
                self.thread_4.start()
            if self.search_times>10:
                self.search_times=0
                if self.search_bar_update:
                    self.parent.f_search_refresh()
                    print('Search bar refreshed')
        except:
            traceback.print_exc()
    '''0.2 - Search'''
    def show_time(self, label):
        current_time = time.time()
        time_diff = current_time - self.last_time
        time_diff_ms = round(time_diff * 1000, 2)
        current_time_formatted = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(current_time))
        print(f"{label}: {current_time_formatted} (间隔: {time_diff_ms}ms)")
        self.last_time = current_time

    def f_button_pro_search(self,quo_num):
        try:
            status_not_change = False
            email_text = self.parent.ui.findChild(QLineEdit, 'text_user_email')
            self.user_email = email_text.text()
            if self.ui.text_proinfo_1_quonum.text()!='':
                self.f_clear2default(True)
        except:
            traceback.print_exc()
        try:
            self.ui.combobox_pro_1_state.currentIndexChanged.disconnect(self.f_pro_state_change)
        except:
            pass
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM bridge.last_update WHERE quotation_number = %s", (quo_num,))
            update_database = mycursor.fetchall()
            if len(update_database)>0:
                last_update_time=str(update_database[0][2])
                user=update_database[0][1]
                time_last=datetime.strptime(last_update_time, "%Y-%m-%d %H:%M:%S")
                time_now = datetime.now()
                time_difference = time_now - time_last
                if time_difference < timedelta(minutes=30):
                    print(last_update_time)
                    message(user+' is using this project, only Copilot can be used for this project now.',self.parent.ui)
                    return
        except:
            traceback.print_exc()
        finally:
            mycursor.close()
            mydb.close()
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            self.ui.text_proinfo_1_quonum.setText(quo_num.upper())
            self.datanow ={"Client id":None,"Main contact id":None,"Xero client id":None,"Xero main contact id":None,}
            '''Project table'''
            mycursor.execute("SELECT * FROM bridge.projects WHERE quotation_number = %s", (quo_num,))
            project_database = mycursor.fetchall()
            if len(project_database)==1:
                project_database=project_database[0]
                pro_num = project_database[1]
                self.ui.text_proinfo_1_pronum.setText(pro_num)
                pro_name = project_database[2]
                self.ui.text_proinfo_1_proname.setText(pro_name)
                pro_state = project_database[3]
                self.ui.combobox_pro_1_state.setCurrentText(pro_state)
                address_to = project_database[4]
                if address_to=='Client':
                    self.radbuttongroup_proinfo_2_clienttype.button(1).setChecked(True)
                else:
                    self.radbuttongroup_proinfo_2_clienttype.button(2).setChecked(True)
                shop_name = project_database[5]
                self.ui.text_proinfo_1_shopname.setText(shop_name)
                pro_type = project_database[7]
                pro_type_list = ['Restaurant', 'Office', 'Commercial', 'Group House', 'Apartment', 'Mixed-use Complex','School', 'Others']
                try:
                    pro_type_id=pro_type_list.index(pro_type)
                    self.radiobuttongroup_proinfo_1_projecttype.button(pro_type_id+1).setChecked(True)
                except:
                    pass
                service_include_list = [project_database[i] for i in range(8, 17)]
                for i in range(len(service_include_list)):
                    if service_include_list[i]==1:
                        checkbox=self.ui_name_dict["checkbox_proinfo_1_servicettype_" + str(i+1)]
                        checkbox.setChecked(True)
                        self.f_checkboxgroup_proinfo_1_servicettype(checkbox)
                proposal_type = project_database[6]
                asana_apt=project_database[19]
                asana_basement = project_database[20]
                asana_feature = project_database[21]
                pro_note = project_database[22]
                if proposal_type=='Minor':
                    self.radiobuttongroup_proinfo_1_proposaltype.button(1).setChecked(True)
                    self.ui.text_proinfo_3_1_area.setText(asana_apt)
                    self.ui.text_proinfo_3_1_carspot.setText(asana_basement)
                    self.ui.text_proinfo_3_1_note.setText(asana_feature)
                elif proposal_type=='Major':
                    self.radiobuttongroup_proinfo_1_proposaltype.button(2).setChecked(True)
                    self.ui.text_proinfo_3_2_area.setText(asana_apt)
                    self.ui.text_proinfo_3_2_carspot.setText(asana_basement)
                    self.ui.text_proinfo_3_2_note.setText(asana_feature)
                self.ui.text_proinfo_3_2_pronote.setText(pro_note)
                self.f_proinfo_1_proposaltype_change()
                asana_id = project_database[23]
                if asana_id==None or asana_id=='None':
                    asana_id=''
                self.datanow["Asana id"]=asana_id
                asana_url = project_database[24]
                if asana_url==None or asana_url=='None':
                    asana_url=''
                self.datanow["Asana url"] = asana_url
                current_folder_address = project_database[25]
                self.datanow["Current folder"] = current_folder_address
                try:
                    email = project_database[26]
                    self.ui.text_proinfo_email.setText(email)
                    fee_accepted_date = project_database[27]
                    self.datanow["Fee accepted date"] = fee_accepted_date
                except:
                    pass
                id_list = [project_database[17], project_database[18]]
                '''Clients and main contact'''
                for i in range(len(id_list)):
                    client_id=id_list[i]
                    text_item_name=['client','contact'][i]
                    mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                    client_database = mycursor.fetchall()
                    if len(client_database) == 1:
                        client_database = client_database[0]
                        xero_client_id=client_database[1]
                        if i==0:
                            self.datanow["Client id"]=client_id
                            self.datanow["Xero client id"]=xero_client_id
                        else:
                            self.datanow["Main contact id"] = client_id
                            self.datanow["Xero main contact id"] = xero_client_id
                        ui_text_list=['name','','company','adress','abn','phone','email']
                        for j in range(7):
                            if j==1:
                                contact_type=client_database[3]
                                contact_list=['Architect','Builder','Certifier','Contractor','Developer','Engineer','Government','RDM','Strata','Supplier','Owner/Tenant','Others']
                                if contact_type=='None' or contact_type==None:
                                    contact_type_id=11
                                else:
                                    contact_type_id=contact_list.index(contact_type)
                                if i==0:
                                    self.radbuttongroup_proinfo_2_clientcontacttype.button(contact_type_id+1).setChecked(True)
                                else:
                                    self.radbuttongroup_proinfo_2_contactcontacttype.button(contact_type_id + 1).setChecked(True)
                            else:
                                content=client_database[j+2]
                                text_type=ui_text_list[j]
                                content_ui=self.ui_name_dict["text_proinfo_2_" + text_item_name+text_type]
                                content_ui.setText(content)
                '''Building features - Minor'''
                if proposal_type == 'Minor':
                    mycursor.execute("SELECT * FROM bridge.levels WHERE quotation_number = %s", (quo_num,))
                    building_feature_minor_database = mycursor.fetchall()
                    if len(building_feature_minor_database) == 5:
                        for i in range(5):
                            index = building_feature_minor_database[i][4]
                            for j in range(3):
                                content_j=str(building_feature_minor_database[i][j+1])
                                if content_j!='None':
                                    self.ui.table_proinfo_3_1_area.setItem(index, j, QTableWidgetItem(content_j))
                        self.f_set_table_style(self.ui.table_proinfo_3_1_area)
                '''Building features - Major'''
                if proposal_type == 'Major':
                    mycursor.execute("SELECT * FROM bridge.rooms WHERE quotation_number = %s", (quo_num,))
                    building_feature_major_database1 = mycursor.fetchall()
                    if len(building_feature_major_database1) == 8:
                        for i in range(8):
                            index = building_feature_major_database1[i][9]
                            for j in range(8):
                                content_j=str(building_feature_major_database1[i][j+1])
                                if content_j!='None':
                                    self.ui.table_proinfo_3_2_carpark.setItem(index, j, QTableWidgetItem(content_j))
                    self.ui.table_proinfo_3_2_carpark.itemChanged.disconnect(self.table_proinfo_3_2_carpark_func)
                    self.f_set_table_style(self.ui.table_proinfo_3_2_carpark)
                    self.ui.table_proinfo_3_2_carpark.itemChanged.connect(self.table_proinfo_3_2_carpark_func)

                    mycursor.execute("SELECT * FROM bridge.apartments WHERE quotation_number = %s", (quo_num,))
                    building_feature_major_database2 = mycursor.fetchall()
                    if len(building_feature_major_database2) == 20:
                        for i in range(20):
                            index = building_feature_major_database2[i][9]
                            for j in range(8):
                                content_j=str(building_feature_major_database2[i][j+1])
                                if content_j!='None':
                                    self.ui.table_proinfo_3_2_area.setItem(index, j, QTableWidgetItem(content_j))
                    self.ui.table_proinfo_3_2_area.itemChanged.disconnect(self.table_proinfo_3_2_area_func)
                    self.f_set_table_style(self.ui.table_proinfo_3_2_area)
                    self.ui.table_proinfo_3_2_area.itemChanged.connect(self.table_proinfo_3_2_area_func)
                    if proposal_type == 'Minor':
                        self.ui.text_proinfo_3_1_area.setText(asana_apt)
                        self.ui.text_proinfo_3_1_carspot.setText(asana_basement)
                    elif proposal_type == 'Major':
                        self.ui.text_proinfo_3_2_area.setText(asana_apt)
                        self.ui.text_proinfo_3_2_carspot.setText(asana_basement)

                '''Drawing number'''
                mycursor.execute("SELECT * FROM bridge.drawings WHERE quotation_number = %s", (quo_num,))
                drawing_num_database=mycursor.fetchall()
                if len(drawing_num_database) == 12:
                    for i in range(12):
                        index = drawing_num_database[i][4]
                        for j in range(3):
                            content_j = str(drawing_num_database[i][j + 1])
                            if content_j != 'None':
                                self.ui.table_proinfo_4_drawing.setItem(index, j, QTableWidgetItem(content_j))
                self.f_set_table_style(self.ui.table_proinfo_4_drawing)
                '''Scope of work'''
                self.f_set_scope_of_work_tables(quo_num, proposal_type)
                '''Stage table set'''
                try:
                    mycursor.execute("SELECT * FROM bridge.major_stage WHERE quotation_number = %s",(quo_num,))
                    stage_database=mycursor.fetchall()
                    for i in range(len(stage_database)):
                        stage=stage_database[i][2]
                        content=stage_database[i][3]
                        row_index=stage_database[i][4]
                        include=stage_database[i][1]
                        table=self.ui_name_dict["table_fee_stage_" + str(stage)]
                        if content!= None and content!='None':
                            table.setItem(row_index, 1, QTableWidgetItem(content))
                        checkbox = table.item(row_index, 0)
                        if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                            if include == 1:
                                checkbox.setCheckState(Qt.Checked)
                            else:
                                checkbox.setCheckState(Qt.Unchecked)
                    for i in range(4):
                        table =self.ui_name_dict["table_fee_stage_" + str(i+1)]
                        self.f_set_table_style(table)
                except:
                    pass
                mycursor.execute("SELECT * FROM bridge.fee WHERE quotation_number = %s",(quo_num,))
                fee_stage_database=mycursor.fetchall()[0]
                installation_date=fee_stage_database[1]
                installation_revision=fee_stage_database[2]
                installation_program=fee_stage_database[3]
                if installation_date!='' and installation_date!=None and installation_date!='None':
                    self.ui.text_fee_3_install_date.setText(installation_date.strftime("%d-%b-%Y"))
                if is_float(installation_revision):
                    self.ui.text_fee_3_install_rev.setText(str(installation_revision))
                if installation_program!=None or installation_program!='None':
                    self.ui.text_fee_3_install_program.setText(installation_program)
                fee_date=fee_stage_database[4]
                fee_revision=fee_stage_database[5]
                if fee_date!=None:
                    fee_date_set=datetime.strptime(str(fee_date), '%Y-%m-%d').strftime('%d-%b-%Y')
                    self.ui.text_fee_1_date.setText(str(fee_date_set))
                if fee_revision!=None:
                    self.ui.text_fee_1_revision.setText(str(fee_revision))



                self.f_set_four_buttons()
                for i in range(3):
                    fee_day=fee_stage_database[i+6]
                    if fee_day!='' and fee_day!=None and fee_day!='None':
                        text_fee_day=self.ui_name_dict["text_fee_2_period" + str(i+1)]
                        text_fee_day.setText(fee_day)
                for i in range(4):
                    stage_name=fee_stage_database[9+i*2]
                    stage_include=fee_stage_database[10+i*2]
                    if stage_name!='' and stage_name!=None and stage_name!='None':
                        text_name=self.ui_name_dict["text_stage_table_name" + str(i+1)]
                        text_name.setText(stage_name)
                    checkbox=self.ui_name_dict["checkbox_fee_stage_" + str(i+1)]
                    for checkbox_i, function in self.checkbox_to_function_2.items():
                        checkbox_i.stateChanged.disconnect(function)
                    if stage_include==1:
                        checkbox.setChecked(True)
                    else:
                        checkbox.setChecked(False)
                    for checkbox_i, function in self.checkbox_to_function_2.items():
                        checkbox_i.stateChanged.connect(function)
                area_price=fee_stage_database[17]
                apt_price=fee_stage_database[18]
                if is_float(area_price):
                    self.ui.text_fee_5_minor_price.setText(str(area_price))
                if is_float(apt_price):
                    self.ui.text_fee_5_major_price.setText(str(apt_price))
                total_area=fee_stage_database[19]
                total_apt=fee_stage_database[20]
                if is_float(total_area):
                    self.ui.text_fee_5_minor_area.setText(str(total_area))
                if is_float(total_apt):
                    self.ui.text_fee_5_major_apt.setText(str(total_apt))
                '''Fee table set'''
                try:
                    for text, function in self.text_to_function_3.items():
                        text.textChanged.disconnect(function)
                except:
                    traceback.print_exc()
                try:
                    for service in self.table_fee_4_cal_dict:
                        table=self.table_fee_4_cal_dict[service][0]
                        func=self.table_fee_4_cal_dict[service][1]
                        table.itemChanged.disconnect(func)
                except:
                    traceback.print_exc()
                mycursor.execute("SELECT * FROM bridge.fee_item WHERE quotation_number = %s",(quo_num,))
                fee_database=mycursor.fetchall()
                for item in fee_database:
                    service=item[1]
                    content=item[2]
                    amount=item[3]
                    column_index=item[4]
                    row_index=item[5]
                    name_short=conf['Service short dict'][service]
                    if content!='' and content!=None and content!='None':
                        table=self.ui_name_dict["table_fee_4_" + name_short]
                        table.setItem(row_index, 0, QTableWidgetItem(content))
                        if is_float(amount):
                            table.setItem(row_index, 1, QTableWidgetItem(str(amount)))
                            if is_float(column_index):
                                radiobutton=self.ui_name_dict["radbutton_finan_1_" + name_short+'_'+str(row_index+1)+'_'+str(column_index+1)]
                                radiobutton.setChecked(True)

                service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'install', 'mis', 'mechrev', 'var']
                for service in service_list:
                    self.f_text_finan_3_profit_calculation(item_name=service)
                for service in self.table_fee_4_cal_dict:
                    table=self.table_fee_4_cal_dict[service][0]
                    self.f_set_table_style(table)
                for service in service_list:
                    self.f_fee_4_total(table_type=service)
                self.f_radbutton_finan_change()
                try:
                    for text, function in self.text_to_function_3.items():
                        text.textChanged.connect(function)
                except:
                    traceback.print_exc()
                try:
                    for service in self.table_fee_4_cal_dict:
                        table=self.table_fee_4_cal_dict[service][0]
                        func=self.table_fee_4_cal_dict[service][1]
                        table.itemChanged.connect(func)
                except:
                    traceback.print_exc()
                '''Remittance set'''
                mycursor.execute("SELECT * FROM bridge.remittances WHERE quotation_number = %s",(quo_num,))
                remit_database=mycursor.fetchall()
                for i in range(len(remit_database)):
                    inv_num = remit_database[i][5]
                    for j in range(3):
                        remit_i=remit_database[i][j+2]
                        if is_float(remit_i):
                            text_remit_i=self.ui_name_dict["text_finan_4_remit_" + str(inv_num+1)+'_'+str(j+1)]
                            text_remit_i.setText(str(remit_i))

                mycursor.execute("SELECT * FROM bridge.remittance_upload WHERE quotation_number = %s",(quo_num,))
                remit_upload_database=mycursor.fetchall()
                for i in range(len(remit_upload_database)):
                    id=remit_upload_database[i][7]+1
                    button_onoff_list=[remit_upload_database[i][j+1] for j in range(6)]
                    button_list = [self.ui_name_dict["button_finan_4_preview" + str(id)],
                                   self.ui_name_dict["button_finan_4_email" + str(id)],
                                   self.ui_name_dict["button_finan_4_remitfull" + str(id)]]
                    for j in range(3):
                        button_list.append(self.ui_name_dict["button_finan_4_remit" + str(id)+'_'+str(j+1)])
                    for j in range(len(button_onoff_list)):
                        if button_onoff_list[j]==1:
                            button_list[j].setStyleSheet(conf["Gray button style"])
                '''Bill table set'''
                mycursor.execute("SELECT * FROM bridge.bills WHERE quotation_number = %s",(quo_num,))
                bill_database=mycursor.fetchall()
                asana_bill_list=[None for _ in range(27)]
                xero_bill_list=[None for _ in range(27)]
                client_bill_id_list=[None for _ in range(27)]
                bill_in_date_list=[None for _ in range(27)]
                bill_paid_date_list=[None for _ in range(27)]
                for i in range(len(bill_database)):
                    state=bill_database[i][2]
                    bill_num=bill_database[i][0]
                    asana_bill_id=bill_database[i][3]
                    xero_id=bill_database[i][4]
                    client_id=bill_database[i][5]
                    bill_in_date=bill_database[i][6]
                    paid_date=bill_database[i][7]
                    bill_type=bill_database[i][8]
                    service=bill_database[i][9]
                    fee_amount=bill_database[i][10]
                    no_gst=bill_database[i][11]
                    heads_up=bill_database[i][12]
                    row_index=bill_database[i][13]
                    service_short=conf['Service short dict'][service]
                    text_num=self.ui_name_dict["text_finan_2_" + service_short+'_'+str(row_index+1)+'_1']
                    text_num.setText(str(row_index))
                    combobox=self.ui_name_dict["combobox_finan_2_" + service_short+'_'+str(row_index+1)]
                    combobox.setCurrentText(bill_type)
                    text_headsup=self.ui_name_dict["text_finan_2_" + service_short+'_'+str(row_index+1)+'_3']
                    text_headsup.setText(heads_up)
                    text_amount=self.ui_name_dict["text_finan_2_" + service_short+'_'+str(row_index+1)+'_4']
                    text_amount.setText(str(fee_amount))
                    checkbox=self.ui_name_dict["checkbox_finan_2_" + service_short+'_'+str(row_index+1)]
                    if no_gst==1:
                        checkbox.setChecked(True)
                    else:
                        checkbox.setChecked(False)
                    bill_state_color_dict = {"Draft": conf['Text gray style'],"Awaiting Approval": conf['Text red style'],"Awaiting Payment": conf['Text yellow style'],"Paid": conf['Text green style'], "Voided": conf['Text purple style']}
                    style = bill_state_color_dict[state]
                    text_total_amount=self.ui_name_dict["text_finan_2_" + service_short+'_'+str(row_index+1)+'_5']
                    text_total_amount.setStyleSheet(style)
                    asana_bill_list[char2num(bill_num[-1])] = asana_bill_id
                    xero_bill_list[char2num(bill_num[-1])] = xero_id
                    client_bill_id_list[char2num(bill_num[-1])] = client_id
                    bill_in_date_list[char2num(bill_num[-1])] = bill_in_date
                    bill_paid_date_list[char2num(bill_num[-1])] = paid_date
                    if client_id!=None:
                        mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                        client_database = mycursor.fetchall()
                        if len(client_database) == 1:
                            client_database = client_database[0]
                            client_name = client_database[2]
                            name_text=self.ui_name_dict["text_finan_2_" +service_short+'_'+ str(row_index+1)+'_2']
                            name_text.setText(client_name)


                self.datanow["Asana bill id"] = asana_bill_list
                self.datanow["Xero bill"] = xero_bill_list
                self.datanow["Client bill id"] = client_bill_id_list
                self.datanow["Bill in date"] = bill_in_date_list
                self.datanow["Bill paid date"] = bill_paid_date_list

                mycursor.execute("SELECT * FROM bridge.bill_upload WHERE quotation_number = %s",(quo_num,))
                bill_upload_database=mycursor.fetchall()
                for i in range(len(bill_upload_database)):
                    service=bill_upload_database[i][1]
                    upload=bill_upload_database[i][2]
                    row_index=bill_upload_database[i][3]
                    if upload==1:
                        service_short=conf['Service short dict'][service]
                        button=self.ui_name_dict["button_finan_2_" + service_short+'_upload'+str(row_index+5)]
                        button.setStyleSheet(conf["Gray button style"])
                '''Bill supplier set'''
                mycursor.execute("SELECT * FROM bridge.supplier_fee WHERE quotation_number = %s",(quo_num,))
                supplier_database=mycursor.fetchall()
                for i in range(len(supplier_database)):
                    service=supplier_database[i][1]
                    clients=supplier_database[i][2]
                    gst=supplier_database[i][3]
                    service_short=conf['Service short dict'][service]
                    text_clients=self.ui_name_dict["text_finan_2_" + service_short+'_0_1']
                    text_clients.setText(clients)
                    checkbox=self.ui_name_dict["checkbox_finan_2_" + service_short+'_0']
                    if gst==1:
                        checkbox.setChecked(True)
                    else:
                        checkbox.setChecked(False)
                    for j in range(4):
                        text_bill_invoice=self.ui_name_dict["text_finan_2_" + service_short+'_0_'+str(j+2)]
                        version_text=supplier_database[i][j+4]
                        if version_text!=None:
                            text_bill_invoice.setText(str(version_text))
                '''bridge.supplier_upload'''
                mycursor.execute("SELECT * FROM bridge.supplier_upload WHERE quotation_number = %s",(quo_num,))
                supplier_upload_database=mycursor.fetchall()
                for i in range(len(supplier_upload_database)):
                    service_full=supplier_upload_database[i][1]
                    service_short=conf['Service short dict'][service_full]
                    for j in range(4):
                        upload=supplier_upload_database[i][j+2]
                        if upload==1:
                            button=self.ui_name_dict["button_finan_2_" + service_short + '_upload'+str(j+1)]
                            button.setStyleSheet(conf["Gray button style"])

                '''Fee acceptance set'''
                mycursor.execute("SELECT * FROM bridge.fee_acceptance WHERE quotation_number = %s",(quo_num,))
                supplier_database=mycursor.fetchall()
                if len(supplier_database)==1:
                    supplier_database=supplier_database[0]
                    for i in range(3):
                        note=supplier_database[i+1]
                        if note!='' and note!=None and note!='None':
                            text_note=self.ui_name_dict["text_finan_5_note" + str(i+1)]
                            text_note.setText(note)
                '''bridge.acceptance_upload'''

                mycursor.execute("SELECT * FROM bridge.acceptance_upload WHERE quotation_number = %s",(quo_num,))
                acceptance_upload_database=mycursor.fetchall()
                if len(acceptance_upload_database)==1:
                    v1_upload=acceptance_upload_database[0][1]
                    v1_confirm=acceptance_upload_database[0][2]
                    if v1_upload==1:
                        self.ui.button_finan_5_upload.setStyleSheet(conf["Gray button style"])
                    if v1_confirm==1:
                        self.ui.button_finan_5_confirm.setStyleSheet(conf["Gray button style"])
                '''Email date set'''

                mycursor.execute("SELECT * FROM bridge.emails WHERE quotation_number = %s",(quo_num,))
                emails_database=mycursor.fetchall()
                if len(emails_database)==1:
                    emails_database=emails_database[0]
                    fee_proposal_date=emails_database[2]
                    fee_rev=emails_database[6]
                    if fee_proposal_date!=None:
                        new_date=datetime.strptime(str(fee_proposal_date), '%Y-%m-%d').strftime('%d-%b-%Y')
                        self.ui.text_fee_1_date.setText(new_date)

                        if pro_state in ['Set Up','Gen Fee Proposal','Email Fee Proposal']:
                            try:
                                pro_info = quo_num + '-' + str(pro_num) + '-' + pro_name
                                color_tuple = ("green", "green", "green", "yellow")
                                self.f_set_pro_buttons_color(color_tuple)
                                self.ui.combobox_pro_1_state.setCurrentText("Chase Fee Acceptance")
                                status_not_change=True
                                update_asana = messagebox('Update Asana', "Do you want to update Asana?",
                                                          self.parent.ui)
                                if update_asana:
                                    update_asana_result = self.f_button_pro_func_updateasana(False)
                                    if update_asana_result == 'Success':
                                        message(f'Project {pro_info} \nAsana Updated Successfully', self.parent.ui)
                                    elif update_asana_result == 'Fail':
                                        message(f'Project {pro_info} \nUpdate Asana failed, please contact admin',
                                                self.parent.ui)
                                    else:
                                        message(update_asana_result, self.parent.ui)
                            except:
                                traceback.print_exc()

                '''Invoice set'''

                mycursor.execute("SELECT * FROM bridge.invoices WHERE quotation_number = %s",(quo_num,))
                invoice_database=mycursor.fetchall()
                asana_invoice_list=[None for _ in range(8)]
                xero_invoice_list=[None for _ in range(8)]
                client_invoice_list=[None for _ in range(8)]
                for i in range(len(invoice_database)):
                    invoice_number=invoice_database[i][0]
                    state=invoice_database[i][2]
                    asana_invoice_id=invoice_database[i][3]
                    xero_id=invoice_database[i][4]
                    client_id=invoice_database[i][5]
                    payment_date=invoice_database[i][6]
                    payment_amount=invoice_database[i][7]
                    fee_amount=invoice_database[i][8]
                    row_index=invoice_database[i][9]
                    text_inv_num=self.ui_name_dict["text_finan_1_inv" + str(row_index+1)]
                    finan_text_list=[self.ui_name_dict["text_finan_1_total" + str(row_index+1)+'_'+str(j+1)] for j in range(5)]
                    text_inv_num.setText(invoice_number)
                    invoice_state_color_dict = {"Backlog": conf['Text gray style'],
                                                "Sent":conf['Text red style'],
                                                "Paid":conf['Text green style'],
                                                "Voided":conf['Text purple style'],}
                    invoice_state_color_dict2 = {"Backlog": conf['Text gray_white style'],
                                                "Sent":conf['Text red_white style'],
                                                "Paid":conf['Text green_white style'],
                                                "Voided":conf['Text purple_white style'],}
                    style = invoice_state_color_dict[state]
                    style2 = invoice_state_color_dict2[state]
                    if text_inv_num.text().find('-')==-1:
                        text_inv_num.setStyleSheet(style)
                        for finan_text in finan_text_list:
                            finan_text.setStyleSheet(style2)
                    asana_invoice_list[row_index]=asana_invoice_id
                    xero_invoice_list[row_index]=xero_id
                    client_invoice_list[row_index]=client_id
                    if payment_amount!='' and payment_amount!=None and payment_amount!='None':
                        text_payment_amount=self.ui_name_dict["text_finan_1_total" + str(row_index + 1)+'_3']
                        text_payment_amount.setText(str(payment_amount))
                    if payment_date!='' and payment_date!=None and payment_date!='None':
                        text_payment_date=self.ui_name_dict["text_finan_1_total" + str(row_index + 1)+'_4']
                        text_payment_date.setText(str(payment_date))

                for i in range(len(client_invoice_list)):
                    client_id=client_invoice_list[i]
                    if client_id!=None:
                        mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                        client_database = mycursor.fetchall()
                        if len(client_database) == 1:
                            client_database = client_database[0]
                            client_name = client_database[2]
                            company_name=client_database[4]
                            name_text=self.ui_name_dict["text_finan_1_invclient_" + str(i + 1)]
                            client_company_name=str(client_name)+'--'+str(company_name)
                            name_text.setText(client_company_name)
                self.ui.table_finan_inv_client_search.hide()

                self.datanow["Asana invoice id"]=asana_invoice_list
                self.datanow["Xero invoice id"] = xero_invoice_list
                self.datanow["Client invoice id"] = client_invoice_list
                '''User history set'''

                mycursor.execute("SELECT * FROM bridge.user_history WHERE quotation_number = %s", (quo_num,))
                history_database = mycursor.fetchall()
                if len(history_database)!=0:
                    for i in range(len(history_database)):
                        user_email=history_database[i][1]
                        action_time=history_database[i][2]
                        action_content=history_database[i][3]
                        table_row_now = self.ui.table_proinfo_history.rowCount()
                        self.ui.table_proinfo_history.insertRow(table_row_now)
                        self.ui.table_proinfo_history.setItem(i, 0, QTableWidgetItem(user_email))
                        self.ui.table_proinfo_history.setItem(i, 1, QTableWidgetItem(str(action_time)))
                        self.ui.table_proinfo_history.setItem(i, 2, QTableWidgetItem(str(action_content)))
                self.f_set_table_style(self.ui.table_proinfo_history)
                self.f_set_table_style(self.ui.table_fee_5_cal_carpark)
                '''Button status'''

                mycursor.execute("SELECT * FROM bridge.status_colour WHERE quotation_number = %s",(quo_num,))
                status_database=mycursor.fetchall()
                color_dict={0:'red',1:'yellow',2:'green'}
                state=self.ui.combobox_pro_1_state.currentText()
                if state=='Set Up':
                    color_setup='yellow'
                else:
                    color_setup = 'green'
                if status_not_change==False:
                    if len(status_database)==1:
                        status_database=status_database[0]
                        gen_color=status_database[1]
                        email_color=status_database[2]
                        chase_color=status_database[3]
                        color_tuple=(color_setup,color_dict[gen_color],color_dict[email_color],color_dict[chase_color])
                        self.f_set_pro_buttons_color(color_tuple)
                '''tables to hide'''
                self.ui.table_client_search.hide()
                self.ui.table_contact_search.hide()
                self.ui.table_finan_inv_client_search.hide()
                self.ui.table_finan_bill_client_search.hide()

                self.thread_3.start()

            else:
                message("No such project.",self.parent.ui)
                return



        except:
            traceback.print_exc()
        finally:
            mycursor.close()
            mydb.close()
            self.ui.combobox_pro_1_state.currentIndexChanged.connect(self.f_pro_state_change)
        try:
            state = self.ui.combobox_pro_1_state.currentText()
            if state not in ['Set Up', 'Gen Fee Proposal','Email Fee Proposal']:
                if self.ui.button_fee_lock.text() == 'Lock':
                    self.f_button_lock(lock_name="Fee")
                if self.ui.button_proinfo_search_unlock1.text() == 'Lock':
                    self.f_button_proinfo_search_unlock(contact_type='client')
                if self.ui.button_proinfo_search_unlock2.text() == 'Lock':
                    self.f_button_proinfo_search_unlock(contact_type='contact')


        except:
            traceback.print_exc()

    '''1 - Set four buttons color'''
    def f_set_pro_buttons_color(self,color_tuple):
        try:
            button_list = [self.ui.button_pro_state_setup, self.ui.button_pro_state_genfee,self.ui.button_pro_state_emailfee, self.ui.button_pro_state_chasefee]
            button_style = {"yellow": "QPushButton {background-color: rgb(255, 165, 0);color: rgb(0, 0, 0);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;padding: 6px;} QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}",
                            "green": "QPushButton {background-color: rgb(0, 128, 0);color: rgb(0, 0, 0);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}",
                            "red": "QPushButton {background-color: rgb(165, 42, 42);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}", }
            for i in range(4):
                button = button_list[i]
                color = color_tuple[i]
                button.setStyleSheet(button_style[color])
        except:
            traceback.print_exc()
    '''1.1 - Set up'''
    def f_button_pro_state_setup(self):
        try:
            quo_num=self.ui.text_proinfo_1_quonum.text()
            if quo_num=='':
                message("Please create a quotation number first",self.parent.ui)
                return
            current_folder_address=self.datanow["Current folder"]
            pro_name=self.ui.text_proinfo_1_proname.text()
            if pro_name=='':
                message('Please insert project name',self.parent.ui)
                return
            if pro_name[0] == " " or pro_name[-1] == " ":
                message("You cant have empty space in front or end of the project name",self.parent.ui)
                return
            special_char_list = ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]
            for char in pro_name:
                if char in special_char_list:
                    message("You cant have special character in your project name",self.parent.ui)
                    return
            if self.radiobuttongroup_proinfo_1_proposaltype.checkedId()== -1:
                message('Please select proposal type',self.parent.ui)
                return
            if self.radiobuttongroup_proinfo_1_projecttype.checkedId()== -1:
                message('Please select project type',self.parent.ui)
                return
            service_list=self.f_get_pro_service_type('Set up')
            if len(service_list)==0:
                message('Please choose service type',self.parent.ui)
                return
            if self.ui.text_proinfo_2_clientname.text()=='' and self.ui.text_proinfo_2_contactname.text()=='':
                message('Please choose contact type',self.parent.ui)
                return
            if self.radiobuttongroup_proinfo_1_proposaltype.checkedId() == 1:
                if self.ui.text_proinfo_3_1_area.text()=='':
                    message('Please insert building features',self.parent.ui)
                    return
            else:
                if self.ui.text_proinfo_3_2_area.text()=='' and self.ui.text_proinfo_3_2_carspot.text()=='':
                    message('Please insert building features',self.parent.ui)
                    return
            folder_name=os.path.join(conf["working_dir"],quo_num+ "-" +pro_name)
            if self.ui.combobox_pro_1_state.currentText()!='Set Up':
                message("You already setup the project",self.parent.ui)
                return
            if current_folder_address!=folder_name:
                try:
                    os.rename(current_folder_address, folder_name)
                    self.datanow["Current folder"]=folder_name
                    message(f"Project {quo_num} set up successfully",self.parent.ui)
                except:
                    traceback.print_exc()
                    message("Please Close the file relate to folder before rename",self.parent.ui)
                    return
            self.ui.combobox_pro_1_state.setCurrentText("Gen Fee Proposal")
            self.f_add_user_history('Finish Set Up')
            color_tuple=("green","yellow", "red", "red",)
            self.f_set_pro_buttons_color(color_tuple)
        except:
            traceback.print_exc()
    '''1.2 - Gen fee proposal'''
    def f_button_pro_state_genfee(self):
        try:
            rev_same = False
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            check_or_create_folder(accounting_dir)
            proposal_type=self.get_proposal_type()
            pro_name=self.ui.text_proinfo_1_proname.text()
            fee_rev = self.ui.text_fee_1_revision.text()
            service_type_list = self.f_get_pro_service_type('Fee proposal')
            if len(service_type_list) == 0:
                message("Please at least select 1 service",self.parent.ui)
                return
            if len(service_type_list) > 5:
                message("Excess the maximum value of service, please contact administrator",self.parent.ui)
                return
            state=self.ui.combobox_pro_1_state.currentText()
            if state=='Set Up':
                message("Please finish Set Up first",self.parent.ui)
                return
            if is_float(self.ui.text_fee_4_ex.text()):
                if float(self.ui.text_fee_4_ex.text())<=0:
                    message("There are error in the fee proposal section, please fix the fee section before generate the fee proposal",self.parent.ui)
                    return
            else:
                message("There are error in the fee proposal section, please fix the fee section before generate the fee proposal",self.parent.ui)
                return
            total_fee = self.ui.text_fee_4_ex.text()
            total_ist = self.ui.text_fee_4_in.text()
            service_name = self.f_get_service_name_for_pdf()
            pdf_name=f'{service_name} Fee Proposal for {pro_name} Rev {fee_rev}'+".pdf"
            pdf_dir=os.path.join(accounting_dir,pdf_name)
            excel_name = f'{pro_name} Back Up.xlsx'
            excel_dir=os.path.join(accounting_dir,excel_name)
            contact_type = self.get_contact_type()
            if contact_type == 'Client':
                full_name = self.ui.text_proinfo_2_clientname.text()
                company = self.ui.text_proinfo_2_clientcompany.text()
                address = self.ui.text_proinfo_2_clientadress.text()
            else:
                full_name = self.ui.text_proinfo_2_contactname.text()
                company = self.ui.text_proinfo_2_contactcompany.text()
                address = self.ui.text_proinfo_2_contactadress.text()
            first_name = get_first_name(full_name)
            quo_num = self.ui.text_proinfo_1_quonum.text()
            email_fee_proposal_date = self.ui.text_fee_1_date.text()
            time_fee_proposal_Fee_proposal = self.ui.text_fee_2_period1.text()
            time_fee_proposal_Pre_design = self.ui.text_fee_2_period2.text()
            time_fee_proposal_Documentation = self.ui.text_fee_2_period3.text()
            project_type = self.f_get_project_type()
            pro_note = self.ui.text_proinfo_3_2_pronote.toPlainText()
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.past_project WHERE project_type = %s", (project_type,))
                project_history_database = mycursor.fetchall()
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
            past_projects=[]
            if len(project_history_database) > 0:
                for i in range(len(project_history_database)):
                    content = project_history_database[i][1]
                    past_projects.append(content)
            services_content_dict={}
            for i in range(len(service_type_list)):
                service_type = service_type_list[i]
                service_type_short = conf['Service short dict'][service_type]
                service_content_i_dict={}
                for j in range(3):
                    content_type = ['Extent', 'Clarifications', 'Deliverables'][j]
                    table=self.ui_name_dict["table_fee_3_" + service_type_short + '_' + str(j + 1)]
                    content_list = []
                    for row in range(table.rowCount()):
                        checkbox = table.item(row, 0)
                        if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                            if checkbox.checkState() == Qt.Checked:
                                content = table.item(row, 1).text()
                                content_list.append(content)
                    service_content_i_dict[content_type]=content_list
                table_fee=self.ui_name_dict["table_fee_4_" + service_type_short]
                fee = table_fee.item(table_fee.rowCount() - 1, 1).text()
                fee_total = table_fee.item(table_fee.rowCount() - 1, 2).text()
                service_content_i_dict['Fee']=[fee,fee_total]
                services_content_dict[service_type]=service_content_i_dict
                if not is_float(fee):
                    message('Please go to fee proposal page to complete the fee first',self.parent.ui)
                    return
                if float(fee)==0:
                    message('Please go to fee proposal page to complete the fee first',self.parent.ui)
                    return
            pdf_list = [file for file in os.listdir(accounting_dir) if "Fee Proposal" in str(file) and str(file).endswith(".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]
            if len(pdf_list) != 0:
                current_pdf = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
                current_revision = current_pdf.split(" ")[-1].split(".")[0]
                if fee_rev == current_revision or fee_rev == str(int(current_revision) + 1):
                    old_pdf_path = os.path.join(accounting_dir, current_pdf)
                    open_pdf_thread(old_pdf_path)
                    if fee_rev == current_revision:
                        overwrite = messagebox("Warning",f"Revision {current_revision} found, do you want to overwrite?", self.parent.ui)
                        rev_same=True
                    else:
                        overwrite = messagebox("Warning",f"Revision {current_revision} found, do you want to generate Revision {fee_rev}", self.parent.ui)
                    if not overwrite:
                        return
                    else:
                        try:
                            for proc in psutil.process_iter():
                                if proc.name() == "Acrobat.exe":
                                    proc.kill()
                        except:
                            pass
                else:
                    message(f'Current revision is {current_revision}, you can not use revision {fee_rev}',self.parent.ui)
                    return
            else:
                if fee_rev != "1":
                    message("There is no other existing fee proposal found, can only have revision 1",self.parent.ui)
                    return

            if proposal_type=="Minor Project":
                shutil.copy(os.path.join(conf["template_dir"], "xlsx", f"fee_proposal_minor_template.xlsx"),excel_dir)
                email_fee_proposal_date=convert_date(email_fee_proposal_date)
                write_excel=f_write_minor_fee_excel(excel_dir,full_name,company,address,first_name,quo_num,email_fee_proposal_date,fee_rev,pro_name,
                            time_fee_proposal_Fee_proposal,time_fee_proposal_Pre_design,time_fee_proposal_Documentation,
                            service_type_list,services_content_dict,past_projects,total_fee,total_ist)
                if not write_excel:
                    message(write_excel,self.parent.ui)
                    return
                export_excel_thread(excel_dir, pdf_dir)
            else:
                stage_content_dict={}
                stage_name_list=[]
                for i in range(4):
                    checkbox=self.ui_name_dict["checkbox_fee_stage_" + str(i+1)]
                    if checkbox.isChecked():
                        content_list=[]
                        table=self.ui_name_dict["table_fee_stage_" + str(i+1)]
                        for row in range(table.rowCount()):
                            checkbox = table.item(row, 0)
                            if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                                if checkbox.checkState() == Qt.Checked:
                                    content = table.item(row, 1).text()
                                    content_list.append(content)
                        stage_name=self.ui_name_dict["text_stage_table_name" + str(i+1)].text()
                        stage_content_dict[stage_name]=content_list
                        stage_name_list.append(stage_name)
                if stage_content_dict == {}:
                    message("You need to select at least one stage",self.parent.ui)
                    return
                stage_fee_dict={}
                for service_type in service_type_list:
                    fee_list=[]
                    service_type_short = conf['Service short dict'][service_type]
                    table=self.ui_name_dict["table_fee_4_" + service_type_short]
                    for i in range(len(stage_name_list)):
                        fee_i = table.item(i, 1).text()
                        if is_float(fee_i):
                            fee_i = float(fee_i)
                        else:
                            fee_i = 0
                        fee_list.append(fee_i)
                    stage_fee_dict[service_type]=fee_list
                shutil.copy(os.path.join(conf["template_dir"], "xlsx", f"fee_proposal_major_template.xlsx"), excel_dir)
                email_fee_proposal_date=convert_date(email_fee_proposal_date)
                write_excel=f_write_major_fee_excel(excel_dir,pro_name,service_name,full_name,company,address,email_fee_proposal_date,quo_num,
                            first_name,service_type_list,pro_note,stage_content_dict,services_content_dict,past_projects,
                            stage_name_list,stage_fee_dict,fee_rev)
                if not write_excel:
                    message(write_excel,self.parent.ui)
                    return
                export_excel_thread(excel_dir, pdf_dir)
            action_content='Create Fee Proposal Revision '+fee_rev
            self.f_add_user_history(action_content)
            state=self.ui.combobox_pro_1_state.currentText()
            if not rev_same:
                if state in ['Set Up', 'Gen Fee Proposal','Email Fee Proposal','Chase Fee Acceptance']:
                    self.ui.combobox_pro_1_state.setCurrentText("Email Fee Proposal")
                color_tuple = ("green","green","yellow", "red")
                self.f_set_pro_buttons_color(color_tuple)
        except:
            traceback.print_exc()
            message("Unexpected files in the database, Please put the external files in a SS folder",self.parent.ui)
    '''1.3 - Email fee proposal'''
    def f_button_pro_state_emailfee(self):
        try:
            state=self.ui.combobox_pro_1_state.currentText()
            if state in ['Set Up', 'Gen Fee Proposal']:
                message("Please Generate a pdf first, or delete fee and regenerate.",self.parent.ui)
                return
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            check_or_create_folder(accounting_dir)
            pdf_list = [file for file in os.listdir(accounting_dir) if
                        "Fee Proposal" in str(file) and str(file).endswith(
                            ".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]
            pdf_name = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
            service_name=self.f_get_service_name_for_pdf()
            fee_rev = self.ui.text_fee_1_revision.text()
            pro_name = self.ui.text_proinfo_1_proname.text()
            service_name_for_email = f'{service_name} Fee Proposal for {pro_name} Rev {fee_rev}'

            contact_type = self.get_contact_type()
            if contact_type == 'Client':
                full_name = self.ui.text_proinfo_2_clientname.text()
            else:
                full_name = self.ui.text_proinfo_2_contactname.text()
            first_name = get_first_name(full_name)
            client_email=self.ui.text_proinfo_2_clientemail.text()
            main_contact_email=self.ui.text_proinfo_2_contactemail.text()
            if client_email=='0@com.au':
                client_email=''
            if main_contact_email=='0@com.au':
                main_contact_email=''
            pdf_dir = os.path.join(accounting_dir, pdf_name)
            quo_num=self.ui.text_proinfo_1_quonum.text()
            message2email = f"""
            Dear {first_name},<br>
            <br>
            I hope this email finds you well. Please find the fee proposal attached for your review and approval.<br>
            <br>
            If you have any questions or need more information, please feel free to give us a call.<br>
            <br>
            We look forward to contributing to the project.<br>
            """
            subject = f'{quo_num}-{service_name_for_email}'
            email2cc = ''
            self.thread_1.start()
            f_email_fee(subject,client_email,main_contact_email,email2cc,message2email,pdf_dir)
        except:
            message("Unable to Create Email",self.parent.ui)
    '''1.4 - Chase fee proposal'''
    def f_button_pro_state_chasefee(self):
        try:
            state = self.ui.combobox_pro_1_state.currentText()
            if state in ['Set Up', 'Gen Fee Proposal', 'Email Fee Proposal']:
                message("Please Sent the fee proposal to client first",self.parent.ui)
                return
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            pdf_list = [file for file in os.listdir(os.path.join(accounting_dir)) if
                        "Fee Proposal" in str(file) and str(file).endswith(
                            ".pdf") and not "Mechanical Installation Fee Proposal" in str(file)]
            pdf_name = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
            contact_type = self.get_contact_type()
            if contact_type == 'Client':
                full_name = self.ui.text_proinfo_2_clientname.text()
            else:
                full_name = self.ui.text_proinfo_2_contactname.text()
            first_name = get_first_name(full_name)
            pro_name = self.ui.text_proinfo_1_proname.text()
            client_email = self.ui.text_proinfo_2_clientemail.text()
            main_contact_email = self.ui.text_proinfo_2_contactemail.text()
            if client_email=='0@com.au':
                client_email=''
            if main_contact_email=='0@com.au':
                main_contact_email=''
            message2email = f"""
            Dear {first_name},<br>
            <br>
            I hope this email finds you well.<br>
            <br>
            I am writing to follow up on the fee proposal we sent on <b>{convert_time_format(self.ui.text_fee_1_date.text())}, attached</b>. I wanted to check if there is any update. <br>
            <br>
            If you need any further clarification, please do not hesitate to contact us.<br>
            <br>
            Looking forward to hearing from you and contribute to the project.<br>
            <br>
            """
            quo_num = self.ui.text_proinfo_1_quonum.text()
            subject = f'Re: {quo_num}-{pro_name}'
            email2cc = "felix@pcen.com.au"
            pdf_dir = os.path.join(accounting_dir, pdf_name)
            f_email_fee(subject, client_email, main_contact_email, email2cc, message2email, pdf_dir)
            self.f_add_user_history("Sent a chase email to Client")
        except:
            traceback.print_exc()
    '''2.1~2.3 - Open folder'''
    def f_profolder_open(self,folder_name):
        try:
            if self.ui.text_proinfo_1_quonum.text()=='':
                message('Please login to a project first',self.parent.ui)
                return
            backup_folder=os.path.join(conf["backup_folder"],Path(self.datanow["Current folder"]).name)
            folder_name_dict={"Project":self.datanow["Current folder"],
                              "Backup":backup_folder,
                              "Database":os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())}
            folder_dir=folder_name_dict[folder_name]
            if folder_name=="Backup":
                create_directory(folder_dir)
            if file_exists(folder_dir):
                open_folder(folder_dir)
            else:
                message(f"{folder_dir} does not exist",self.parent.ui)
                return
        except:
            traceback.print_exc()
    '''2.4 - Get filename of design certificate'''
    def f_get_filename(self,folder):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.ReadOnly
            try:
                fileName, _ = QFileDialog.getOpenFileName(self.ui, "QFileDialog.getOpenFileName()",folder,"", options=options)
            except:
                fileName, _ = QFileDialog.getOpenFileName(self.ui, "QFileDialog.getOpenFileName()", "","", options=options)
            return fileName
        except:
            traceback.print_exc()
    '''2.4 - Design certificate'''
    def f_button_pro_func_designcert(self):
        try:
            quo_num=self.ui.text_proinfo_1_quonum.text()
            if quo_num=='':
                message("Please enter a Quotation Number before you load",self.parent.ui)
                return
            pro_name = self.ui.text_proinfo_1_proname.text()
            drawing_content = []
            table = self.ui.table_proinfo_4_drawing
            for i in range(table.rowCount()):
                try:
                    drawing_number = table.item(i, 0).text()
                    drawing_name = table.item(i, 1).text()
                    revision = table.item(i, 2).text()
                    drawing_i = {"Drawing number": drawing_number,
                                 "Drawing name": drawing_name,
                                 "Revision": revision, }
                    drawing_content.append(drawing_i)
                except:
                    pass
            folder_path = self.datanow["Current folder"]
            proposal_type=self.get_proposal_type()
            excel_path = None
            for file in os.listdir(folder_path):
                if file.endswith(".xlsx"):
                    if file.startswith("Preliminary Calculation") or file.startswith("Mech System Calculation"):
                        excel_path = os.path.join(folder_path,file)
                        break
            if excel_path is None:
                message(f"Can not find the Preliminary Calculation or Mech System Calculation file under {folder_path}",self.parent.ui)
                return

            if check_excel_open(excel_path):
                message(f"The excel \n{excel_path}\n is open, please close it.",self.parent.ui)
                return

            excel = win32client.Dispatch("Excel.Application")
            try:
                excel.Visible = True
            except Exception as e:
                print(e)
            work_book = excel.Workbooks.Open(excel_path)
            work_book.Worksheets[0].Activate()
            design_certificate_worksheet = work_book.Worksheets["Mechanical Design Certificate"]
            design_compliance_cert_work_sheet = work_book.Worksheets["Mech Design Compliance Cert"]
            design_certificate_worksheet.Cells(2, 1).Value = pro_name
            design_compliance_cert_work_sheet.Cells(2, 1).Value = pro_name
            i = 1
            design_certificate_first_cell = None
            while i < 100:
                if design_certificate_worksheet.Cells(i, 2).Value == "Drawing Number":
                    design_certificate_first_cell = i + 1
                    break
                i += 1
            if design_certificate_first_cell is None:
                message("Can not found the drawing row of the design certificate",self.parent.ui)
                return
            i = 1
            design_compliance_first_cell = None
            while i < 100:
                if design_compliance_cert_work_sheet.Cells(i, 2).Value == "Drawing Number":
                    design_compliance_first_cell = i + 1
                    break
                i += 1
            if design_compliance_cert_work_sheet is None:
                message("Can not found the drawing row of the design compliance certificate",self.parent.ui)
                return
            if drawing_content != []:
                for i in range(len(drawing_content)):
                    drawing_number = drawing_content[i]["Drawing number"]
                    drawing_name = drawing_content[i]["Drawing name"]
                    drawing_rev = drawing_content[i]["Revision"]
                    design_certificate_worksheet.Cells(design_certificate_first_cell + i, 2).Value = drawing_number
                    design_certificate_worksheet.Cells(design_certificate_first_cell + i, 4).Value = drawing_name
                    design_certificate_worksheet.Cells(design_certificate_first_cell + i, 8).Value = drawing_rev
                    design_compliance_cert_work_sheet.Cells(design_compliance_first_cell + i, 2).Value = drawing_number
                    design_compliance_cert_work_sheet.Cells(design_compliance_first_cell + i, 4).Value = drawing_name
                    design_compliance_cert_work_sheet.Cells(design_compliance_first_cell + i, 8).Value = drawing_rev


            design_cert_dict = {
                "Design": "Mechanical Design Certificate.pdf",
                "Compliance": 'Mechanical Design Compliance Certificate.pdf'}
            dialog = CustomDialog(self.parent.ui)
            if dialog.exec_() == QDialog.Accepted:
                selected_option = dialog.get_selected_option()
                if selected_option!=None:
                    pdf_name=design_cert_dict[selected_option]
                    pdf_path = os.path.join(folder_path, "Plot", pdf_name)
                    if os.path.exists(pdf_path):
                        open_pdf_thread(pdf_path)
                        overwrite=messagebox("File existing",f'{pdf_name} \n exists, do you want to overwrite?', self.parent.ui)
                        if not overwrite:
                            return
                    for proc in psutil.process_iter():
                        if proc.name() == "Acrobat.exe":
                            proc.kill()
                    if selected_option=="Design":
                        design_certificate_worksheet.ExportAsFixedFormat(0, pdf_path)
                        open_pdf_thread(pdf_path)
                    else:
                        design_compliance_cert_work_sheet.ExportAsFixedFormat(0, pdf_path)
                        open_pdf_thread(pdf_path)
        except:
            message('Can not generate design certificate, please end excel task at Task Manager.',self.parent.ui)
            traceback.print_exc()


    '''3.1 - Rename project'''
    def f_button_pro_func_renameproject(self):
        try:
            quo_num=self.ui.text_proinfo_1_quonum.text()
            pro_num=self.ui.text_proinfo_1_pronum.text()
            pro_name=self.ui.text_proinfo_1_proname.text()
            pro_info=quo_num+'-'+pro_num+'-'+pro_name
            if pro_num!='':
                folder_name=pro_num+'-'+pro_name
            else:
                folder_name=self.ui.text_proinfo_1_quonum.text()+'-'+pro_name
            folder_address=os.path.join(conf["working_dir"], folder_name)
            if self.ui.text_proinfo_1_quonum.text()=='':
                message("Please Create an Quotation Number first",self.parent.ui)
                return
            elif os.path.exists(folder_address):
                message(f"{folder_address} exists, you dont need to rename it",self.parent.ui)
                return
            if pro_name[0] == " " or pro_name[-1] == " ":
                message("You cant have empty space in front or end of the project name",self.parent.ui)
                return
            special_char_list = ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]
            for char in pro_name:
                if char in special_char_list:
                    message("You cant have special character in your project name",self.parent.ui)
                    return
            try:
                old_folder=self.datanow["Current folder"]
                os.rename(old_folder, folder_address)
                message(f"Rename Folder from \n {old_folder} \n To \n {folder_address}",self.parent.ui)
                self.datanow["Current folder"]=folder_name
                asana_id = self.datanow["Asana id"]
                if len(asana_id) != 0:
                    update_asana_result=self.f_button_pro_func_updateasana(False)
                    if update_asana_result == 'Success':
                        message(f'Project {pro_info} \nAsana Updated Successfully', self.parent.ui)
                    elif update_asana_result == 'Fail':
                        message(f'Project {pro_info} \nUpdate Asana failed, please contact admin', self.parent.ui)
                    else:
                        message(update_asana_result, self.parent.ui)
                open_folder(self.datanow["Current folder"])
                self.f_add_user_history('Rename the Folder')
            except:
                traceback.print_exc()
                message("Bridge wont able to rename the folder, Please close anything related to the folder",self.parent.ui)
        except:
            traceback.print_exc()
    '''3.2 - Delete project'''
    def f_button_proinfo_deletepro(self):
        try:
            if self.ui.text_proinfo_1_quonum.text()=='':
                message("You can't delete an empty project",self.parent.ui)
                return
            messagebox_answer=messagebox("Delete","Are you sure you want to delete the Project?", self.parent.ui)
            if messagebox_answer:
                recycle_folder=self.ui.text_proinfo_1_quonum.text()+ "-" + datetime.now().strftime("%y%m%d%H%M%S")
                recycle_bin_dir = os.path.join(conf["recycle_bin_dir"], recycle_folder)
                os.mkdir(recycle_bin_dir)
                if os.path.exists(self.datanow["Current folder"]):
                    shutil.move(self.datanow["Current folder"], recycle_bin_dir)
                    message(f"Project {self.ui.text_proinfo_1_quonum.text()} deleted, and you have to delete Asana task manually.",self.parent.ui)
                    quo_num=self.ui.text_proinfo_1_quonum.text()
                    delete_project_database(quo_num)
                    self.f_clear2default(False)
        except:
            traceback.print_exc()
    '''4.1 - Update Asana'''
    def f_button_pro_func_updateasana_button(self, update_sub):
        try:
            quo_num=self.ui.text_proinfo_1_quonum.text()
            pro_num=self.ui.text_proinfo_1_pronum.text()
            pro_name=self.ui.text_proinfo_1_proname.text()
            pro_info=quo_num+'-'+pro_num+'-'+pro_name
            update_asana_result=self.f_button_pro_func_updateasana(update_sub)
            if update_asana_result == 'Success':
                message(f'Project {pro_info} \nAsana Updated Successfully', self.parent.ui)
            elif update_asana_result == 'Fail':
                message(f'Project {pro_info} \nUpdate Asana failed, please contact admin', self.parent.ui)
            else:
                message(update_asana_result, self.parent.ui)
        except:
            message('Update Asana failed, please contact admin', self.parent.ui)

    def f_button_pro_func_updateasana(self,update_sub):
        try:
            self.f_button_pro_func_refreshxero()
            project_type=self.f_get_project_type()
            contact_type=self.get_contact_type()
            project_state=self.ui.combobox_pro_1_state.currentText()
            service_type_list=self.f_get_pro_service_type('Asana')
            shop_name=self.ui.text_proinfo_1_shopname.text()
            if service_type_list==[]:
                return "Choose at least one service type"
            proposal_type=self.get_proposal_type()
            if proposal_type=='Minor Project':
                asana_area=self.ui.text_proinfo_3_1_area.text()
                asana_basement=self.ui.text_proinfo_3_1_carspot.text()
                asana_note=self.ui.text_proinfo_3_1_note.text()
            else:
                asana_area=self.ui.text_proinfo_3_2_area.text()
                asana_basement=self.ui.text_proinfo_3_2_carspot.text()
                asana_note=self.ui.text_proinfo_3_2_note.text()

            if is_float(self.ui.text_fee_4_ex.text()):
                total_fee=float(self.ui.text_fee_4_ex.text())
            else:
                total_fee=0
            if is_float(self.ui.text_finan_1_total0_3.text()):
                total_paid_fee=float(self.ui.text_finan_1_total0_3.text())
            else:
                total_paid_fee=0
            overdue_fee=self.f_get_overdue_fee()

            client_name_list=[self.ui.text_proinfo_2_clientcompany.text(), self.ui.text_proinfo_2_clientname.text()]
            client_name= "-".join(client_name_list)
            main_contact_list=[self.ui.text_proinfo_2_contactcompany.text(),self.ui.text_proinfo_2_contactname.text()]
            main_contact_name = "-".join(main_contact_list)

            pro_contact_type=self.get_contact_type()
            if pro_contact_type=='Client':
                con_type = self.radbuttongroup_proinfo_2_clientcontacttype.checkedId()
            else:
                con_type = self.radbuttongroup_proinfo_2_contactcontacttype.checkedId()
            con_type_list = ['Architect', 'Builder', 'Certifier', 'Contractor', 'Developer', 'Engineer', 'Government',
                             'RDM', 'Strata', 'Supplier', 'Owner/Tenant', 'Others']
            if con_type == -1:
                con_type=12
            contact_type_database = con_type_list[con_type - 1]

            if self.datanow["Asana id"]==None or self.datanow["Asana id"]=='':
                new_asana_id,new_asana_url=create_asana_project(project_type=project_type)
                self.datanow["Asana id"] = new_asana_id
                self.datanow["Asana url"] = new_asana_url

            update_asana_project_tags(self.datanow["Asana id"],project_type)
            update_asana_project(asana_id=self.datanow["Asana id"], name=self.datanow["Current folder"], status=project_state,
                                 services=service_type_list, shop_name=shop_name, apt=asana_area, basement=asana_basement,
                                 notes=asana_note, client_name=client_name,main_contact_name=main_contact_name,
                                 main_contact_type=contact_type_database, total_fee=total_fee, total_paid_fee=total_paid_fee, overdue_fee=overdue_fee)

            if project_state in ['Set Up','Gen Fee Proposal','Email Fee Proposal','Chase Fee Acceptance','Quote Unsuccessful','Pending']:
                update_asana_email(self.datanow["Asana id"], self.ui.text_proinfo_email.toPlainText())
                all_subtasks = get_asana_sub_tasks(self.datanow["Asana id"])
                first_task_subtask_id = list(all_subtasks["Subtasks"].keys())[0]
                update_asana_sub_task(first_task_subtask_id, self.user_email, date.today(), False)
            else:
                content2email=get_asana_email(self.datanow["Asana id"])
                self.ui.text_proinfo_email.setText(content2email)
                if update_sub:
                    all_subtasks = get_asana_sub_tasks(self.datanow["Asana id"])["Subtasks"]
                    first = True
                    for subtask_id in all_subtasks.keys():
                        if first:
                            update_asana_sub_task(subtask_id, self.user_email, date.today(), True)
                            first=False
                        else:
                            update_asana_sub_task(subtask_id, self.user_email, date.today())

                inv_total_list = self.f_get_inv_update_list()
                if inv_total_list!=[]:
                    for i in inv_total_list:
                        if self.datanow["Asana invoice id"][i]==None:

                            asana_invoice_id = create_asana_invoice(self.datanow["Asana id"])
                            self.datanow["Asana invoice id"][i] = asana_invoice_id
                        invoice=self.f_get_invoice_dict(i+1)
                        if invoice["Invoice status"] in ["Sent","Paid"]:
                            update_asana_sub_task(self.datanow["Asana invoice id"][i],self.user_email,date.today()+timedelta(days=7))
                        update_asana_invoice(asana_invoice_id=self.datanow["Asana invoice id"][i], name=invoice["name"],
                                             status=invoice["Invoice status"], payment_date=invoice["Payment Date"],
                                             payment_ingst=invoice["Payment InGST"], net=invoice["Net"])

                bill_total_dict = self.f_get_bill_total_dict()
                if bill_total_dict!={}:
                    for bill_num,bill_info in bill_total_dict.items():
                        bill_id=char2num(bill_num)
                        service_short=bill_info[0]
                        bill_subid=bill_info[1]
                        if self.datanow["Asana bill id"][bill_id]==None:
                            asana_bill_id = create_asana_bill(self.datanow["Asana id"])
                            self.datanow["Asana bill id"][bill_id]=asana_bill_id
                        bill=self.f_get_bill_dict(service_short,bill_subid,bill_id)
                        if bill["Bill status"] in ["Awaiting Approval", "Paid","Awaiting Payment"]:
                            update_asana_sub_task(self.datanow["Asana bill id"][bill_id],self.user_email,date.today()+timedelta(days=7))
                        update_asana_bill(asana_bill_id=self.datanow["Asana bill id"][bill_id], name=bill["name"], status=bill["Bill status"],
                                          contact=bill["From"], bill_in_date=bill["Bill in date"], paid_date=bill["Paid date"], type=bill["Type"],
                                          amount_exgst=float(bill["Amount Excl GST"]),amount_ingst=float(bill["Amount Incl GST"]), headup=bill["HeadsUp"])


            self.f_add_user_history('Update Asana')
            return 'Success'
        except Exception as e:
            print(e)
            traceback.print_exc()
            return 'Fail'
    '''4.2 - Open asana'''
    def f_button_pro_func_openasana(self):
        try:
            if self.ui.text_proinfo_1_quonum.text()=='':
                message('Please login a Project before open asana',self.parent.ui)
                return
            if self.datanow["Asana id"]==None or self.datanow["Asana id"]=='':
                message('You should update Asana first before you open asana',self.parent.ui)
                return
            if self.datanow["Asana url"]==None or self.datanow["Asana url"]=='':
                message('Can not find the asana link, please contact Admin',self.parent.ui)
                return
            link=self.datanow["Asana url"]
            if link[-2:]!="?focus=true":
                link=link+"?focus=true"
            open_link_with_edge(link)
        except:
            traceback.print_exc()
    '''5.1 - Update Xero'''
    def f_button_pro_func_updatexero(self):
        try:
            if self.ui.text_proinfo_1_pronum=='':
                message('Cant Find the Project Number, Please sent first invoice to the client.',self.parent.ui)
                return
            state = self.ui.combobox_pro_1_state.currentText()
            if state in ['Set Up','Gen Fee Proposal','Email Fee Proposal','Chase Fee Acceptance']:
                message("Please upload a fee acceptance before you update xero",self.parent.ui)
                return
            self.f_button_pro_func_refreshxero()
            project_name = self.ui.text_proinfo_1_proname.text()
            project_type = self.f_get_project_type()
            service_list = self.f_get_pro_service_type('Frame_with_var')
            xero_invoice_status_map = {"DRAFT": "Sent",
                                       "SUBMITTED": "Sent",
                                       "AUTHORISED": "Sent",
                                       "PAID": "Paid",
                                       "VOIDED": "Voided",
                                       "Draft": "Sent",
                                       "Submitted": "Sent",
                                       "Authorised": "Sent",
                                       "Paid": "Paid",
                                       "Voided": "Voided",
                                       "Sent":"Sent",
                                       }
            xero_bill_status_map = {
                "DRAFT": "Draft",
                "SUBMITTED": "Awaiting Approval",
                "AUTHORISED": "Awaiting Payment",
                "PAID": "Paid",
                "VOIDED": "Voided",
                "Draft": "Draft",
                "Submitted": "Awaiting Approval",
                "Authorised": "Awaiting Payment",
                "Paid": "Paid",
                "Voided": "Voided",
                "Awaiting Approval":"Awaiting Approval",
                "Awaiting Payment":"Awaiting Payment",
            }
            invoice_state_color_dict = {"Backlog": conf['Text gray style'],"Sent":conf['Text red style'],"Paid":conf['Text green style'],"Voided": conf['Text purple style']}
            invoice_state_color_dict2 = {"Backlog": conf['Text gray_white style'], "Sent": conf['Text red_white style'],
                                        "Paid": conf['Text green_white style'], "Voided": conf['Text purple_white style']}
            bill_state_color_dict = {"Draft": conf['Text gray style'],"Awaiting Approval": conf['Text red style'],"Awaiting Payment":conf['Text yellow style'],"Paid":conf['Text green style'],"Voided": conf['Text purple style']}

            for i in range(8):
                inv_text_name=self.ui_name_dict["text_finan_1_inv" + str(i + 1)]
                finan_text_list=[self.ui_name_dict["text_finan_1_total" + str(i + 1)+'_'+str(j+1)] for j in range(5)]
                inv_num=inv_text_name.text()
                if inv_num!='' and inv_num.find('-')==-1:
                    invoice_xero_id=self.datanow["Xero invoice id"][i]
                    if invoice_xero_id!=None:
                        invoice_info_dict=get_xero_invoice_status(invoice_xero_id)
                        status=invoice_info_dict["status"]
                        last_payment_date=invoice_info_dict["last_payment_date"]
                        payment_amount=invoice_info_dict["payment_amount"]
                        pay_or_not=False
                        if last_payment_date!=None:
                            date_text=self.ui_name_dict["text_finan_1_total" + str(i + 1)+'_4']
                            date_text.setText(str(last_payment_date))
                        if is_float(payment_amount):
                            amount_text=self.ui_name_dict["text_finan_1_total" + str(i + 1)+'_3']
                            amount_text.setText(str(round(float(payment_amount),1)))
                            if float(payment_amount)>0:
                                pay_or_not = True
                        if status in ["PAID", "Paid"] or pay_or_not:
                            contact_id = self.datanow["Client invoice id"][i]
                            contact_xero_id = f_get_client_xero_id(contact_id)
                            update_xero_invoice_contact(invoice_xero_id, contact_xero_id)
                        elif status not in ["VOIDED", "Voided"]:
                            contact_id = self.datanow["Client invoice id"][i]
                            contact_xero_id = f_get_client_xero_id(contact_id)
                            items = []
                            for service_full_name in service_list:
                                service_short_name = conf['Service short dict'][service_full_name]
                                for j in range(4):
                                    radiobutton=self.ui_name_dict["radbutton_finan_1_" + service_short_name + '_' + str(j + 1) + '_' + str(i+1)]
                                    if radiobutton.isChecked():
                                        doc_name=self.ui_name_dict["text_finan_1_" + service_short_name + "_doc" + str(j + 1)].text()
                                        fee = round(float(self.ui_name_dict["text_finan_1_" + service_short_name + "_fee" + str(j + 1) + '_1'].text()), 1)
                                        item_i = {"Item": doc_name, "Fee": fee}
                                        items.append(item_i)
                            update_xero_invoice(invoice_xero_id, inv_num, items, project_type, project_name,contact_xero_id)
                        inv_status=xero_invoice_status_map[status]
                        style = invoice_state_color_dict[inv_status]
                        style2 = invoice_state_color_dict2[inv_status]
                        inv_text_name.setStyleSheet(style)
                        for finan_text in finan_text_list:
                            finan_text.setStyleSheet(style2)
            for service in service_list:
                service_short = conf['Service short dict'][service]
                for j in range(3):
                    bill_letter=self.ui_name_dict["text_finan_2_" + service_short + '_' + str(j + 1) + '_1'].text()
                    if bill_letter != '':
                        bill_num = char2num(bill_letter)
                        bill_id=self.datanow["Xero bill"][bill_num]
                        if bill_id!=None:
                            bill_amount_text=self.ui_name_dict["text_finan_2_" + service_short + '_' + str(j + 1) + '_4'].text()
                            amount = round(float(bill_amount_text), 1)
                            combobox=self.ui_name_dict["combobox_finan_2_" + service_short + '_' + str(j+1)]
                            bill_type = combobox.currentText()
                            checkbox=self.ui_name_dict["checkbox_finan_2_" + service_short + '_' + str(j+1)]

                            if checkbox.isChecked():
                                no_gst = True
                            else:
                                no_gst = False
                            contact_id = self.datanow["Client bill id"][bill_num]
                            contact_xero_id = f_get_client_xero_id(contact_id)
                            bill_info=get_xero_bill_status(bill_id)
                            status=bill_info["status"]
                            paid_date=bill_info["paid_data"]
                            if status not in ["PAID","VOIDED","Paid","Voided"]:
                                update_xero_bill(bill_id, service, amount, bill_type, no_gst, contact_xero_id)
                            bill_status=xero_bill_status_map[status]
                            style=bill_state_color_dict[bill_status]
                            bill_gst_text=self.ui_name_dict["text_finan_2_" + service_short + '_' + str(j + 1) + '_5']
                            bill_gst_text.setStyleSheet(style)
                            self.datanow["Bill paid date"][bill_num]=paid_date
            update_asana_result=self.f_button_pro_func_updateasana(False)
            quo_num=self.ui.text_proinfo_1_quonum.text()
            pro_num=self.ui.text_proinfo_1_pronum.text()
            pro_name=self.ui.text_proinfo_1_proname.text()
            pro_info=quo_num+'-'+pro_num+'-'+pro_name
            if update_asana_result == 'Success':
                message(f'Project {pro_info} \nAsana and Xero Updated Successfully', self.parent.ui)
            else:
                message(f'Project {pro_info} \nCan not update Xero or Asana, please contact admin.', self.parent.ui)
        except:
            message('Can not update Xero or Asana, please contact admin',self.parent.ui)
            traceback.print_exc()
    '''5.2 - Refresh Xero'''
    def f_button_pro_func_refreshxero(self):
        try:
            refresh_token()
        except:
            traceback.print_exc()
    '''5.3 - Login xero'''
    def f_button_pro_func_loginxero(self):
        try:
            login_xero2()
        except:
            traceback.print_exc()
    '''5.4 - Sync Xero'''
    def f_button_pro_func_syncxero(self):
        try:
            exe=r'B:\02.Copilot\Daily_update_script.exe'
            subprocess.Popen([exe], creationflags=subprocess.CREATE_NEW_CONSOLE)
        except:
            traceback.print_exc()



    '''=============================Project info page==============================='''
    '''1 - Project information text to split'''
    def f_text_proinfo_split(self, text_name):
        try:
            text=text_name.text()
            if text.find('\n') == -1:
                return
            modified_text = ' '.join(text.splitlines())
            text_name.setText(modified_text)
        except:
            traceback.print_exc()
    '''1 - Set color and content when state changed'''
    def f_pro_state_change(self):
        try:
            state=self.ui.combobox_pro_1_state.currentText()
            if state=='Quote Unsuccessful':
                self.f_add_user_history('Quote Unsuccessful')
                self.f_quote_unseccussful()
            if self.last_state=='Quote Unsuccessful' and state!='Quote Unsuccessful':
                self.f_add_user_history('Restore Project')
                self.f_project_restore()
            if state not in ['Set Up','Gen Fee Proposal','Email Fee Proposal']:
                if self.ui.button_fee_lock.text()=='Lock':
                    self.f_button_lock(lock_name = "Fee")
                if self.ui.button_proinfo_search_unlock1.text() == 'Lock':
                    self.f_button_proinfo_search_unlock(contact_type='client')
                if self.ui.button_proinfo_search_unlock2.text() == 'Lock':
                    self.f_button_proinfo_search_unlock(contact_type='contact')

            self.last_state=state
            self.f_set_four_buttons()
        except:
            traceback.print_exc()

    def f_quote_unseccussful(self):
        try:
            quo_num=self.ui.text_proinfo_1_quonum
            if quo_num=='':
                message('You need to load a project first',self.parent.ui)
                return
            quote=messagebox('Quote',"Are you sure you want to put this project into Quote Unsuccessful", self.parent.ui)
            if quote:
                if self.datanow["Asana id"]!=None and self.datanow["Asana id"]!='':
                    update_asana_result=self.f_button_pro_func_updateasana(False)
                    if update_asana_result == 'Success':
                        message('Quote Unsuccessful and Asana Updated', self.parent.ui)
                    else:
                        message('Quote Unsuccessful but Asana updated failed, please contact admin', self.parent.ui)
                else:
                    message("Quote Unsuccessful",self.parent.ui)
        except:
            traceback.print_exc()

    def f_project_restore(self):
        try:
            quo_num=self.ui.text_proinfo_1_quonum
            if quo_num=='':
                message('You need to load a project first',self.parent.ui)
                return
            restore=messagebox('Quote',"Are you sure you want to restore this project", self.parent.ui)
            if restore:
                if self.datanow["Asana id"]!=None and self.datanow["Asana id"]!='':
                    update_asana_result=self.f_button_pro_func_updateasana(False)
                    if update_asana_result == 'Success':
                        message('Project Restore and Asana Updated', self.parent.ui)
                    else:
                        message('Project Restore but Asana updated failed, please contact admin', self.parent.ui)
                else:
                    message("Project Restore",self.parent.ui)
        except:
            traceback.print_exc()
    '''1 - check similar project name'''
    def f_button_pro_checksimilar(self):
        try:
            project_similar_list=[]
            project_name_now=self.ui.text_proinfo_1_proname.text()
            quo_num_now=self.ui.text_proinfo_1_quonum.text()
        except:
            traceback.print_exc()
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * from bridge.projects")
            search_results = mycursor.fetchall()
        except:
            traceback.print_exc()
            return
        finally:
            mycursor.close()
            mydb.close()
        try:
            for projects in search_results:
                project_name=str(projects[2])
                edit_distance = Levenshtein.distance(project_name_now, project_name)
                max_len = max(len(project_name_now), len(project_name))
                similarity = (1 - edit_distance / max_len)
                if similarity>0.7:
                    quo_num=projects[0]
                    if quo_num_now!=quo_num:
                        pro_num=str(projects[1])
                        similar_pro=quo_num+'-'+pro_num+'-'+project_name
                        project_similar_list.append(similar_pro)
            if project_similar_list==[]:
                message('No similar project found',self.parent.ui)
            else:
                similar_project_str=''
                for project_info in project_similar_list:
                    similar_project_str+=project_info+'\n'
                message('Similar projects found:\n'+similar_project_str,self.parent.ui)
        except:
            traceback.print_exc()


    '''1 - Project number change'''
    def f_text_proinfo_1_pronum_change(self):
        try:
            if self.ui.text_proinfo_1_pronum.text()!='' and self.ui.button_finan_1_lock.text()=='Lock':
                self.f_button_lock(lock_name = "Financial")
        except:
            traceback.print_exc()
    '''1 - Change proposal type'''
    def f_proinfo_1_proposaltype_change(self):
        try:
            quo_num=self.ui.text_proinfo_1_quonum.text()
            if self.radiobuttongroup_proinfo_1_proposaltype.checkedId()==1:
                self.ui.frame_proinfo_minorfeature.setVisible(True)
                self.ui.frame_proinfo_majorfeature.setVisible(False)
                self.ui.frame_fee_stage.setVisible(False)
                self.ui.frame_fee_timeframe.setVisible(True)
                self.f_set_scope_of_work_tables(quo_num, 'Minor')
            else:
                self.ui.frame_proinfo_minorfeature.setVisible(False)
                self.ui.frame_proinfo_majorfeature.setVisible(True)
                self.ui.frame_fee_stage.setVisible(True)
                self.ui.frame_fee_timeframe.setVisible(False)
                self.f_set_scope_of_work_tables(quo_num, 'Major')
        except:
            traceback.print_exc()
    '''1 - Servicettype group change'''
    def f_checkboxgroup_proinfo_1_servicettype(self,button):
        try:
            button_id = self.checkboxgroup_proinfo_1_servicettype.id(button)
            service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'mech', 'mechrev', 'mis','install']
            frame_scope_list=[self.ui_name_dict["frame_fee_scope_" + service_list[i]] for i in range(len(service_list))]
            frame_feedetail_list=[self.ui_name_dict["frame_fee_pro_" + service_list[i]] for i in range(len(service_list))]
            frame_invoice_list =[self.ui_name_dict["frame_finan_1_" + service_list[i]] for i in range(len(service_list))]
            frame_bill_list =[self.ui_name_dict["frame_finan_2_" + service_list[i]] for i in range(len(service_list))]
            frame_profit_list =[self.ui_name_dict["frame_finan_3_" + service_list[i]] for i in range(len(service_list))]
            if button_id==1 or button_id==6:
                state=self.checkboxgroup_proinfo_1_servicettype.button(1).isChecked() or self.checkboxgroup_proinfo_1_servicettype.button(6).isChecked()
            else:
                state=self.checkboxgroup_proinfo_1_servicettype.button(button_id).isChecked()
            frame_scope_list[button_id-1].setVisible(state)
            frame_feedetail_list[button_id - 1].setVisible(state)
            frame_invoice_list[button_id - 1].setVisible(state)
            frame_bill_list[button_id - 1].setVisible(state)
            frame_profit_list[button_id - 1].setVisible(state)
        except:
            traceback.print_exc()
    '''2 - Client/Main contact search'''
    def f_client_search(self,client_type):
        try:
            client_info_list=[]
            client_info_dict={"Client":(self.ui.text_proinfo_2_clientname,self.ui.table_client_search),
                              "Main Contact":(self.ui.text_proinfo_2_contactname,self.ui.table_contact_search)}
            search_text=client_info_dict[client_type][0]
            search_table=client_info_dict[client_type][1]
            search_content=search_text.text()
            if search_content=='':
                search_table.hide()
                return
            if search_content=='-':
                search_table.setRowCount(0)
                search_table.show()
                table_row_now = search_table.rowCount()
                search_table.insertRow(table_row_now)
                search_table.setItem(table_row_now, 0, QTableWidgetItem('Clear'))
                search_table.setItem(table_row_now, 1, QTableWidgetItem('----'))
                search_table.setItem(table_row_now, 2, QTableWidgetItem('----'))
                search_table.setItem(table_row_now, 3, QTableWidgetItem('----'))
                self.f_set_table_style(search_table)
                return
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * from bridge.clients WHERE LOWER(full_name) LIKE LOWER(%s) OR LOWER(company) LIKE LOWER(%s) OR LOWER(contact_email) LIKE LOWER(%s)",
                                 (f'%{search_content}%',f'%{search_content}%',f'%{search_content}%',))
                search_results = mycursor.fetchall()
            except:
                traceback.print_exc()
                return
            finally:
                mycursor.close()
                mydb.close()
            for row in search_results:
                client_info_list_i=[row[0],row[2],row[4],row[8]]
                client_info_list.append(client_info_list_i)
            search_table.setRowCount(0)
            search_table.show()
            for client_info in client_info_list:
                table_row_now = search_table.rowCount()
                search_table.insertRow(table_row_now)
                search_table.setItem(table_row_now, 0, QTableWidgetItem(str(client_info[0])))
                search_table.setItem(table_row_now, 1, QTableWidgetItem(str(client_info[1])))
                search_table.setItem(table_row_now, 2, QTableWidgetItem(str(client_info[2])))
                search_table.setItem(table_row_now, 3, QTableWidgetItem(str(client_info[3])))
            self.f_set_table_style(search_table)

        except:
            traceback.print_exc()
    '''2 - Client/Main contact search table clicked'''
    def f_search_table_click(self, contact_type):
        try:
            table_info_dict={"Client":(self.ui.table_client_search,'client',),
                             "Main Contact":(self.ui.table_contact_search,'contact',),}
            table=table_info_dict[contact_type][0]
            text_item_name=table_info_dict[contact_type][1]
            current_row = table.currentRow()
            if current_row >= 0:
                client_id=table.item(current_row, 0).text()
                if client_id=='Clear':
                    if text_item_name == 'client':
                        self.datanow["Client id"] = None
                        self.datanow["Xero client id"] = None
                    else:
                        self.datanow["Main contact id"] = None
                        self.datanow["Xero main contact id"] = None
                    ui_text_list = ['name', 'company', 'adress', 'abn', 'phone', 'email']
                    for ui_text in ui_text_list:
                        content_ui = self.ui_name_dict["text_proinfo_2_" + text_item_name + ui_text]
                        content_ui.setText('')
                    table.hide()
                    return
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                    client_database = mycursor.fetchall()
                except:
                    traceback.print_exc()
                    return
                finally:
                    mycursor.close()
                    mydb.close()
                if len(client_database) == 1:
                    client_database = client_database[0]
                    xero_client_id = client_database[1]
                    if text_item_name=='client':
                        self.datanow["Client id"]=client_id
                        self.datanow["Xero client id"] = xero_client_id
                    else:
                        self.datanow["Main contact id"] = client_id
                        self.datanow["Xero main contact id"] = xero_client_id
                    ui_text_list = ['name', '', 'company', 'adress', 'abn', 'phone', 'email']
                    for j in range(7):
                        if j == 1:
                            contact_type = client_database[3]
                            contact_list = ['Architect', 'Builder', 'Certifier', 'Contractor', 'Developer', 'Engineer',
                                            'Government', 'RDM', 'Strata', 'Supplier', 'Owner/Tenant', 'Others']
                            if contact_type == 'None':
                                contact_type_id = 11
                            else:
                                contact_type_id = contact_list.index(contact_type)
                            if text_item_name=='client':
                                self.radbuttongroup_proinfo_2_clientcontacttype.button(contact_type_id + 1).setChecked(True)
                            else:
                                self.radbuttongroup_proinfo_2_contactcontacttype.button(contact_type_id + 1).setChecked(True)
                        else:
                            content = client_database[j + 2]
                            text_type = ui_text_list[j]
                            content_ui=self.ui_name_dict["text_proinfo_2_" + text_item_name + text_type]
                            content_ui.setText(content)
                    table.hide()
        except:
            traceback.print_exc()
    '''2 - Client/Main contact content unlock'''
    def f_button_proinfo_search_unlock(self,contact_type):
        try:
            ui_text_list = ['name', 'company', 'adress', 'abn', 'phone', 'email']
            text_list=[self.ui_name_dict["text_proinfo_2_" + contact_type + ui_name] for ui_name in ui_text_list]
            radbuttongroup_list=[self.ui_name_dict["radbutton_proinfo_2_"+contact_type+"contacttype_" + str(i + 1)] for i in range(12)]
            if contact_type=='client':
                button=self.ui.button_proinfo_search_unlock1
            else:
                button=self.ui.button_proinfo_search_unlock2
            if button.text()=='Unlock':
                set_state=True
                set_style='font: 12pt "Calibri";color: rgb(0, 0, 0);border-color: rgb(255, 255, 255);background-color: rgb(255, 255, 255);'
                new_name='Lock'
            else:
                set_state=False
                set_style ='font: 12pt "Calibri";color: rgb(0, 0, 0);border-color: rgb(255, 255, 255);background-color: rgb(200, 200, 200);'
                new_name = 'Unlock'
            for text_name in text_list:
                if set_state==True:
                    text_name.setReadOnly(False)
                else:
                    text_name.setReadOnly(True)
            for radiobutton in radbuttongroup_list:
                radiobutton.setEnabled(set_state)
            text_name = self.ui_name_dict["text_proinfo_2_" + contact_type + 'name']
            text_name.setStyleSheet(set_style)
            button.setText(new_name)
        except:
            traceback.print_exc()
    '''2 - Client/Main contact update'''
    def f_button_proinfo_search_update(self,contact_type):
        update_xero_state = False
        try:
            full_name=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'name'].text())
            company=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'company'].text())
            address=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'adress'].text())
            ABN=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'abn'].text())
            contact_number=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'phone'].text())
            contact_email=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'email'].text())
            if contact_email==None:
                contact_email='0@com.au'
                self.ui_name_dict["text_proinfo_2_" + contact_type + 'email'].setText(contact_email)
            if contact_type=='client':
                con_type=self.radbuttongroup_proinfo_2_clientcontacttype.checkedId()
                client_id = self.datanow["Client id"]
            else:
                con_type = self.radbuttongroup_proinfo_2_contactcontacttype.checkedId()
                client_id = self.datanow["Main contact id"]
            if con_type == -1:
                message("Please choose contact type",self.parent.ui)
                return
            con_type_list = ['Architect', 'Builder', 'Certifier', 'Contractor', 'Developer', 'Engineer','Government', 'RDM', 'Strata', 'Supplier', 'Owner/Tenant', 'Others']
            contact_type_database = con_type_list[con_type - 1]
            try:
                project_client_num=0
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",
                                               database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.projects WHERE client_id = %s or main_contact_id=%s", (client_id,client_id,))
                project_client_num = len(mycursor.fetchall())
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
            update=messagebox('Confirm',f'Updating client information in the database will affect approximately {project_client_num} projects. Do you want to proceed?', self.parent.ui)
            if update:
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                    client_database = mycursor.fetchall()
                    mycursor.execute("UPDATE bridge.clients SET full_name = %s, company = %s, address = %s, ABN = %s, contact_number = %s, contact_email = %s, contact_type=%s"
                                     " WHERE client_id=%s",(full_name, company, address, ABN, contact_number, contact_email, contact_type_database, client_id,))
                    mydb.commit()

                    if len(client_database) == 1:
                        xero_client_id = client_database[0][1]
                        contact_name=None
                        if company==None:
                            if full_name==None:
                                message('Company or name can not be None',self.parent.ui)
                            else:
                                contact_name=full_name
                        else:
                            if full_name == None:
                                contact_name=company
                            else:
                                contact_name=company+'-'+full_name
                        if contact_name!=None:
                            # todo 2
                            update_xero_contact(contact_id=xero_client_id, contact_name=contact_name, account_number=client_id, first_name=full_name, last_name=None, email=contact_email,
                                                phone_number=contact_number, abn=ABN, address=address)
                            update_xero_state=True
                except:
                    traceback.print_exc()
                finally:
                    mycursor.close()
                    mydb.close()
                if update_xero_state:
                    message('Update client in database and Xero successfully', self.parent.ui)
                else:
                    message('Update client unsuccessfully, please contact admin', self.parent.ui)


        except:
            traceback.print_exc()
    '''2 - Client/Main contact create'''
    def f_button_proinfo_search_create(self,contact_type):
        client_id=None
        xero_client_id = None
        error_reason=''
        try:
            full_name=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'name'].text().strip())
            company=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'company'].text().strip())
            address=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'adress'].text().strip())
            ABN=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'abn'].text().strip())
            contact_number=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'phone'].text().strip())
            contact_email=self.f_change2none(self.ui_name_dict["text_proinfo_2_" + contact_type + 'email'].text().strip())
            if contact_email==None:
                contact_email='0@com.au'
                self.ui_name_dict["text_proinfo_2_" + contact_type + 'email'].setText(contact_email)
            else:
                if contact_email.find(';')!=-1:
                    contact_email=contact_email.split(';')[0].strip()
            if not is_valid_email(contact_email):
                message('Invalid email address',self.parent.ui)
                return
            if self.contains_non_unicode_characters([full_name,company,address,ABN,contact_number,contact_email]):
                message('Information contains incorrect content',self.parent.ui)
                return
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.clients where full_name=%s and company=%s", (full_name,company,))
                clients_similar_database=mycursor.fetchall()
                if len(clients_similar_database) > 0:
                    message('Similar client found, check it again',self.parent.ui)
                    return
                mycursor.execute("SELECT MAX(client_id) AS max_value FROM bridge.clients")
                current_client_id=int(mycursor.fetchall()[0][0])
                client_id_num=str(current_client_id+1)
                client_id = client_id_num.zfill(8)
                if contact_type=='client':
                    con_type=self.radbuttongroup_proinfo_2_clientcontacttype.checkedId()
                else:
                    con_type = self.radbuttongroup_proinfo_2_contactcontacttype.checkedId()
                if con_type == -1:
                    message("Please choose contact type",self.parent.ui)
                    return
                con_type_list = ['Architect', 'Builder', 'Certifier', 'Contractor', 'Developer', 'Engineer','Government', 'RDM', 'Strata', 'Supplier', 'Owner/Tenant', 'Others']
                contact_type_database = con_type_list[con_type - 1]
                if company == None:
                    if full_name == None:
                        message('Company and name can not both be None', self.parent.ui)
                        return
                    else:
                        contact_name = full_name
                else:
                    if full_name == None:
                        contact_name = company
                    else:
                        contact_name = company + '-' + full_name

                xero_client_id,error_reason=create_xero_contact(contact_name=contact_name, account_number=client_id, full_name=full_name, email=contact_email, phone_number=contact_number, abn=ABN, address=address)
                if xero_client_id!=None:
                    mycursor.execute("INSERT INTO bridge.clients VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)",
                                     (client_id,xero_client_id,full_name,contact_type_database,company,address,ABN,contact_number,contact_email,))
                    mydb.commit()
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
            if xero_client_id!=None:
                if contact_type=='client':
                    self.datanow["Client id"]=client_id
                    self.datanow["Xero client id"]=xero_client_id
                else:
                    self.datanow["Main contact id"]=client_id
                    self.datanow["Xero main contact id"]=xero_client_id
                message('Create client in database and Xero successfully',self.parent.ui)
            else:
                real_error_reason=error_reason.split('HTTP')[0]
                message(f'Create client unsuccessfully.\n{real_error_reason}',self.parent.ui)
            self.ui.table_client_search.hide()
            self.ui.table_contact_search.hide()
        except:
            traceback.print_exc()

    '''=============================Fee page========================================'''
    '''0 - Page lock'''
    def f_button_lock(self,lock_name):
        try:
            service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis','var']
            frame_scope_list=[self.ui_name_dict["frame_fee_scope_" + service_list[i]] for i in range(len(service_list)-1)]
            frame_feedetail_list=[self.ui_name_dict["frame_fee_pro_" + service_list[i]] for i in range(len(service_list))]
            frame_invoice_list =[self.ui_name_dict["frame_finan_1_" + service_list[i]] for i in range(len(service_list))]
            frame_bill_list =[self.ui_name_dict["frame_finan_2_" + service_list[i]] for i in range(len(service_list))]
            frame_profit_list =[self.ui_name_dict["frame_finan_3_" + service_list[i]] for i in range(len(service_list))]
            frame_stage_list=[self.ui_name_dict["table_fee_stage_" + str(i+1)] for i in range(4)]
            fee_lock_text_list=[self.ui.text_fee_1_date,self.ui.text_fee_2_period1,
                                self.ui.text_fee_2_period2,self.ui.text_fee_2_period3,
                                self.ui.text_fee_3_install_date,self.ui.text_fee_3_install_rev,self.ui.text_fee_3_install_program]
            fee_lock_frame_list=[self.ui.frame_fee_reference,self.ui.frame_fee_timeframe,self.ui.frame_fee_1_total]+frame_scope_list+frame_feedetail_list+frame_stage_list
            finan_lock_frame_list = [self.ui.frame_finan_1_total]
            finan_radiobutton_list=[self.ui_name_dict["radbutton_finan_1_" + service_i + '_' + str(i + 1) + '_' + str(j+1)] for service_i in service_list for i in range(4) for j in range(8)]
            lock_dict={"Fee":(self.ui.button_fee_lock,fee_lock_frame_list),
                       "Financial":(self.ui.button_finan_1_lock,finan_lock_frame_list)}
            button=lock_dict[lock_name][0]
            frames=lock_dict[lock_name][1]
            if button.text()=='Lock':
                for frame in frames:
                    frame.setEnabled(False)
                if lock_name=='Fee':
                    for ui_text in fee_lock_text_list:
                        ui_text.setStyleSheet('font: 12pt "Calibri";border-color: rgb(85, 255, 255);background-color: rgb(200, 200, 200);color: rgb(0, 0, 0);')
                else:
                    for radiobutton in finan_radiobutton_list:
                        radiobutton.setEnabled(False)
                button.setText('Unlock')
            else:
                for frame in frames:
                    frame.setEnabled(True)
                if lock_name=='Fee':
                    for ui_text in fee_lock_text_list:
                        ui_text.setStyleSheet('font: 12pt "Calibri";color: rgb(0, 0, 0);border-color: rgb(255, 255, 255);background-color: rgb(255, 255, 255);')
                else:
                    for radiobutton in finan_radiobutton_list:
                        radiobutton.setEnabled(True)
                button.setText('Lock')

        except:
            traceback.print_exc()
    '''1 - Today'''
    def f_gettodaydate(self,text_name):
        try:
            today_dict={"Service":self.ui.text_fee_1_date,
                        "Install":self.ui.text_fee_3_install_date}
            text=today_dict[text_name]
            date_today=date.today().strftime("%d-%b-%Y")
            text.setText(date_today)
        except:
            traceback.print_exc()
    '''1 - Fee revision added'''
    def f_toolbutton_fee_revision_add(self):
        try:
            rev_text=self.ui.text_fee_1_revision.text()
            if is_float(rev_text):
                rev_new=int(rev_text)+1
                self.ui.text_fee_1_revision.setText(str(rev_new))
            else:
                self.ui.text_fee_1_revision.setText('1')
            color_tuple = ("green", "yellow", "red", "red")
            self.f_set_pro_buttons_color(color_tuple)
        except:
            traceback.print_exc()
    '''1 - Set four buttons color and content when fee revision added'''
    def f_set_four_buttons(self):
        try:
            state=self.ui.combobox_pro_1_state.currentText()
            if state=='Set Up':
                color_tuple=("yellow","red","red","red",)
                self.f_set_pro_buttons_color(color_tuple)
            elif state not in ['Gen Fee Proposal','Email Fee Proposal','Chase Fee Acceptance']:
                color_tuple = ("green","green","green","green",)
                self.f_set_pro_buttons_color(color_tuple)
            rev=self.ui.text_fee_1_revision.text()
            self.ui.button_pro_state_genfee.setText('Gen fee proposal v'+rev)
            self.ui.button_pro_state_emailfee.setText('Email fee proposal v'+rev)
            self.ui.button_pro_state_chasefee.setText('Chase fee proposal v'+rev)

        except:
            traceback.print_exc()
    '''3 - Stage content save as default'''
    def f_save_stage2default(self, stage):
        try:
            table=self.ui_name_dict["table_fee_stage_" +str(stage+1)]
            content_list=[]
            for row in range(table.rowCount()):
                content = self.f_tableitem2none(table.item(row, 1))
                if content!=None:
                    content_list.append(content)
        except:
            traceback.print_exc()
            return
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("DELETE FROM bridge.default_stage WHERE stage=%s",(stage,))
            for i in range(len(content_list)):
                mycursor.execute("INSERT INTO bridge.default_stage VALUES (%s, %s, %s)",
                         (stage, content_list[i], i))
            mydb.commit()
            message('Stage table set as default',self.parent.ui)
        except:
            traceback.print_exc()
            return
        finally:
            mycursor.close()
            mydb.close()
    '''3 - Stage checkbox checked'''
    def f_stage_total_check(self, index):
        try:
            checkbox_total=self.ui_name_dict["checkbox_fee_stage_" + str(index)]
            table=self.ui_name_dict["table_fee_stage_" + str(index)]
            if checkbox_total.isChecked():
                for row in range(table.rowCount()):
                    checkbox = table.item(row, 0)
                    content=table.item(row, 1).text()
                    if content!='':
                        checkbox.setCheckState(Qt.Checked)
            else:
                for row in range(table.rowCount()):
                    checkbox = table.item(row, 0)
                    checkbox.setCheckState(Qt.Unchecked)
        except:
            traceback.print_exc()
    '''4 - Set scope tables'''
    def f_set_scope_of_work_tables(self,quo_num,proposal_type):
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM bridge.scope WHERE quotation_number = %s and minor_or_major = %s", (quo_num,proposal_type,))
            scope_of_work_database=mycursor.fetchall()
        except:
            traceback.print_exc()
            return
        finally:
            mycursor.close()
            mydb.close()
        try:
            table_max_row_dict={"Mechanical Service":[0,0,0],
                                 "Mechanical Review":[0,0,0],
                                 "CFD Service": [0,0,0],
                                 "Electrical Service": [0,0,0],
                                 "Hydraulic Service": [0,0,0],
                                 "Fire Service": [0,0,0],
                                 "Miscellaneous": [0,0,0],
                                 "Installation": [0,0,0],}
            scope_dict={'Extent':0,'Clarifications':1,'Deliverables':2}
            for item in scope_of_work_database:
                service_name = item[2]
                scope_type = item[3]
                row_index = item[6]
                table_max_row_dict[service_name][scope_dict[scope_type]] = row_index + 1
            for service in table_max_row_dict:
                for j in range(3):
                    table=self.ui_name_dict["table_fee_3_" + conf['Service short dict'][service] + '_' + str(j + 1)]
                    row_max=table_max_row_dict[service][j]
                    row_max=max(row_max,1)
                    table.setRowCount(row_max)
                    for row in range(table.rowCount()):
                        table.setRowHeight(row, 20)
                        checkbox_item = QTableWidgetItem()
                        checkbox_item.setFlags(checkbox_item.flags() | Qt.ItemIsUserCheckable)
                        checkbox_item.setCheckState(Qt.Unchecked)
                        table.setItem(row, 0, checkbox_item)
                    table.setFixedHeight(int(20*row_max*1.2+25))

            table_name_list=['Extent','Clarifications','Deliverables']
            for item in scope_of_work_database:
                service_name = item[2]
                scope_type = item[3]
                include=item[4]
                content = item[5]
                row_index = item[6]
                table=self.ui_name_dict["table_fee_3_" + conf['Service short dict'][service_name]+'_'+str(table_name_list.index(scope_type)+1)]
                table.setItem(row_index, 1, QTableWidgetItem(content))
                checkbox=table.item(row_index, 0)
                if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                    if include==1:
                        checkbox.setCheckState(Qt.Checked)
                    else:
                        checkbox.setCheckState(Qt.Unchecked)
            service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis']
            for service in service_list:
                for i in range(3):
                    table=self.ui_name_dict["table_fee_3_" + service+'_'+str(i + 1)]
                    self.f_set_table_style(table)
        except:
            traceback.print_exc()
    '''4 - Scope of work table up'''
    def f_fee_scope_table_up(self,service,id):
        try:
            table=self.ui_name_dict["table_fee_3_" + service + "_"+id]
            selected_items = table.selectedItems()
            if selected_items:
                row = selected_items[0].row()
                if row!=0:
                    item_current = self.f_get_table_item(table, row, 1)
                    item_above = self.f_get_table_item(table, row-1, 1)
                    table.setItem(row, 1, QTableWidgetItem(item_above))
                    table.setItem(row - 1, 1, QTableWidgetItem(item_current))
                    checkbox_current=table.item(row, 0)
                    checkbox_current_state=Qt.Unchecked
                    if checkbox_current is not None and checkbox_current.flags() & Qt.ItemIsUserCheckable:
                        checkbox_current_state=checkbox_current.checkState()
                    checkbox_above=table.item(row-1, 0)
                    checkbox_above_state=Qt.Unchecked
                    if checkbox_above is not None and checkbox_above.flags() & Qt.ItemIsUserCheckable:
                        checkbox_above_state=checkbox_above.checkState()
                        checkbox_above.setCheckState(checkbox_current_state)
                    if checkbox_current is not None and checkbox_current.flags() & Qt.ItemIsUserCheckable:
                        checkbox_current.setCheckState(checkbox_above_state)
                    table.setCurrentCell(row - 1, 1)
                    self.f_set_table_style(table)

        except:
            traceback.print_exc()
    '''4 - Scope of work table down'''
    def f_fee_scope_table_down(self,service,id):
        try:
            table=self.ui_name_dict["table_fee_3_" + service + "_"+id]
            selected_items = table.selectedItems()
            if selected_items:
                row = selected_items[0].row()
                if int(row) != int(table.rowCount()) - 1:
                    item_current = self.f_get_table_item(table, row, 1)
                    item_down=self.f_get_table_item(table, row + 1, 1)
                    table.setItem(row, 1, QTableWidgetItem(item_down))
                    table.setItem(row + 1, 1, QTableWidgetItem(item_current))
                    checkbox_current=table.item(row, 0)
                    checkbox_current_state=Qt.Unchecked
                    if checkbox_current is not None and checkbox_current.flags() & Qt.ItemIsUserCheckable:
                        checkbox_current_state=checkbox_current.checkState()
                    checkbox_down=table.item(row+1, 0)
                    checkbox_down_state=Qt.Unchecked
                    if checkbox_down is not None and checkbox_down.flags() & Qt.ItemIsUserCheckable:
                        checkbox_down_state=checkbox_down.checkState()
                        checkbox_down.setCheckState(checkbox_current_state)
                    if checkbox_current is not None and checkbox_current.flags() & Qt.ItemIsUserCheckable:
                        checkbox_current.setCheckState(checkbox_down_state)
                    table.setCurrentCell(row + 1, 1)
                    self.f_set_table_style(table)
        except:
            traceback.print_exc()
    '''4 - Scope of work table add'''
    def f_fee_scope_table_add(self,service,id):
        try:
            table=self.ui_name_dict["table_fee_3_" +service+ '_'+id]
            current_row = table.currentRow()
            if current_row >= 0:
                table_row_now=table.rowCount()
                table.insertRow(table_row_now)
                table.setRowHeight(table_row_now, 20)
                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(checkbox_item.flags() | Qt.ItemIsUserCheckable)
                checkbox_item.setCheckState(Qt.Unchecked)
                table.setItem(table_row_now, 0, checkbox_item)
                for i in range(table_row_now-current_row-1):
                    row=table_row_now-i-1
                    item_current = self.f_get_table_item(table,row, 1)
                    table.setItem(row+1, 1, QTableWidgetItem(item_current))
                    checkbox_current = table.item(row, 0)
                    checkbox_down = table.item(row+1, 0)
                    if checkbox_current is not None and checkbox_current.flags() & Qt.ItemIsUserCheckable:
                        if checkbox_current.checkState() == Qt.Checked:
                            checkbox_down.setCheckState(Qt.Checked)
                        else:
                            checkbox_down.setCheckState(Qt.Unchecked)
                        if row == current_row+1:
                            checkbox_current.setCheckState(Qt.Unchecked)
                table.setItem(current_row + 1, 1, QTableWidgetItem())
                table.setFixedHeight(int(20 * (table_row_now+1) * 1.2 + 25))
                self.f_set_table_style(table)
        except:
            traceback.print_exc()
    '''4 - Installation proposal'''
    def f_button_fee_3_install_preview(self):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            check_or_create_folder(accounting_dir)
            pro_name = self.ui.text_proinfo_1_proname.text()
            fee_installation_rev=self.ui.text_fee_3_install_rev.text()
            if not is_float(fee_installation_rev):
                message("Please choose revision",self.parent.ui)
                return
            pdf_name = f"Mechanical Installation Fee Proposal for {pro_name} Rev {fee_installation_rev}.pdf"
            pdf_dir=os.path.join(accounting_dir, pdf_name)
            excel_name = f'Mechanical Installation {pro_name} Back Up.xlsx'
            excel_dir=os.path.join(accounting_dir, excel_name)
            project_type = self.f_get_project_type()
            contact_type = self.get_contact_type()
            if contact_type == 'Client':
                full_name = self.ui.text_proinfo_2_clientname.text()
                company = self.ui.text_proinfo_2_clientcompany.text()
                address = self.ui.text_proinfo_2_clientadress.text()
            else:
                full_name = self.ui.text_proinfo_2_contactname.text()
                company = self.ui.text_proinfo_2_contactcompany.text()
                address = self.ui.text_proinfo_2_contactadress.text()
            quo_num = self.ui.text_proinfo_1_quonum.text()
            install_fee_date = self.ui.text_fee_3_install_date.text()
            drawing_content = []
            table = self.ui.table_proinfo_4_drawing
            for i in range(table.rowCount()):
                try:
                    drawing_number = table.item(i, 0).text()
                    drawing_name = table.item(i, 1).text()
                    revision = table.item(i, 2).text()
                    drawing_i = {"Drawing number": drawing_number,
                                 "Drawing name": drawing_name,
                                 "Revision": revision, }
                    drawing_content.append(drawing_i)
                except:
                    pass
            content_type_list = ['Extent', 'Clarifications', 'Deliverables']
            content_dict={}
            for i in range(3):
                content_type=content_type_list[i]
                content_list_i=[]
                table=self.ui_name_dict["table_fee_3_install_" + str(i + 1)]
                for row in range(table.rowCount()):
                    checkbox = table.item(row, 0)
                    if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                        if checkbox.checkState() == Qt.Checked:
                            content = table.item(row, 1).text()
                            content_list_i.append(content)
                content_dict[content_type]=content_list_i
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.past_project WHERE project_type = %s", (project_type,))
                project_history_database = mycursor.fetchall()
            except:
                traceback.print_exc()
            finally:
                mycursor.close()
                mydb.close()
            past_projects = []
            if len(project_history_database) > 0:
                for i in range(len(project_history_database)):
                    content = project_history_database[i][1]
                    past_projects.append(content)
            program = self.ui.text_fee_3_install_program.toPlainText()
            table = self.ui.table_fee_4_install
            installation_content_list = []
            for i in range(4):
                try:
                    document = table.item(i, 0).text()
                    fee = table.item(i, 1).text()
                    fee_total = table.item(i, 2).text()
                    if document!='':
                        installation_content_list.append([document, fee, fee_total])
                except:
                    pass
            pdf_list = [file for file in os.listdir(accounting_dir) if str(file).endswith(".pdf") and "Mechanical Installation Fee Proposal" in str(file)]
            if len(pdf_list) != 0:
                current_pdf = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
                current_revision = current_pdf.split(" ")[-1].split(".")[0]
                if fee_installation_rev== current_revision or fee_installation_rev == str(int(current_revision) + 1):
                    old_pdf_path = os.path.join(accounting_dir, current_pdf)
                    open_pdf_thread(old_pdf_path)
                    overwrite = messagebox(f"Warning",f"Revision {current_revision} found, do you want to generate Revision {fee_installation_rev}", self.parent.ui)
                    if not overwrite:
                        return
                    else:
                        for proc in psutil.process_iter():
                            if proc.name() == "Acrobat.exe":
                                proc.kill()
                else:
                    message(f'Current revision is {current_revision}, you can not use revision {fee_installation_rev}',self.parent.ui)
                    return
            else:
                if fee_installation_rev != "1":
                    message("There is no other existing fee proposal found, can only have revision 1",self.parent.ui)
                    return

            try:
                shutil.copy(os.path.join(conf["template_dir"], "xlsx", f"installation_fee_proposal_template.xlsx"),os.path.join(accounting_dir, excel_name))
            except PermissionError:
                message("Please close the excel before you generate",self.parent.ui)
                return
            install_fee_date=convert_date(install_fee_date)
            write_excel=f_write_installation_fee_excel(excel_dir,full_name,company,address,quo_num,install_fee_date,
                                   fee_installation_rev,pro_name,drawing_content,content_dict,
                                   past_projects,program,installation_content_list)
            if not write_excel:
                message(write_excel,self.parent.ui)
                return
            export_excel_thread(excel_dir, pdf_dir)
        except:
            traceback.print_exc()
    '''4 - Email installation proposal'''
    def f_button_fee_3_install_email(self):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            check_or_create_folder(accounting_dir)
            pdf_list = [file for file in os.listdir(accounting_dir) if "Mechanical Installation Fee Proposal" in str(file) and str(file).endswith(".pdf")]
            pdf_name = sorted(pdf_list, key=lambda pdf: str(pdf).split(" ")[-1].split(".")[0])[-1]
            pdf_dir = os.path.join(accounting_dir, pdf_name)
            quo_num=self.ui.text_proinfo_1_quonum.text()
            revision=self.ui.text_fee_1_revision.text()
            if not os.path.exists(pdf_dir):
                message(f'Python cant found fee proposal for {quo_num} revision {revision}',self.parent.ui)
                return
            contact_type = self.get_contact_type()
            if contact_type == 'Client':
                full_name = self.ui.text_proinfo_2_clientname.text()
            else:
                full_name = self.ui.text_proinfo_2_contactname.text()
            first_name = get_first_name(full_name)
            client_email = self.ui.text_proinfo_2_clientemail.text()
            main_contact_email = self.ui.text_proinfo_2_contactemail.text()
            if client_email=='0@com.au':
                client_email=''
            if main_contact_email=='0@com.au':
                main_contact_email=''
            pro_name = self.ui.text_proinfo_1_proname.text()
            message2email = f"""
                Dear {first_name},<br>
                <br>
                I hope this email finds you well. Please find the fee proposal attached for your review and approval.<br>
                <br>
                If you have any questions or need more information, please feel free to give us a call.<br>
                <br>
                We look forward to contributing to the project.<br>
                """
            subject = f'{quo_num}-Mechanical Installation Fee Proposal for {pro_name}'
            email2cc = ''
            f_email_fee(subject,client_email,main_contact_email,email2cc,message2email,pdf_dir)
            update_asana = messagebox('Update Asana', "Do you want to update Asana?", self.parent.ui)
            if update_asana:
                update_asana_result=self.f_button_pro_func_updateasana(False)
                quo_num = self.ui.text_proinfo_1_quonum.text()
                pro_num = self.ui.text_proinfo_1_pronum.text()
                pro_name = self.ui.text_proinfo_1_proname.text()
                pro_info = quo_num + '-' + pro_num + '-' + pro_name
                if update_asana_result == 'Success':
                    message(f'Project {pro_info} \nAsana Updated Successfully', self.parent.ui)
                elif update_asana_result == 'Fail':
                    message(f'Project {pro_info} \nUpdate Asana failed, please contact admin', self.parent.ui)
                else:
                    message(update_asana_result, self.parent.ui)

        except:
            traceback.print_exc()
            message("Unable to Create Email",self.parent.ui)
    '''4 - Scope content save as default'''
    def f_save_scope2default(self, service_short):
        try:
            if self.radiobuttongroup_proinfo_1_proposaltype.checkedId()==1:
                scope_type='Minor'
            elif self.radiobuttongroup_proinfo_1_proposaltype.checkedId()==2:
                scope_type='Major'
            else:
                message('Choose proposal type first',self.parent.ui)
            service=conf['Service full dict'][service_short]
            stage_name_list=['Extent','Clarifications','Deliverables']
            content_list=[]
            for i in range(3):
                table=self.ui_name_dict["table_fee_3_"+ service_short +'_'+str(i+1)]
                selected_rows = table.selectionModel().selectedRows()
                if not selected_rows:
                    message('All tables in the section to be selected to set as default.',self.parent.ui)
                    return
                selected_row_indices = [row.row() for row in selected_rows]
                for j in range(len(selected_row_indices)):
                    row=selected_row_indices[j]
                    stage=stage_name_list[i]
                    content = self.f_tableitem2none(table.item(row, 1))
                    checkbox = table.item(row, 0)
                    include=0
                    if checkbox is not None and checkbox.flags() & Qt.ItemIsUserCheckable:
                        if checkbox.checkState() == Qt.Checked:
                            include=1
                    content_dict_i = {"stage":stage,"content":content,"include":include,"row_index":j}
                    content_list.append(content_dict_i)
        except:
            traceback.print_exc()
            message('Can not set scope as default',self.parent.ui)
            return
        try:
            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
            mycursor = mydb.cursor()
            mycursor.execute("DELETE FROM bridge.default_scope WHERE scope_type=%s AND service=%s", (scope_type,service))
            for scope_dict in content_list:
                mycursor.execute("INSERT INTO bridge.default_scope VALUES (%s, %s, %s, %s, %s, %s)",
                                 (scope_type, service, scope_dict["stage"], scope_dict["content"], scope_dict["include"], scope_dict["row_index"]))
            mydb.commit()
            message('Scope table set as default',self.parent.ui)
        except:
            traceback.print_exc()
        finally:
            mycursor.close()
            mydb.close()
    '''5 - Fee calculation table calculate'''
    def f_fee_2_calculate_total_fee(self,table_name):
        try:
            fee_calculation_table_dict = {"Minor": (self.ui.text_fee_5_minor_area, self.ui.text_fee_5_minor_price, self.ui.text_fee_5_minor_total,self.ui.table_fee_5_cal_area),
                                          "Major": (self.ui.text_fee_5_major_apt, self.ui.text_fee_5_major_price,self.ui.text_fee_5_major_total, self.ui.table_fee_5_cal_apt)}
            if is_float(fee_calculation_table_dict[table_name][0].text()):
                basic_num = float(fee_calculation_table_dict[table_name][0].text())
            else:
                basic_num=0
            if is_float(fee_calculation_table_dict[table_name][1].text()):
                price=float(fee_calculation_table_dict[table_name][1].text())
            else:
                price=0
            fee_text=fee_calculation_table_dict[table_name][2]
            fee_total=round(basic_num*price,2)
            fee_text.setText(str(fee_total))
        except:
            traceback.print_exc()
    def f_fee_calculation_table_change(self,table_name):
        try:
            fee_calculation_table_dict={"Minor":(self.ui.text_fee_5_minor_area,self.ui.text_fee_5_minor_price,self.ui.text_fee_5_minor_total,self.ui.table_fee_5_cal_area),
                                        "Major":(self.ui.text_fee_5_major_apt,self.ui.text_fee_5_major_price,self.ui.text_fee_5_major_total,self.ui.table_fee_5_cal_apt)}
            basic_num=fee_calculation_table_dict[table_name][0]
            table=fee_calculation_table_dict[table_name][3]
            for row in range(table.rowCount()):
                set=int(table.columnCount()/2)
                for set_i in range(set):
                    column_price=set_i*2
                    column_fee = set_i*2+1
                    item = table.item(row,column_price).text()
                    price_i=item[1:]
                    if is_float(price_i) and is_float(basic_num.text()):
                        fee_i=float(price_i)*float(basic_num.text())
                    else:
                        fee_i=0
                    table.setItem(row, column_fee, QTableWidgetItem(str(fee_i)))
            self.f_set_table_style(table)
            self.f_fee_2_calculate_total_fee(table_name)


        except:
            traceback.print_exc()

    '''=============================Financial page=================================='''
    '''0 - Invoice '''
    def f_text_finan_inv_num(self, index):
        try:
            text_name=self.ui_name_dict['text_finan_1_inv'+str(index)]
        except:
            traceback.print_exc()
    '''0 - Invoice contact text changed'''
    def f_text_finan_client(self, text_id):
        try:
            text_name=self.ui_name_dict["text_finan_1_invclient_" + str(text_id)]
            self.last_inv_client_id=text_id
            search_table=self.ui.table_finan_inv_client_search
            client_info_list=[]
            search_content=text_name.text()
            if search_content=='':
                search_table.hide()
                return
            if search_content=='-':
                contact_id_list=[]
                contact_type_list=[]
                if self.datanow["Client id"]!=None:
                    contact_id_list.append(self.datanow["Client id"])
                    contact_type_list.append('Client')
                if self.datanow["Main contact id"]!=None:
                    contact_id_list.append(self.datanow["Main contact id"])
                    contact_type_list.append('Main contact')
                if contact_id_list!=[]:
                    self.ui.table_finan_inv_client_search.setRowCount(0)
                    for i in range(len(contact_id_list)):
                        client_id=contact_id_list[i]
                        try:
                            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                            mycursor = mydb.cursor()
                            mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                            client_database = mycursor.fetchall()
                            client_database = client_database[0]
                            full_name = client_database[2]
                            company = client_database[4]
                            email = client_database[8]
                            self.ui.table_finan_inv_client_search.insertRow(i)
                            self.ui.table_finan_inv_client_search.setItem(i, 0, QTableWidgetItem(str(client_id)))
                            self.ui.table_finan_inv_client_search.setItem(i, 1, QTableWidgetItem(contact_type_list[i]))
                            self.ui.table_finan_inv_client_search.setItem(i, 2, QTableWidgetItem(str(full_name)))
                            self.ui.table_finan_inv_client_search.setItem(i, 3, QTableWidgetItem(str(company)))
                            self.ui.table_finan_inv_client_search.setItem(i, 4, QTableWidgetItem(str(email)))
                        except:
                            traceback.print_exc()
                        finally:
                            mycursor.close()
                            mydb.close()
                    self.f_set_table_style(self.ui.table_finan_inv_client_search)
                    self.ui.table_finan_inv_client_search.show()
            else:
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("SELECT * from bridge.clients WHERE LOWER(full_name) LIKE LOWER(%s) OR LOWER(company) LIKE LOWER(%s) OR LOWER(contact_email) LIKE LOWER(%s)",
                                     (f'%{search_content}%',f'%{search_content}%',f'%{search_content}%',))
                    search_results = mycursor.fetchall()
                except:
                    traceback.print_exc()
                    return
                finally:
                    mycursor.close()
                    mydb.close()
                for row in search_results:
                    client_info_list_i=[row[0],row[2],row[4],row[8]]
                    client_info_list.append(client_info_list_i)
                search_table.setRowCount(0)
                search_table.show()
                for clent_info in client_info_list:
                    table_row_now = search_table.rowCount()
                    search_table.insertRow(table_row_now)
                    search_table.setItem(table_row_now, 0, QTableWidgetItem(str(clent_info[0])))
                    search_table.setItem(table_row_now, 2, QTableWidgetItem(str(clent_info[1])))
                    search_table.setItem(table_row_now, 3, QTableWidgetItem(str(clent_info[2])))
                    search_table.setItem(table_row_now, 4, QTableWidgetItem(str(clent_info[3])))
                self.f_set_table_style(search_table)
        except:
            traceback.print_exc()
    '''0 - Invoice contact search table clicked'''
    def f_search_table_click2(self):
        try:
            table=self.ui.table_finan_inv_client_search
            current_row = table.currentRow()
            if current_row >= 0:
                client_id=table.item(current_row, 0).text()
                client_name=table.item(current_row, 2).text()
                company_name=table.item(current_row, 3).text()
                client_company_name=str(client_name)+'--'+str(company_name)
                text_name=self.ui_name_dict["text_finan_1_invclient_" + str(self.last_inv_client_id)]
                text_name.setText(client_company_name)
                self.datanow["Client invoice id"][self.last_inv_client_id-1]=client_id
                table.hide()
        except:
            traceback.print_exc()
    '''0 - Bill contact text changed'''
    def f_text_bill_client(self,service, text_id):
        try:
            text_name=self.ui_name_dict["text_finan_2_" + service + '_' + str(text_id) + '_2']
            self.last_bill_info=(service,text_id)
            search_table=self.ui.table_finan_bill_client_search
            client_info_list=[]
            search_content=text_name.text()
            if search_content=='':
                search_table.hide()
                text_bill_num=self.ui_name_dict["text_finan_2_" + service + '_' + str(text_id) + '_1']
                text_bill_num.setText('')
                return
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * from bridge.clients WHERE LOWER(full_name) LIKE LOWER(%s) OR LOWER(company) LIKE LOWER(%s) OR LOWER(contact_email) LIKE LOWER(%s)",
                                 (f'%{search_content}%',f'%{search_content}%',f'%{search_content}%',))
                search_results = mycursor.fetchall()
            except:
                traceback.print_exc()
                return
            finally:
                mycursor.close()
                mydb.close()
            for row in search_results:
                client_info_list_i=[row[0],row[2],row[4],row[8]]
                client_info_list.append(client_info_list_i)
            search_table.setRowCount(0)
            search_table.show()
            for clent_info in client_info_list:
                table_row_now = search_table.rowCount()
                search_table.insertRow(table_row_now)
                search_table.setItem(table_row_now, 0, QTableWidgetItem(str(clent_info[0])))
                search_table.setItem(table_row_now, 1, QTableWidgetItem(str(clent_info[1])))
                search_table.setItem(table_row_now, 2, QTableWidgetItem(str(clent_info[2])))
                search_table.setItem(table_row_now, 3, QTableWidgetItem(str(clent_info[3])))
            self.f_set_table_style(search_table)
        except:
            traceback.print_exc()
    '''0 - Bill contact search table clicked'''
    def f_search_table_click3(self):
        try:
            table=self.ui.table_finan_bill_client_search
            current_row = table.currentRow()
            if current_row >= 0:
                client_id=table.item(current_row, 0).text()
                client_name=table.item(current_row, 1).text()
                text_name=self.ui_name_dict["text_finan_2_" + self.last_bill_info[0] + '_' + str(self.last_bill_info[1]) + '_2']
                bill_num_text=self.ui_name_dict["text_finan_2_" + self.last_bill_info[0] + '_' + str(self.last_bill_info[1]) + '_1']
                bill_num_text.setText(' ')
                text_name.setText(client_name)
                bill_num=self.ui_name_dict["text_finan_2_" + self.last_bill_info[0] + '_' + str(self.last_bill_info[1]) + '_1'].text()
                self.datanow["Client bill id"][char2num(bill_num)]=client_id
                table.hide()
        except:
            traceback.print_exc()
    '''0 - Get new invoice number'''
    def f_text_gen_quo_inv_num(self, text_id):
        try:
            fee_text=self.ui_name_dict["text_finan_1_total" + str(text_id)+'_1']
            inv_text=self.ui_name_dict["text_finan_1_inv" + str(text_id)]
            if inv_text.text()=='' and is_float(fee_text.text()):
                if float(fee_text.text())>0:
                    quo_num=self.ui.text_proinfo_1_quonum.text()
                    new_inv_num=quo_num[:3]+quo_num[-2:]+'-'+str(text_id)
                    inv_text.setText(new_inv_num)
        except:
            traceback.print_exc()
    '''0 - Gen invoice number'''
    def f_button_finan_1_gen(self):
        try:
            state=self.ui.combobox_pro_1_state.currentText()
            if state in ['Set Up','Gen Fee Proposal','Email Fee Proposal','Chase Fee Acceptance']:
                message("You need to upload a fee acceptance first",self.parent.ui)
                return
            generate = messagebox('Generate invoice',"Are you sure you want to generate the invoice number?", self.parent.ui)
            if generate:
                inv_now=8
                for i in range(8):
                    inv_text=self.ui_name_dict["text_finan_1_inv" + str(i+1)]
                    if inv_text.text()=='' or inv_text.text().find('-')!=-1:
                        inv_now=i
                        break
                if inv_now==8:
                    message("You cant generate more than 8 invoices",self.parent.ui)
                    return
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("SELECT MAX(invoice_number) AS max_value FROM invoices WHERE invoice_number NOT LIKE '%-%' AND invoice_number NOT LIKE '9%'")
                    current_inv_number = str(mycursor.fetchall()[0][0])
                except:
                    traceback.print_exc()
                    return
                finally:
                    mycursor.close()
                    mydb.close()
                if current_inv_number.startswith(date.today().strftime("%y")[1]):
                    current_number = current_inv_number[3:6]
                    res = date.today().strftime("%y%m")[1:] + str(int(current_number) + 1).zfill(3)
                else:
                    res = date.today().strftime("%y%m")[1:] + "001"
                text_inv_new=self.ui_name_dict["text_finan_1_inv" + str(inv_now+1)]
                finan_text_list=[self.ui_name_dict["text_finan_1_total" + str(inv_now + 1)+'_'+str(j+1)] for j in range(5)]
                self.last_inv_client_id=inv_now+1
                text_inv_new.setText(res)
                text_inv_new.setStyleSheet(conf['Text gray style'])
                for finan_text in finan_text_list:
                    finan_text.setStyleSheet(conf['Text gray_white style'])
                action_content='Generate Invoice '+res
                self.f_add_user_history(action_content)
                quotation_number=self.ui.text_proinfo_1_quonum.text()
                fee_amount=self.ui_name_dict["text_finan_1_total" + str(inv_now+1)+'_1'].text()
                contact_id_list=[]
                contact_type_list=[]
                if self.datanow["Client id"]!=None:
                    contact_id_list.append(self.datanow["Client id"])
                    contact_type_list.append('Client')
                if self.datanow["Main contact id"]!=None:
                    contact_id_list.append(self.datanow["Main contact id"])
                    contact_type_list.append('Main contact')
                print(contact_id_list)
                if contact_id_list!=[]:
                    self.ui.table_finan_inv_client_search.setRowCount(0)
                    for i in range(len(contact_id_list)):
                        client_id=contact_id_list[i]
                        try:
                            mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                            mycursor = mydb.cursor()
                            mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                            client_database = mycursor.fetchall()
                            client_database = client_database[0]
                            full_name = client_database[2]
                            company = client_database[4]
                            email = client_database[8]
                            self.ui.table_finan_inv_client_search.insertRow(i)
                            self.ui.table_finan_inv_client_search.setItem(i, 0, QTableWidgetItem(str(client_id)))
                            self.ui.table_finan_inv_client_search.setItem(i, 1, QTableWidgetItem(contact_type_list[i]))
                            self.ui.table_finan_inv_client_search.setItem(i, 2, QTableWidgetItem(str(full_name)))
                            self.ui.table_finan_inv_client_search.setItem(i, 3, QTableWidgetItem(str(company)))
                            self.ui.table_finan_inv_client_search.setItem(i, 4, QTableWidgetItem(str(email)))
                        except:
                            traceback.print_exc()
                        finally:
                            mycursor.close()
                            mydb.close()
                    self.f_set_table_style(self.ui.table_finan_inv_client_search)
                    self.ui.table_finan_inv_client_search.show()
                try:
                    mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                    mycursor = mydb.cursor()
                    mycursor.execute("INSERT INTO bridge.invoices VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                         (res, quotation_number, 'Backlog',None,None,None,None,0,fee_amount,inv_now))
                    mydb.commit()
                except:
                    traceback.print_exc()
                    return
                finally:
                    mycursor.close()
                    mydb.close()
        except:
            traceback.print_exc()

    '''0 - Delete invoice number'''
    def f_button_finan_1_del(self):
        try:
            del_inv_num=''
            for i in range(8):
                inv_num_text_origin=self.ui_name_dict["text_finan_1_inv" + str(8-i)]
                inv_num=inv_num_text_origin.text()
                palette = inv_num_text_origin.palette()
                background_color = palette.color(palette.Window).name()
                if inv_num.find('-')!=-1 or background_color=="#d2d2d2":
                    del_inv_i=i
                    del_inv_num=inv_num
                    break
            if del_inv_num!='':
                fee_inv_del=self.ui_name_dict["text_finan_1_total" + str(8-del_inv_i)+'_1'].text()
                if is_float(fee_inv_del):
                    if float(fee_inv_del)>0:
                        message('Amount not equals to 0, please change amount.',self.parent.ui)
                        return

                delete_inv=messagebox('delete',f'Do you want to delete {del_inv_num}?',self.parent.ui)
                if delete_inv:
                    inv_text=self.ui_name_dict["text_finan_1_inv" + str(8-del_inv_i)]
                    inv_text.setText('')
                    try:
                        mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                        mycursor = mydb.cursor()
                        mycursor.execute("DELETE FROM bridge.invoices"
                                         " WHERE invoice_number=%s", (del_inv_num,))
                        mydb.commit()
                        message(f'Invoice {del_inv_num} is deleted, and you have to delete related files manually.',self.parent.ui)
                    except:
                        traceback.print_exc()
                    finally:
                        mycursor.close()
                        mydb.close()
        except:
            traceback.print_exc()

    '''1 - Calculate payment amount'''
    def f_calculate_sum_payment(self):
        try:
            payment_sum=0
            for i in range(8):
                payment=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_3'].text()
                if is_float(payment):
                    payment_sum+=float(payment)
            self.ui.text_finan_1_total0_3.setText(str(payment_sum))
        except:
            traceback.print_exc()
    '''1 - Set payment time'''
    def f_set_payment_date(self):
        try:
            for i in range(8):
                payment_date_1=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_4'].text()
                payment_date_2=payment_date_1.replace("-", "")
                payment_date_2_text=self.ui_name_dict["text_finan_1_total" + str(i+1)+'_5']
                payment_date_2_text.setText(payment_date_2)
        except:
            traceback.print_exc()


    '''1- Calculate invoice total'''
    def f_radbutton_finan_change(self):
        try:
            service_full_list=self.f_get_pro_service_type('Frame_with_var')
            service_list=[]
            for service in service_full_list:
                service_short=conf['Service short dict'][service]
                service_list.append(service_short)
            fee_total_list=[0 for _ in range(8)]
            for i in range(len(service_list)):
                for j in range(4):
                    fee=self.ui_name_dict["text_finan_1_" +service_list[i]+ '_fee'+str(j+1)+'_1'].text()
                    button_group=self.radiobutton_groups_dict[service_list[i] + '_' + str(j + 1)]
                    selected_id=button_group.checkedId()
                    if selected_id!=-1:
                        if is_float(fee):
                            fee_total_list[selected_id-1]+=float(fee)
            for i in range(8):
                total_text=self.ui_name_dict["text_finan_1_total" + str(i+1) + "_1"]
                total_text.setText(str(fee_total_list[i]))
                total_gst_text=self.ui_name_dict["text_finan_1_total" + str(i+1) + "_2"]
                total_gst_text.setText(str(round(fee_total_list[i]*1.1,1)))
                self.f_finan_total_profit()
        except:
            traceback.print_exc()
    '''2- Bill invoice upload'''
    def f_button_upload_subconsultant_invoice(self,service_name,id):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            bill_description=self.ui_name_dict["text_finan_2_" +service_name+ '_0_1'].text()
            service=conf['Service full dict'][service_name]
            origin_dict={"1":'Origin',"2":'V1',"3":'V2',"4":'V3'}
            origin=origin_dict[id]
            if bill_description=='':
                message("You need to upload the description",self.parent.ui)
                return
            if service == "Fire Service":
                filename = f"{service}-{bill_description}-{origin}"
            else:
                filename = f"{service.split(' ')[0]}-{bill_description}-{origin}"
            file_dir=os.path.join(accounting_dir,filename)
            if file_exists(file_dir):
                rewrite=messagebox('Overwrite',"Existing file found, Do you want to rewrite", self.parent.ui)
                if not rewrite:
                    return
            fileName = self.f_get_filename(accounting_dir)
            if fileName == "":
                return
            try:
                folder_dir = os.path.join(accounting_dir, filename + os.path.splitext(fileName)[1])
                shutil.move(fileName, folder_dir)
            except PermissionError:
                traceback.print_exc()
                message("Please Close the file before you upload it",self.parent.ui)
                return
            except Exception as e:
                traceback.print_exc()
                message("Some error occurs, please contact Administrator",self.parent.ui)
                return
            message(f'Upload from \n {fileName} \n to \n {folder_dir}',self.parent.ui)
            button=self.ui_name_dict["button_finan_2_" +service_name+ '_upload'+id]
            button.setStyleSheet(conf["Gray button style"])

        except:
            traceback.print_exc()
    '''2 - Bill invoice price change connection'''
    def f_text_finan_2_price_change(self, item_name):
        try:
            text_list=[self.ui_name_dict["text_finan_2_" + item_name + "_0_" + str(i + 2)] for i in range(4)]
            price_sum_text=self.ui_name_dict["text_finan_2_" + item_name + "_0_6"]
            total_price=0
            for i in range(len(text_list)):
                price_i=text_list[i].text()
                if is_float(price_i):
                    total_price+=float(price_i)
            if total_price>0:
                price_sum_text.setText(str(total_price))
        except:
            traceback.print_exc()
    '''2 - Bill num change connection'''
    def f_text_finan_2_num_change(self,item_name):
        try:
            ui_text_name=self.ui_name_dict["text_finan_2_" + item_name + '_1']
            if ui_text_name.text()!='':
                num_allow_all_list=['ERROR','A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J','K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T','U', 'V', 'W', 'X', 'Y', 'Z']
                service_list = ['mech', 'cfd', 'install', 'ele', 'hyd', 'fire', 'mechrev', 'mis', 'var']
                num_text_list=[]
                for i in range(len(service_list)):
                    for j in range(3):
                        text=self.ui_name_dict["text_finan_2_" +service_list[i]+ '_'+str(j+1)+'_1'].text()
                        if text!='':
                            num_text_list.append(text)
                if num_text_list!=[]:
                    num_allow=num_allow_all_list[len(num_text_list)]
                else:
                    num_allow = 'A'
                ui_text_name.setText(num_allow)
        except:
            traceback.print_exc()
    '''2 - Bill fee change connection'''
    def f_text_finan_2_fee_change(self,item_name):
        try:
            fee1_text=self.ui_name_dict["text_finan_2_" + item_name+"_4"]
            if is_float(fee1_text.text()):
                checkbox=self.ui_name_dict["checkbox_finan_2_" + item_name]
                fee2_text=self.ui_name_dict["text_finan_2_" + item_name + "_5"]
                if not checkbox.isChecked():
                    fee2=round(float(fee1_text.text())*1.1,1)
                    fee2_text.setText(str(fee2))
                else:
                    fee2_text.setText(fee1_text.text())

            service=item_name.split('_')[0]
            total_fee_text1_list=[self.ui_name_dict["text_finan_2_" + service+'_'+str(i+1)+"_4"] for i in range(3)]
            total_fee_text2_list=[self.ui_name_dict["text_finan_2_" + service + '_' + str(i + 1) + "_5"] for i in range(3)]
            total_fee_text1=self.ui_name_dict["text_finan_2_" + service+'_4_1']
            total_fee_text2=self.ui_name_dict["text_finan_2_" + service + '_4_2']
            total_fee1=0
            total_fee2=0
            for i in range(len(total_fee_text1_list)):
                if is_float(total_fee_text1_list[i].text()):
                    total_fee1+=float(total_fee_text1_list[i].text())
                if is_float(total_fee_text2_list[i].text()):
                    total_fee2 += float(total_fee_text2_list[i].text())
            total_fee_text1.setText(str(round(total_fee1,1)))
            total_fee_text2.setText(str(round(total_fee2, 1)))


            service_list = ['mech', 'cfd', 'ele', 'hyd', 'fire', 'mechrev', 'mis', 'install', 'var']
            fee1_list=[self.ui_name_dict["text_finan_2_" + service_list[i] + "_4_1"] for i in range(len(service_list))]
            fee2_list=[self.ui_name_dict["text_finan_2_" + service_list[i] + "_4_2"] for i in range(len(service_list))]
            fee1_total=0
            fee2_total=0
            for i in range(len(fee1_list)):
                if is_float(fee1_list[i].text()):
                    fee1_total+=float(fee1_list[i].text())
                if is_float(fee2_list[i].text()):
                    fee2_total += float(fee2_list[i].text())
            self.ui.text_finan_2_total1.setText(str(round(fee1_total,1)))
            self.ui.text_finan_2_total2.setText(str(round(fee2_total,1)))
            self.f_finan_total_profit()


        except:
            traceback.print_exc()
    '''2 - Bill price sum change connection'''
    def f_text_finan_2_price_sum_change(self,item_name):
        try:
            price_sum_text=self.ui_name_dict["text_finan_2_" + item_name + "_0_6"]
            price_total_text=self.ui_name_dict["text_finan_2_" + item_name + "_0_7"]
            checkbox=self.ui_name_dict["checkbox_finan_2_" + item_name + "_0"]
            if is_float(price_sum_text.text()):
                if checkbox.isChecked():
                    price_total=round(float(price_sum_text.text()),1)
                else:
                    price_total = round(float(price_sum_text.text())*1.1,1)
                price_total_text.setText(str(price_total))
            self.f_text_finan_3_profit_calculation(item_name)
        except:
            traceback.print_exc()
    '''2- Bill upload'''
    def f_button_upload_subconsultant_bill(self, service_name, id):
        try:
            pass
        except:
            traceback.print_exc()
    '''2 - Upload bill to xero'''
    def f_finan_bill_upload(self,service,id):
        try:
            self.f_button_pro_func_refreshxero()
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            pro_num=self.ui.text_proinfo_1_pronum.text()
            if pro_num=='':
                message("This Project dont have a Project Number yet",self.parent.ui)
                return
            bill_letter=self.ui_name_dict["text_finan_2_" + service + '_' + str(id)+'_1'].text()
            bill_number=pro_num+bill_letter
            bills_dir = os.path.join(conf["bills_dir"], date.today().strftime("%Y%m"))
            bills_dir_out=conf["bills_dir"]
            if bill_letter=='':
                message("You need to enter a bill number",self.parent.ui)
                return
            contact_name=self.ui_name_dict["text_finan_2_" + service + '_' + str(id)+'_2'].text()
            if contact_name=='':
                message("Please enter the contact before you upload bill",self.parent.ui)
                return
            bill_amount_text=self.ui_name_dict["text_finan_2_" + service + '_'+str(id)+'_4'].text()
            if bill_amount_text=='':
                message("Please enter the bill amount before you upload bill",self.parent.ui)
                return
            fee_invoice_text=self.ui_name_dict["text_finan_2_" + service + '_0_6'].text()
            fee_bill_text=self.ui_name_dict["text_finan_2_" + service + '_4_1'].text()
            if is_float(fee_invoice_text) and is_float(fee_bill_text):
                if float(fee_invoice_text)<float(fee_bill_text):
                    message("Bill Amount, exceed subbie quote amount",self.parent.ui)
                    return
            else:
                message("Fee amount error",self.parent.ui)
                return
            button=self.ui_name_dict["button_finan_2_" + service + '_upload'+str(id+4)]
            color_dict={"#a52a2a":False,"#c8c8c8":True}
            palette = button.palette()
            background_color = palette.color(palette.Window).name()
            state = color_dict[background_color]
            if state:
                rewrite = messagebox('Rewrite',f"You uploaded this bill before, Do you want to rewrite", self.parent.ui)
                if not rewrite:
                    return
            file = self.f_get_filename(accounting_dir)
            if file == "":
                return
            filename = "BIL " + bill_number + "-" + os.path.basename(file).replace(" ", "_")
            if not os.path.exists(bills_dir):
                os.makedirs(bills_dir)
            # try:
            #     for old_file in os.listdir(accounting_dir):
            #         if old_file.startswith("BIL " + bill_number + "-"):
            #             os.remove(os.path.join(accounting_dir, old_file))
            #     for old_file in os.listdir(bills_dir):
            #         if old_file.startswith("BIL " + bill_number + "-"):
            #             os.remove(os.path.join(bills_dir, old_file))
            #     for old_file in os.listdir(bills_dir_out):
            #         if old_file.startswith("BIL " + bill_number + "-"):
            #             os.remove(os.path.join(bills_dir_out, old_file))
            # except PermissionError:
            #     message("Please Close the old file before you upload the bill, Bridge need to remove the old file",self.parent.ui)
            #     return
            # except:
            #     traceback.print_exc()
            #     message("Some error occurs, please contact Administrator",self.parent.ui)
            #     return
            try:
                folder_path = os.path.join(accounting_dir, filename)
                shutil.move(file, folder_path)
                bill_path_out = os.path.join(bills_dir_out, filename)
                shutil.copy(folder_path, bill_path_out)
            except PermissionError:
                message("Please Close the file before you upload it",self.parent.ui)
                return
            except AssertionError:
                message("You Can not delete a bill in Awaiting payment or Paid Stage",self.parent.ui)
                return
            except:
                traceback.print_exc()
                message("Some error occurs, please contact Administrator",self.parent.ui)
                return
            self.f_button_pro_func_refreshxero()
            try:
                service_full_name=conf['Service full dict'][service]
                amount=round(float(bill_amount_text),1)
                combobox=self.ui_name_dict["combobox_finan_2_" + service + '_'+str(id)]
                bill_type=combobox.currentText()
                checkbox=self.ui_name_dict["checkbox_finan_2_" + service + '_'+str(id)]
                if checkbox.isChecked():
                    no_gst=True
                else:
                    no_gst=False
                row_index = char2num(bill_letter)
                contact_id = self.datanow["Client bill id"][row_index]
                contact_xero_id = f_get_client_xero_id(contact_id)

                try:
                    bill_xero_id=create_xero_bill(service_full_name,amount, bill_type,no_gst,contact_xero_id, bill_path_out)
                except:
                    message(f'Create bill {bill_number} unsuccessfully.',self.parent.ui)
                    traceback.print_exc()
                    return
                self.datanow["Xero bill"][row_index]=bill_xero_id
                self.datanow["Bill in date"][row_index]=date.today().strftime("%Y-%m-%d")
                self.f_button_pro_func_updatexero()
            except:
                traceback.print_exc()
                message("Unable to upload the File to xero",self.parent.ui)
                return
            message(f'Upload from \n {file} \n to \n {folder_path} \n And the bill has been sent to Xero',self.parent.ui)
            button.setStyleSheet(conf["Gray button style"])
        except:
            traceback.print_exc()
    '''3 - Profit change connection'''
    def f_text_finan_3_profit_calculation(self,item_name):
        try:
            text_invoice_ex=self.ui_name_dict["text_finan_1_" + item_name + "_fee0_1"].text()
            text_bill_ex=self.ui_name_dict["text_finan_2_" + item_name + "_0_6"].text()
            text_invoice_in=self.ui_name_dict["text_finan_1_" + item_name + "_fee0_2"].text()
            text_bill_in=self.ui_name_dict["text_finan_2_" + item_name + "_0_7"].text()
            text_profit_ex=self.ui_name_dict["text_finan_3_" + item_name]
            text_profit_in=self.ui_name_dict["text_finan_3_" + item_name+'_in']
            if not is_float(text_invoice_ex):
                text_invoice_ex=0
            if not is_float(text_bill_ex):
                text_bill_ex=0
            if not is_float(text_invoice_in):
                text_invoice_in=0
            if not is_float(text_bill_in):
                text_bill_in=0
            profit_ex=round(float(text_invoice_ex)-float(text_bill_ex),1)
            text_profit_ex.setText(str(profit_ex))
            if profit_ex<0:
                style = conf['Text red style']
            else:
                style = conf['Text not change style']
            text_profit_ex.setStyleSheet(style)
            profit_in=round(float(text_invoice_in)-float(text_bill_in),1)
            text_profit_in.setText(str(profit_in))
            text_profit_in.setStyleSheet(style)



        except:
            traceback.print_exc()
    '''3 - Total profit calculation'''
    def f_finan_total_profit(self):
        try:
            fee_inv1=self.ui.text_finan_1_total0_1.text()
            fee_inv2 = self.ui.text_finan_1_total0_2.text()
            fee_bill1=self.ui.text_finan_2_total1.text()
            fee_bill2 = self.ui.text_finan_2_total2.text()
            if is_float(fee_inv1) and is_float(fee_bill1):
                profit1=float(fee_inv1)-float(fee_bill1)
                self.ui.text_finan_3_total1.setText(str(round(profit1,1)))
            if is_float(fee_inv2) and is_float(fee_bill2):
                profit2=float(fee_inv2)-float(fee_bill2)
                self.ui.text_finan_3_total2.setText(str(round(profit2,1)))

        except:
            traceback.print_exc()
    '''4 - Invoice preview'''
    def f_invoice_preview(self,id):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            check_or_create_folder(accounting_dir)
            inv_num=self.ui_name_dict["text_finan_1_inv" + str(id)].text()
            if inv_num.find('-')!=-1 or inv_num=='':
                message('You need to generate a invoice number before you generate the Invoice',self.parent.ui)
                return
            fee=self.ui_name_dict["text_finan_1_total" + str(id)+'_1'].text()
            pro_name=self.ui.text_proinfo_1_proname.text()
            if fee=='0' or fee=='':
                process=messagebox('Invoice','The Invoice amount is 0, Do you want to proceed?', self.parent.ui)
                if not process:
                    return
            excel_name = f'PCE INV {inv_num}.xlsx'
            excel_dir=os.path.join(accounting_dir, excel_name)
            invoice_name = f'PCE INV {inv_num}.pdf'
            invoice_dir=os.path.join(accounting_dir, invoice_name)
            if os.path.exists(invoice_dir):
                old_pdf_path = invoice_dir
                open_pdf_thread(old_pdf_path)
                rewrite=messagebox('Warning',f"Existing file PCE INV {inv_num} do you want to rewrite?", self.parent.ui)
                if not rewrite:
                    return
            for proc in psutil.process_iter():
                if proc.name() == "Acrobat.exe":
                    proc.kill()
            shutil.copy(os.path.join(conf["template_dir"], "xlsx", f"invoice_template.xlsx"),excel_dir)
            client_id = self.datanow["Client invoice id"][int(id) - 1]
            try:
                mydb = mysql.connector.connect(host="192.168.1.199", user="bridge", password="PcE$yD2023",database="bridge")
                mycursor = mydb.cursor()
                mycursor.execute("SELECT * FROM bridge.clients WHERE client_id = %s", (client_id,))
                client_database = mycursor.fetchall()
            except:
                traceback.print_exc()
                message('Cant find client information.',self.parent.ui)
                return
            finally:
                mycursor.close()
                mydb.close()
            if len(client_database) == 1:
                client_database = client_database[0]
                full_name = client_database[2]
                company = client_database[4]
                address = client_database[5]
                abn = client_database[6]
            else:
                message('Cant find client information.',self.parent.ui)
            info_list = [full_name, company, address, abn]


            service_full_list = self.f_get_pro_service_type('Frame_with_var')
            total_fee = 0
            total_inGST = 0
            invoice_info_dict={}
            for service_full in service_full_list:
                service_short=conf['Service short dict'][service_full]
                if service_short=='var':
                    if self.ui.text_finan_1_var_fee0_1.text() == '' or self.ui.text_finan_1_var_fee0_1.text() == '0':
                        continue
                invoice_info_dict[service_full]={}
                doc_list=[]
                for i in range(4):
                    doc_name=self.ui_name_dict["text_finan_1_" + service_short + "_doc" + str(i + 1)].text()
                    fee=self.ui_name_dict["text_finan_1_" + service_short + "_fee" + str(i + 1) + '_1'].text()
                    fee_ingst=self.ui_name_dict["text_finan_1_" + service_short + "_fee" + str(i + 1) + '_2'].text()
                    if doc_name!='' and fee!='' and fee != '0' and fee != '0.0':
                        radiobutton=self.ui_name_dict["radbutton_finan_1_" + service_short + '_' + str(i + 1) + '_' + str(id)]
                        if radiobutton.isChecked():
                            check=True
                            total_fee += float(fee)
                            total_inGST += float(fee_ingst)
                        else:
                            check=False
                        doc_list_i=[doc_name,fee,fee_ingst,check]
                        doc_list.append(doc_list_i)
                invoice_info_dict[service_full]['doc_list']=doc_list
                fee_service=self.ui_name_dict["text_finan_1_" + service_short + "_fee0_1"].text()
                invoice_info_dict[service_full]['fee_service']=fee_service

            write_excel=f_write_invoice_excel(excel_dir,info_list,inv_num,pro_name,invoice_info_dict,total_fee,total_inGST)
            if not write_excel:
                message(write_excel,self.parent.ui)
                return
            export_excel_thread(excel_dir, invoice_dir)
            button=self.ui_name_dict["button_finan_4_preview" + str(id)]
            button.setStyleSheet(conf["Gray button style"])
        except:
            traceback.print_exc()
    '''4 - Email invoice'''
    def f_finan_email_invoice(self,id):
        try:
            pro_name = self.ui.text_proinfo_1_proname.text()
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            inv_number=self.ui_name_dict["text_finan_1_inv" + str(id)].text()
            pdf_name='PCE INV '+inv_number+'.pdf'
            pdf_dir=os.path.join(accounting_dir,pdf_name)
            if not file_exists(pdf_dir):
                message('Please generate the invoice before you send to client',self.parent.ui)
                return
            subject=f'PCE INV {inv_number}-{pro_name}'
            client_email = self.ui.text_proinfo_2_clientemail.text()
            main_contact_email = self.ui.text_proinfo_2_contactemail.text()
            if client_email=='0@com.au':
                client_email=''
            if main_contact_email=='0@com.au':
                main_contact_email=''
            contact_type = self.get_contact_type()
            if contact_type == 'Client':
                full_name = self.ui.text_proinfo_2_clientname.text()
            else:
                full_name = self.ui.text_proinfo_2_contactname.text()
            first_name = get_first_name(full_name)
            message2email = f"""
            Hi {first_name},<br>
            <br>
            I hope this email finds you well.<br>
            <br>
            Please find the attached invoice for the payment. We appreciate your prompt attention to this matter.<br>
            <br>
            If you have any questions or concerns regarding the invoice, please do not hesitate to contact us.<br>
            <br>
            """
            f_email_invoice(subject, client_email, main_contact_email, message2email, pdf_dir)
        except:
            traceback.print_exc()

    def f_cal_time(self,message,last_time,time_now):
        try:
            print(f"Time {message}: {time_now - last_time:.2f} seconds")

        except:
            traceback.print_exc()

    def f_finan_email(self,id):
        try:
            inv_number=self.ui_name_dict["text_finan_1_inv" + str(id)].text()
            if inv_number.find('-')!=-1:
                message('Generate valid invoice number first.',self.parent.ui)
                return
            project_num = self.ui.text_proinfo_1_pronum.text()
            if id == 1 and project_num == '':
                if not self.check_folder_files_open(self.datanow["Current folder"]):
                    message(f'File in {self.datanow["Current folder"]} is open. Close files and click the button again.',self.parent.ui)
                    return
            self.thread_2.start()
            self.f_finan_email_invoice(id)
            self.f_button_pro_func_refreshxero()
            service_list = self.f_get_pro_service_type('Frame_with_var')
            contact_id = self.datanow["Client invoice id"][id - 1]
            contact_xero_id=f_get_client_xero_id(contact_id)
            project_name = self.ui.text_proinfo_1_proname.text()
            project_type = self.f_get_project_type()
            items = []
            for service_full_name in service_list:
                service_short_name = conf['Service short dict'][service_full_name]
                for i in range(4):
                    radiobutton=self.ui_name_dict["radbutton_finan_1_" + service_short_name + '_' + str(i + 1) + '_' + str(id)]
                    if radiobutton.isChecked():
                        doc_name=self.ui_name_dict["text_finan_1_" + service_short_name + "_doc" + str(i + 1)].text()
                        fee = round(float(self.ui_name_dict["text_finan_1_" + service_short_name + "_fee" + str(i + 1) + '_1'].text()), 1)
                        item_i = {"Item": doc_name, "Fee": fee}
                        items.append(item_i)
            number=self.ui_name_dict["text_finan_1_inv" + str(id)].text()
            if id==1 and project_num=='':
                current_folder=self.datanow["Current folder"]
                if not file_exists(current_folder):
                    message(f"Can not find the folder {current_folder}",self.parent.ui)
                    return
                old_folder, new_folder, update = change_quotation_number(current_folder, number)
                if not update:
                    message(f"Fail to rename the folder from {old_folder} to {new_folder}, Please close all the file relate in the folder",self.parent.ui)
                    return
                self.ui.text_proinfo_1_pronum.setText(number)
                self.datanow["Current folder"]=new_folder
                invoice_id=self.datanow["Xero invoice id"][id-1]
                if invoice_id==None:
                    try:
                        new_invoice_id=create_xero_invoice(number, items, project_type, project_name, contact_xero_id)
                    except:
                        message(f'Create invoice {number} unsuccessfullfy.',self.parent.ui)
                        traceback.print_exc()
                        return
                    self.datanow["Xero invoice id"][id - 1]=new_invoice_id
                update = messagebox('Update',f'Folder renamed from \n {old_folder} \n to \n {new_folder}. \nDo you want to update Xero and Asana?', self.parent.ui)
                if update:
                    self.f_button_pro_func_updatexero()
            else:
                invoice_id=self.datanow["Xero invoice id"][id-1]
                if invoice_id==None:
                    try:
                        new_invoice_id=create_xero_invoice(number, items, project_type, project_name, contact_xero_id)
                    except:
                        message(f'Create invoice {number} unsuccessfully.',self.parent.ui)
                        traceback.print_exc()
                        return
                    self.datanow["Xero invoice id"][id - 1]=new_invoice_id
                update = messagebox('Update',f'Do you want to update Xero and Asana?', self.parent.ui)
                if update:
                    self.f_button_pro_func_updatexero()
            button=self.ui_name_dict["button_finan_4_email" + str(id)]
            button.setStyleSheet(conf["Gray button style"])
        except:
            traceback.print_exc()
    '''4 - Invoice remit'''
    def f_invoice_remit(self,part,id):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            remittance_dir = os.path.join(conf["remittances_dir"], date.today().strftime("%Y%m"))
            color_dict={"#a52a2a":False,"#c8c8c8":True}
            button_list=[self.ui_name_dict["button_finan_4_remitfull" + str(id)]]

            for i in range(3):
                button_list.append(self.ui_name_dict["button_finan_4_remit" + str(id)+'_'+str(i+1)])
            remit_state_list=[]
            for i in range(len(button_list)):
                palette = button_list[i].palette()
                background_color = palette.color(palette.Window).name()
                state = color_dict[background_color]
                remit_state_list.append(state)
            inv_num=self.ui_name_dict["text_finan_1_inv" + str(id)].text()
            fee_in_total=self.ui_name_dict["text_finan_1_total" + str(id)+'_2'].text()
            fee1=self.ui_name_dict["text_finan_4_remit_" + str(id)+'_1'].text()
            fee2=self.ui_name_dict["text_finan_4_remit_" + str(id) + '_2'].text()
            fee3=self.ui_name_dict["text_finan_4_remit_" + str(id) + '_3'].text()
            if part=='Full':
                if remit_state_list[1]==True or remit_state_list[2]==True or remit_state_list[3]==True:
                    message("This invoice already has partial amount remittance",self.parent.ui)
                    return
                filename = f"{inv_num}, ${fee_in_total}-${fee_in_total}"
            else:
                if remit_state_list[0]:
                    message("This invoice already has full amount remittance",self.parent.ui)
                    return
                if part == "Part1":
                    if not is_float(fee1):
                        message("The first remittance amount is wrong",self.parent.ui)
                        return
                    if float(fee1)>float(fee_in_total):
                        message("The amount exceed the full amount",self.parent.ui)
                        return
                    amount = float(fee1)
                elif part == "Part2":
                    if not remit_state_list[1]:
                        message("Please upload the first remittance first",self.parent.ui)
                        return
                    if not is_float(fee1):
                        message("The first remittance amount is wrong",self.parent.ui)
                        return
                    if not is_float(fee2):
                        message("The second remittance amount is wrong",self.parent.ui)
                        return
                    if float(fee1)+float(fee2)>float(fee_in_total):
                        message("The amount exceed the full amount",self.parent.ui)
                        return
                    amount = float(fee1)+float(fee2)
                else:
                    if not remit_state_list[1]:
                        message("Please upload the first remittance first",self.parent.ui)
                        return
                    if not remit_state_list[2]:
                        message("Please upload the second remittance first",self.parent.ui)
                        return
                    if not remit_state_list[3]:
                        message("The first remittance amount is wrong",self.parent.ui)
                        return
                    if not is_float(fee1):
                        message("The first remittance amount is wrong",self.parent.ui)
                        return
                    if not is_float(fee2):
                        message("The second remittance amount is wrong",self.parent.ui)
                        return
                    if not is_float(fee3):
                        message("The third remittance amount is wrong",self.parent.ui)
                        return
                    if float(fee1) + float(fee2) + float(fee3) > float(fee_in_total):
                        message("The amount exceed the full amount",self.parent.ui)
                        return
                    amount = float(fee1) + float(fee2) + float(fee3)
                    if amount!=fee_in_total:
                        message(f"The Amount does not match, input amount {amount}, invoice ingst {fee_in_total}",self.parent.ui)
                        return
                filename = f"{inv_num}, ${amount}-${fee_in_total}"
            part_dict={"Full":(remit_state_list[0],0),"Part1":(remit_state_list[1],1),"Part2":(remit_state_list[2],2),"Part3":(remit_state_list[3],3)}
            already_upload=part_dict[part][0]
            if already_upload:
                rewrite = messagebox('Rewrite',"Existing file found, Do you want to rewrite", self.parent.ui)
                if not rewrite:
                    return
            file = self.f_get_filename(accounting_dir)
            if file == "":
                return
            try:
                folder_path = os.path.join(accounting_dir, filename + os.path.splitext(file)[1])
                remittance_path = os.path.join(remittance_dir, filename + os.path.splitext(file)[1])
                shutil.move(file, folder_path)
                check_or_create_folder(remittance_dir)
                shutil.copy(folder_path, remittance_path)
            except PermissionError:
                traceback.print_exc()
                message("Please Close the file before you upload it",self.parent.ui)
                return
            except:
                traceback.print_exc()
                message("Some error occurs, please contact Administrator",self.parent.ui)
                return
            button=button_list[part_dict[part][1]]
            button.setStyleSheet(conf["Gray button style"])
            message(f'Upload from \n {file} \n to \n {folder_path}',self.parent.ui)
        except:
            traceback.print_exc()
    '''5- Upload fee acceptance'''
    def f_button_finan_5_upload(self):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            state=self.ui.combobox_pro_1_state.currentText()
            if state in ['Set Up','Gen Fee Proposal','Email Fee Proposal']:
                message("You need to send the Email to Client first",self.parent.ui)
                return
            elif state in ['Design','DWG drawings','Done','Installation','Construction Phase']:
                update=messagebox('Fee acceptance','Fee Already Accepted, do you want to update again?', self.parent.ui)
                if not update:
                    return
            acceptance_list = [file for file in os.listdir(os.path.join(accounting_dir)) if str(file).startswith("Fee Acceptance")]
            if len(acceptance_list) != 0:
                current_revision = str(max([str(pdf).split(" ")[-1].split(".")[0] for pdf in acceptance_list]))
                overwrite=messagebox('Overwrite',f"Current Fee Acceptance {current_revision} found, do you want to Overwrite", self.parent.ui)
                if overwrite:
                    filename = f"Fee Acceptance Rev {current_revision}"
                else:
                    filename = f"Fee Acceptance Rev {str(int(current_revision) + 1)}"
            else:
                filename = "Fee Acceptance Rev 1"
            file = self.f_get_filename(accounting_dir)
            if file == "":
                return
            try:
                folder_dir = os.path.join(accounting_dir, filename + os.path.splitext(file)[1])
                shutil.move(file, folder_dir)
            except PermissionError:
                traceback.print_exc()
                message("Please Close the file before you upload it",self.parent.ui)
                return
            except Exception as e:
                traceback.print_exc()
                message("Some error occurs, please contact Administrator",self.parent.ui)
                return
            self.ui.combobox_pro_1_state.setCurrentText('Design')
            self.ui.button_finan_5_upload.setStyleSheet(conf["Gray button style"])
            self.datanow["Fee accepted date"]=date.today().strftime("%Y-%m-%d")
            self.f_add_user_history('Log fee accept file')
            update_asana=messagebox('Update Asana',f'Upload from \n {file} \n To \n {folder_dir} \n Do you want to Update Asana', self.parent.ui)
            if update_asana:
                update_asana_result=self.f_button_pro_func_updateasana(True)
                quo_num = self.ui.text_proinfo_1_quonum.text()
                pro_num = self.ui.text_proinfo_1_pronum.text()
                pro_name = self.ui.text_proinfo_1_proname.text()
                pro_info = quo_num + '-' + pro_num + '-' + pro_name
                if update_asana_result == 'Success':
                    message(f'Project {pro_info} \nAsana Updated Successfully', self.parent.ui)
                elif update_asana_result == 'Fail':
                    message(f'Project {pro_info} \nUpdate Asana failed, please contact admin', self.parent.ui)
                else:
                    message(update_asana_result, self.parent.ui)
        except:
            traceback.print_exc()
    '''5- Confirm fee acceptance'''
    def f_button_finan_5_confirm(self):
        try:
            accounting_dir = os.path.join(conf["accounting_dir"], self.ui.text_proinfo_1_quonum.text())
            resource_dir = os.path.join(conf["resource_dir"], "txt", "Verbal Fee Acceptance.txt")
            state=self.ui.combobox_pro_1_state.currentText()
            if state in ['Set Up','Gen Fee Proposal','Email Fee Proposal']:
                message("You need to send the Email to Client first",self.parent.ui)
                return
            elif state in ['Design','DWG drawings','Done','Installation','Construction Phase']:
                update=messagebox('Fee acceptance','Fee Already Accepted, do you want to update again?', self.parent.ui)
                if not update:
                    return
            update=messagebox('Confirm','Do you want to verbal confirm?', self.parent.ui)
            if update:
                self.ui.combobox_pro_1_state.setCurrentText('Design')
                self.ui.button_finan_5_confirm.setStyleSheet(conf["Gray button style"])
                shutil.copy(resource_dir, accounting_dir)
                verbal_content=self.ui.text_finan_5_note1.text()+';'+self.ui.text_finan_5_note2.text()+';'+self.ui.text_finan_5_note3.text()
                with open(os.path.join(accounting_dir, "Verbal Fee Acceptance.txt"), "w") as f:
                    f.write(verbal_content)
                f.close()
                self.datanow["Fee accepted date"] = date.today().strftime("%Y-%m-%d")
                self.f_add_user_history('Log fee accept file')
                update_asana = messagebox('Update Asana', "Verbal Fee Acceptance logged. \n Do you want to update Asana?", self.parent.ui)
                if update_asana:
                    update_asana_result=self.f_button_pro_func_updateasana(True)
                    quo_num = self.ui.text_proinfo_1_quonum.text()
                    pro_num = self.ui.text_proinfo_1_pronum.text()
                    pro_name = self.ui.text_proinfo_1_proname.text()
                    pro_info = quo_num + '-' + pro_num + '-' + pro_name
                    if update_asana_result=='Success':
                        message(f'Project {pro_info} \nAsana Updated Successfully',self.parent.ui)
                    elif update_asana_result=='Fail':
                        message(f'Project {pro_info} \nUpdate Asana failed, please contact admin', self.parent.ui)
                    else:
                        message(update_asana_result, self.parent.ui)
        except:
            traceback.print_exc()

    '''=============================Table Calculation==============================='''
    '''1.3 - Table calculation minor'''
    def f_table_proinfo_3_1_area_total(self):
        try:
            total_sum = 0
            self.ui.table_proinfo_3_1_area.itemChanged.disconnect(self.f_table_proinfo_3_1_area_total)
            for row in range(self.ui.table_proinfo_3_1_area.rowCount()-1):
                item = self.ui.table_proinfo_3_1_area.item(row, 2)
                if item:
                    if is_float(item.text()):
                        total_sum += float(item.text())
            self.ui.text_proinfo_3_1_area.setText(str(total_sum)+' m2')
            self.ui.text_fee_5_minor_area.setText(str(total_sum))
            self.ui.table_proinfo_3_1_area.setItem(5, 2, QTableWidgetItem(str(total_sum)))
            self.f_set_table_style(self.ui.table_proinfo_3_1_area)
            self.ui.table_proinfo_3_1_area.itemChanged.connect(self.f_table_proinfo_3_1_area_total)
        except:
            traceback.print_exc()
    '''1.3 - Table calculation major'''
    def f_table_proinfo_3_2_carpark_area_total(self,table_type):
        try:
            proinfo_major_table_dict={"Car park":(self.ui.table_proinfo_3_2_carpark,self.table_proinfo_3_2_carpark_func),
                                      "Apt":(self.ui.table_proinfo_3_2_area,self.table_proinfo_3_2_area_func)}
            table=proinfo_major_table_dict[table_type][0]
            func=proinfo_major_table_dict[table_type][1]
            table.itemChanged.disconnect(func)
            for i in range(2,table.columnCount()-1):
                total_sum = 0
                for row in range(table.rowCount()-2):
                    item = table.item(row, i)
                    if item:
                        if is_float(item.text()):
                            total_sum += int(item.text())
                i_conv=1 if i % 2 == 0 else 0
                table.setItem(table.rowCount()-1-i_conv, i, QTableWidgetItem(str(total_sum)))
                for row in range(table.rowCount()):
                    table_sum_row=0
                    for col in range(2,table.columnCount()-1):
                        item = table.item(row, col)
                        if item:
                            if is_float(item.text()):
                                table_sum_row += int(item.text())
                    table.setItem(row, table.columnCount()-1, QTableWidgetItem(str(table_sum_row)))
                    if row==table.rowCount()-2:
                        if table_type=="Car park":
                            sum_carpark=str(table_sum_row)
                        else:
                            sum_unit=str(table_sum_row)
            if table_type=="Car park":
                level_list=[]
                for row in range(table.rowCount()-2):
                    item = table.item(row, 0)
                    try:
                        if item.text()!='':
                            level_list.append(item.text())
                    except:
                        pass
                self.ui.table_fee_5_cal_carpark.setItem(0, 0, QTableWidgetItem(str(len(level_list))))
                level_text=''
                if level_list!=[]:
                    for level in level_list:
                        level_text+=level+', '
                self.ui.text_proinfo_3_2_carspot.setText(level_text[:-2]+' '+sum_carpark+' Car Spots')
                self.ui.table_fee_5_cal_carpark.setItem(0, 1, QTableWidgetItem(sum_carpark))
                self.f_set_table_style(self.ui.table_fee_5_cal_carpark)
            if table_type=="Apt":
                self.ui.text_fee_5_major_apt.setText(sum_unit)
                if self.ui.text_proinfo_3_2_area.text()=='':
                    self.ui.text_proinfo_3_2_area.setText(sum_unit+' Units')
                else:
                    try:
                        old_message=self.ui.text_proinfo_3_2_area.text()
                        if old_message.find(' Units')!=-1:
                            loc = old_message.find(' Units')
                            unit_message=old_message[loc + 6:]
                            new_message=sum_unit+' Units'+unit_message
                            self.ui.text_proinfo_3_2_area.setText(new_message)
                        else:
                            self.ui.text_proinfo_3_2_area.setText(sum_unit + ' Units')
                    except:
                        traceback.print_exc()
            self.f_set_table_style(table)
            table.itemChanged.connect(func)
        except:
            traceback.print_exc()
    '''2.6 - Carpark calculation'''
    def f_table_fee_5_cal_carpark_cal(self):
        try:
            table=self.ui.table_fee_5_cal_carpark
            table.itemChanged.disconnect(self.f_table_fee_5_cal_carpark_cal)
            for row in range(table.rowCount()):
                item_1 = table.item(row, 0)
                item_2 = table.item(row, 1)
                if item_1 and item_2:
                    if is_float(item_1.text()) and is_float(item_2.text()):
                        level=float(item_1.text())
                        level_factor=0.25*level+0.75
                        car_factor=float(item_2.text())/90
                        complex_factor=level_factor*car_factor
                        table.setItem(row, 2, QTableWidgetItem(str(round(level_factor,2))))
                        table.setItem(row, 3, QTableWidgetItem(str(round(car_factor, 2))))
                        table.setItem(row, 4, QTableWidgetItem(str(round(complex_factor, 2))))
                        if complex_factor<=0.5:
                            cost = 1000
                        elif complex_factor<=1:
                            cost = 2000
                        elif complex_factor<=5:
                            cost = 3000
                        elif complex_factor<=10:
                            cost = 4000
                        elif complex_factor<=15:
                            cost = 5000
                        else:
                            cost = 6000
                        table.setItem(row, 5, QTableWidgetItem(str(cost)))
            self.f_set_table_style(table)
            table.itemChanged.connect(self.f_table_fee_5_cal_carpark_cal)
        except:
            traceback.print_exc()
    '''2.6 - Table calculation'''
    def f_fee_4_total(self,table_type):
        try:
            table=self.table_fee_4_cal_dict[table_type][0]
            func = self.table_fee_4_cal_dict[table_type][1]
            table.itemChanged.disconnect(func)
        except:
            pass
        try:
            for row in range(table.rowCount() - 1):
                item = table.item(row, 1)
                try:
                    if is_float(item.text()):
                        in_gst=round(float(item.text())*1.1,1)
                        table.setItem(row, 2, QTableWidgetItem(str(in_gst)))
                        fee_float=round(float(item.text()),1)
                        table.setItem(row, 1, QTableWidgetItem(str(fee_float)))
                    elif item.text()=='':
                        table.setItem(row, 2, QTableWidgetItem(''))
                    else:
                        message('Cant insert string in fee table, already delete it', self.parent.ui)
                        table.setItem(row, 1, QTableWidgetItem(''))
                except:
                    pass
            for i in [1,2]:
                total_sum = 0
                for row in range(table.rowCount() - 1):
                    item = table.item(row, i)
                    try:
                        if is_float(item.text()):
                            total_sum += float(item.text())
                    except:
                        pass
                table.setItem(table.rowCount() - 1, i, QTableWidgetItem(str(total_sum)))
            service_list = self.f_get_pro_service_type('Frame_with_var')
            table_fee_list = []
            for service_full_name in service_list:
                service = conf['Service short dict'][service_full_name]
                table_fee_list.append(self.ui_name_dict["table_fee_4_" + service])
            ex_sum,in_sum=0,0
            for i in range(len(table_fee_list)):
                try:
                    ex_sum+=round(float(table_fee_list[i].item(4, 1).text()),1)
                    in_sum+=round(float(table_fee_list[i].item(4, 2).text()),1)
                except:
                    pass
            self.ui.text_fee_4_ex.setText(str(ex_sum))
            self.ui.text_fee_4_in.setText(str(in_sum))
            self.ui.text_finan_1_total0_1.setText(str(ex_sum))
            self.ui.text_finan_1_total0_2.setText(str(in_sum))
            finan_doc_list=[self.ui_name_dict["text_finan_1_" + table_type +'_doc'+str(i+1)] for i in range(4)]
            fee_1_list=[self.ui_name_dict["text_finan_1_" + table_type +'_fee'+str(i+1)+'_1'] for i in range(4)]
            fee_2_list=[self.ui_name_dict["text_finan_1_" + table_type + '_fee' + str(i + 1) + '_2'] for i in range(4)]
            fee_sum_1=self.ui_name_dict["text_finan_1_" + table_type +'_fee0_1']
            fee_sum_2=self.ui_name_dict["text_finan_1_" + table_type + '_fee0_2']
            for row in range(table.rowCount() - 1):
                item1 = ''
                item2 = ''
                item3 = ''
                try:
                    item1 = table.item(row, 0).text()
                    item2 = table.item(row, 1).text()
                    item3 = table.item(row, 2).text()
                except:
                    break
                finan_doc_list[row].setText(item1)
                if item1!='' and item2!='':
                    fee_clicked=self.f_fee_clicked(table_type,row)
                    print(table_type,row,fee_clicked)
                    if fee_clicked:
                        if row == 0 or row == 2:
                            finan_doc_list[row].setStyleSheet("""QLineEdit {font: 12pt "Calibri";color: rgb(0, 0, 0);border-style: ridge;
                                                border-color: rgb(100, 100, 100);border-width: 1px;background-color: rgb(240, 240, 240);}""")
                        else:
                            finan_doc_list[row].setStyleSheet("""QLineEdit {font: 12pt "Calibri";color: rgb(0, 0, 0);border-style: ridge;
                                                                        border-color: rgb(100, 100, 100);border-width: 1px;background-color: rgb(220, 220, 220);}""")
                    else:
                        finan_doc_list[row].setStyleSheet("""QLineEdit {font: 12pt "Calibri";color: rgb(0, 0, 0);border-style: ridge;
                        border-color: rgb(100, 100, 100);border-width: 1px;background-color: rgb(220, 0, 0);}""")

                fee_1_list[row].setText(item2)
                fee_2_list[row].setText(item3)
            if is_float(table.item(table.rowCount()-1, 1).text()) and is_float(table.item(table.rowCount()-1, 2).text()):
                fee_sum_1.setText(str(table.item(table.rowCount()-1, 1).text()))
                fee_sum_2.setText(str(table.item(table.rowCount()-1, 2).text()))
        except:
            traceback.print_exc()
        finally:
            self.f_set_table_style(table)
            table.itemChanged.connect(func)

    def f_fee_clicked(self,service,row):
        try:
            fee_clicked = False
            for j in range(8):
                print("radbutton_finan_1_" + service + '_' + str(row + 1) + '_' + str(j + 1))
                radiobutton_name = self.ui_name_dict["radbutton_finan_1_" + service + '_' + str(row + 1) + '_' + str(j + 1)]
                if radiobutton_name.isChecked():
                    fee_clicked = True
                    break
            print(fee_clicked)
            return fee_clicked
        except:
            traceback.print_exc()
            return False


# if __name__ == '__main__':
#     multiprocessing.freeze_support()
#     app = QApplication(sys.argv)
#     ui=uic.loadUi("Copilot-Bridge.ui")
#     stats = Stats_bridge(ui)
#     app.aboutToQuit.connect(stats.on_close)
#     stats.ui.show()
#     sys.exit(app.exec_())

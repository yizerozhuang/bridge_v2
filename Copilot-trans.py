import json
import os.path
import traceback
from pathlib import Path
import time
import pyautogui
from conf.config_copilot import CONFIGURATION as conf
from utility.pdf_utility import init_environment,flatten_pdf,open_in_bluebeam,read_json,is_int,unflatten_pdf
init_environment(conf)
import keyboard
import shutil
import ast
from datetime import datetime
from utility.pdf_tool import PDFTools as PDFTools_v2
PDFTools_v2.SetEnvironment(conf["bluebeam_dir"], conf["bluebeam_engine_dir"], r"C:\Progra~1\Inkscape\bin\inkscape.exe", conf["c_temp_dir"])
import subprocess




position_left_top_corner = (356, 190)
position_right_top_corner = (1438, 190)
position_left_bottom_corner = (356, 952)
position_right_bottom_corner = (1438, 952)

position_doc_1, position_doc_2 = (210, 168), (270, 168)
position_left2_close,position_right_close=(240, 168),(950, 168)
position_close = (1906, 13)
position_left, position_right = (500, 500), (1350, 500)
position_full_page = (296, 986)
point_boundary_left = (193, 328), (872, 808)
point_boundary_right = (922,328), (1588,805)
point_boundary_right_4 =[(922,328),(1588,328),(1588,805),(922,805)]
# point_boundary_right_4 =[(922,328),(1618,328),(1618,805),(922,805)]
position_opacity=(1750,500)
position_sync=(1590,1023)
position_add2layer = [(480, 500), (543, 650), (771, 647), (780, 674)]

position_erase_color=(455, 100)
position_colorprocess=(23,105)
position_color_screen=(850,280)
position_grayscale=(850,324)
position_colorize=(850,310)
position_selectcolor=(852,324)
position_color_red=(860,396)
position_color_cyan=(950,418)
position_color_green=(925,376)
position_color_orange=(880,374)
position_color_purple=(990,396)
position_color_blue=(970,374)
position_color_dark_green=(927,376)
position_color_dark_orange=(880,374)
position_color_processimage=(1028,344)
position_color_ok=(1110,776)


first_point= (860, 357)
y_sampling_point=(850, 365)
step=22
modify_color_background=(32, 32, 32)
position_colorto = (910, 328)
position_colortotransparent = (917, 468)

position_flatten_logo=(845,51)
position_flatten_selectedmarkup=(770,477)
position_flatten_button=(1086,764)

position_changeumi=(942,352)
position_addlumi=(984,317)
position_color_close=(1197,245)

position_layer=(149,167)
position_addlayer=(12,193)
position_addlayer_fromfirstlayer=[(65,221),(119,350),(327,350)]
position_flatten_fisrtlayer=(125,494)

position_arrow=(713,986)
position_pagescale=(789,984)
position_piccolorchange=(1699,426)
position_lefttoclear=((685,920),(210,220))
position_righttoclear=((1560,930),(1100,220))
position_order = (506, 106)

position_blend=(1758,523)
position_darken=(1773,590)

keyboard_save='ctrl+s'
keyboard_chooseall='ctrl+a'
keyboard_copy='ctrl+c'
keyboard_past2sameloc='ctrl+shift+v'
keyboard_delete='delete'
keyboard_snapshot='tab+g'
keyboard_enter='enter'
keyboard_esc='esc'
keyboard_division='ctrl+2'
keyboard_tofirstpage='Home'
keyboard_tonextpage='ctrl+right'

pdf_dir=conf["trans_timewait_dir"]
sleep_dict = read_json(pdf_dir)


def insert_word(word):
    for ch in word:
        keyboard.press_and_release(ch)
        time.sleep(0.1)
def change2click(onoff_list):
    num = [i for i, x in enumerate(onoff_list) if x == True]
    click_list=[num[0]]
    for j in range(len(num)-1):
        click_list.append(num[j+1]-num[j])
    return click_list

def sleep(name,page=8):
    time_tosleep=sleep_dict[name]
    if page>8 and name in ['time_open_FileB','time_division','time_change2FileA','time_change2fisrtpage','time_sync','time_nextpage',
                'time_save','time_fullpage','time_screenshot','time_paste2sameposition','time_flattenpic','time_nextpage','time_changecolor',
                'time_changeopcacity','time_add2layer','time_flattenlayer','time_changepiccolor']:
        time_tosleep=int(time_tosleep*page/8)
    time.sleep(time_tosleep)


class Simulator:
    def __init__(self):
        pass
    def click_at(self,position):
        x, y = position
        pyautogui.moveTo(x, y)
        time.sleep(0.5)
        pyautogui.click()
    def right_click_at(self,position):
        x, y = position
        pyautogui.moveTo(x, y)
        time.sleep(0.5)
        pyautogui.rightClick()
    def move_to(self,position):
        x, y = position
        pyautogui.moveTo(x, y)
        time.sleep(0.5)

    def drag(self,position):
        time.sleep(0.5)
        pyautogui.moveTo(position[0][0], position[0][1])
        pyautogui.mouseDown()
        pyautogui.moveTo(position[1][0], position[1][1], duration=1)
        pyautogui.mouseUp()
        time.sleep(0.5)

    @staticmethod
    def copy_markup(folder_dir):
        try:
            with open(os.path.join(folder_dir, 'file_names_copymarkup.txt'), 'r', encoding='utf-8') as file:
                file_content = file.read()
                content = ast.literal_eval(file_content)
                file_a_dir = content['File1']
                file_b_dir = content['File2']
                copyornot_from_list = content['copyornot'][0]
                copyornot_to_list = content['copyornot'][1]
            click_list1 = change2click(copyornot_from_list)
            click_list2 = change2click(copyornot_to_list)
            print(file_a_dir)
            print(file_b_dir)
            print(copyornot_from_list)
            print(copyornot_to_list)
            print(click_list1)
            print(click_list2)
            folder_path = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_a = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name,'File A.pdf')
            file_b = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File B.pdf')
            file_ss = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File SS.pdf')
            shutil.copy(file_a_dir, file_a)
            shutil.copy(file_b_dir, file_b)
            shutil.copy(file_b_dir, file_ss)
            time.sleep(3)
            page_num_sleep = max(PDFTools_v2.page_count(file_a), PDFTools_v2.page_count(file_b))
            open_in_bluebeam(file_a)
            sleep('time_open_FileA',page_num_sleep)
            open_in_bluebeam(file_b)
            sleep('time_open_FileB',page_num_sleep)
            sim = Simulator()
            keyboard.press_and_release(keyboard_division)
            sleep('time_division',page_num_sleep)
            sim.click_at(position_doc_1)
            sleep('time_change2FileA',page_num_sleep)
            keyboard.press_and_release(keyboard_tofirstpage)
            sleep('time_change2fisrtpage',page_num_sleep)
            sim.click_at(position_sync)
            sleep('time_sync',page_num_sleep)
            for i in range(len(click_list1)):
                sim.click_at(position_left)
                sleep('time_normalwait',page_num_sleep)
                try:
                    for j in range(click_list1[i]):
                        keyboard.press_and_release(keyboard_tonextpage)
                        sleep('time_nextpage',page_num_sleep)
                except:
                    pass
                sim.click_at(position_right)
                sleep('time_normalwait',page_num_sleep)
                try:
                    for j in range(click_list2[i]):
                        keyboard.press_and_release(keyboard_tonextpage)
                        sleep('time_nextpage',page_num_sleep)
                except:
                    pass
                sim.click_at(position_left)
                sleep('time_normalwait',page_num_sleep)
                keyboard.press_and_release(keyboard_chooseall)
                keyboard.press_and_release(keyboard_copy)
                sleep('time_normalwait',page_num_sleep)
                sim.click_at(position_right)
                sleep('time_normalwait',page_num_sleep)
                keyboard.press_and_release(keyboard_past2sameloc)
                sleep('time_normalwait',page_num_sleep)
            keyboard.press_and_release(keyboard_save)
            sleep('time_save',page_num_sleep)
            sim.click_at(position_sync)
            sleep('time_sync',page_num_sleep)
            sim.click_at(position_close)
            shutil.copy(file_b,file_b_dir)
            sleep('time_normalwait',page_num_sleep)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()


    @staticmethod
    def get_rect(file_path, two_point_position, page_number=1, color=None):
        (x1, y1), (x2, y2) = two_point_position
        dct = PDFTools_v2.return_markup_by_page(file_path, page_number)
        w, h = PDFTools_v2.page_size(file_path, page_number - 1)
        rect = [(k, v) for k, v in dct.items() if v.get("subject") == 'Rectangle']
        if color is not None:
            rect = [(k, v) for k, v in rect if v.get('color') == color]
        assert len(rect) == 1
        rect = rect[0][1]
        x, y, width, height, c = float(rect['x']), float(rect['y']), float(rect['width']), float(rect['height']), rect['color']
        a, b, c, d = (x, y), (x + width, y), (x + width, y + height), (x, y + height)
        a1, b1, c1, d1 = [(int(x1 + xi / w * (x2 - x1)), int(y1 + yi / h * (y2 - y1))) for xi, yi in [a, b, c, d]]
        return a1, b1, c1, d1


    @staticmethod
    def setup_drawing(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_setupdrawing.txt'), 'r', encoding='utf-8') as file:
                content = file.read()
                parts = content.split(';')
                file_a_dir = parts[0].strip()
                file_b_dir = parts[1].strip()
            print(file_a_dir,file_b_dir)
            folder_path = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_a = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File A.pdf')
            file_b = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File B.pdf')
            shutil.copy(file_a_dir, file_a)
            shutil.copy(file_b_dir, file_b)
            pages = PDFTools_v2.page_count(file_a)
            points_list=[]
            for i in range(pages):
                point_four = Simulator.get_rect(file_a, point_boundary_left, page_number=i+1, color='#7C0000')
                points_list.append(point_four)
            print(points_list)
            sleep('time_normalwait',pages)
            open_in_bluebeam(file_a)
            sleep('time_open_FileA',pages)
            open_in_bluebeam(file_b)
            sleep('time_open_FileB',pages)
            sim = Simulator()
            keyboard.press_and_release(keyboard_division)
            sleep('time_division',pages)
            sim.click_at(position_doc_1)
            sleep('time_change2FileA',pages)
            keyboard.press_and_release(keyboard_tofirstpage)
            sleep('time_change2FileA',pages)
            sim.click_at(position_full_page)
            sleep('time_fullpage',pages)
            for i in range(pages):
                points = points_list[i]
                sim.click_at(position_left)
                keyboard.press_and_release(keyboard_chooseall)
                sleep('time_normalwait',pages)
                keyboard.press_and_release(keyboard_delete)
                sleep('time_normalwait',pages)
                keyboard.press_and_release(keyboard_snapshot)
                sleep('time_normalwait',pages)
                for position_snap in points:
                    sim.click_at(position_snap)
                sleep('time_normalwait',pages)
                keyboard.press_and_release(keyboard_enter)
                sleep('time_screenshot',pages)
                keyboard.press_and_release(keyboard_esc)
                sleep('time_normalwait',pages)
                sim.click_at(position_right)
                keyboard.press_and_release(keyboard_past2sameloc)
                sleep('time_paste2sameposition',pages)
                sim.click_at(position_flatten_logo)
                sleep('time_normalwait',pages)
                sim.click_at(position_flatten_selectedmarkup)
                sleep('time_normalwait',pages)
                sim.click_at(position_flatten_button)
                sleep('time_flattenpic',pages)
                if i < pages - 1:
                    keyboard.press_and_release(keyboard_tonextpage)
                    sleep('time_nextpage',pages)
                else:
                    keyboard.press_and_release(keyboard_save)
                    sleep('time_save',pages)
                    sim.click_at(position_close)
                    sleep('time_normalwait',pages)
                    keyboard.press_and_release('n')
            shutil.copy(file_b,file_b_dir)
            sleep('time_normalwait',pages)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()


    @staticmethod
    def overlay(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_overlay.txt'), 'r', encoding='utf-8') as file:
                file_content = file.read()
                content = ast.literal_eval(file_content)
            file_a_dir = content['File1']
            file_b_dir = content['File2']
            file_c_dir = content['File3']
            file_d_dir = content['File4']
            color=content['Color']
            layer_name=content['Name']
            service_type=layer_name
            folder_path = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_a = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File A.pdf')
            file_b = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File B.pdf')
            shutil.copy(file_a_dir, file_a)
            shutil.copy(file_b_dir, file_b)
            pages = PDFTools_v2.page_count(file_a)
            print(pages)
            flatten_pdf(file_b, file_b)
            open_in_bluebeam(file_a)
            sleep('time_open_FileA',pages)
            open_in_bluebeam(file_b)
            sleep('time_open_FileB',pages)
            sim = Simulator()
            keyboard.press_and_release(keyboard_division)
            sleep('time_division',pages)
            sim.click_at(position_doc_1)
            sleep('time_change2FileA',pages)
            keyboard.press_and_release(keyboard_tofirstpage)
            sleep('time_change2fisrtpage',pages)
            sim.click_at(position_full_page)
            sleep('time_fullpage',pages)
            if color!=None:
                sim.click_at(position_right)
                sim.click_at(position_colorprocess)
                sleep('time_normalwait',pages)
                sim.click_at(position_color_screen)
                sim.click_at(position_colorize)
                sim.click_at(position_selectcolor)
                if color=='#FF0000':
                    sim.click_at(position_color_red)
                elif color == '#00FFFF':
                    sim.click_at(position_color_cyan)
                elif color == '#008000':
                    sim.click_at(position_color_green)
                elif color == '#FFAA00':
                    sim.click_at(position_color_orange)
                elif color == '#AA00FF':
                    sim.click_at(position_color_purple)
                elif color == '#0000FF':
                    sim.click_at(position_color_blue)
                elif color == '#808000':
                    sim.click_at(position_color_dark_green)
                elif color == '#FF6600':
                    sim.click_at(position_color_dark_orange)
                sim.click_at(position_color_processimage)
                sim.click_at(position_color_ok)
                sleep('time_changecolor',pages)
            sim.click_at(position_left)
            sim.click_at(position_layer)
            sim.click_at(position_addlayer)
            insert_word('None-'+str(datetime.now().strftime("%Y%m%d%H%M%S")))
            sleep('time_normalwait',pages)
            keyboard.press_and_release(keyboard_enter)
            sleep('time_normalwait',pages)
            sim.right_click_at(position_addlayer_fromfirstlayer[0])
            sleep('time_normalwait',pages)
            sim.move_to(position_addlayer_fromfirstlayer[1])
            sleep('time_normalwait',pages)
            sim.click_at(position_addlayer_fromfirstlayer[2])
            sleep('time_normalwait',pages)
            insert_word(service_type)
            sleep('time_normalwait',pages)
            keyboard.press_and_release(keyboard_enter)
            sleep('time_normalwait',pages)
            for i in range(pages):
                sim.click_at(position_right)
                keyboard.press_and_release(keyboard_snapshot)
                sleep('time_normalwait',pages)
                for j in range(4):
                    sim.click_at(point_boundary_right_4[j])
                sleep('time_normalwait',pages)
                keyboard.press_and_release(keyboard_enter)
                sleep('time_screenshot',pages)
                keyboard.press_and_release(keyboard_esc)
                sleep('time_normalwait',pages)
                sim.click_at(position_left)
                keyboard.press_and_release(keyboard_past2sameloc)
                sleep('time_paste2sameposition',pages)
                sim.click_at(position_opacity)
                keyboard.press_and_release('delete')
                sleep('time_normalwait',pages)
                keyboard.press_and_release('delete')
                sleep('time_normalwait',pages)
                keyboard.press_and_release('delete')
                sleep('time_normalwait',pages)
                keyboard.press_and_release('5')
                sleep('time_normalwait',pages)
                keyboard.press_and_release('0')
                sleep('time_normalwait',pages)
                keyboard.press_and_release(keyboard_enter)
                sleep('time_changeopcacity',pages)

                sim.click_at(position_blend)
                time.sleep(0.5)
                sim.click_at(position_darken)
                time.sleep(0.5)

                sim.right_click_at(position_add2layer[0])
                sleep('time_normalwait',pages)
                sim.move_to(position_add2layer[1])
                sleep('time_normalwait',pages)
                sim.move_to(position_add2layer[2])
                sleep('time_normalwait',pages)
                sim.click_at(position_add2layer[3])
                sleep('time_add2layer',pages)
                if color == None:
                    keyboard.press_and_release(keyboard_esc)
                    sleep('time_normalwait', pages)
                    sim.click_at(position_left)
                    sleep('time_normalwait', pages)
                    sim.click_at(position_order)
                    sleep('time_normalwait', pages)


                if i < pages - 1:
                    keyboard.press_and_release(keyboard_tonextpage)
                    sleep('time_nextpage',pages)
                else:
                    sim.click_at(position_left)
                    keyboard.press_and_release(keyboard_save)
                    sleep('time_save',pages)
                    sim.click_at(position_right)
                    keyboard.press_and_release(keyboard_save)
                    sleep('time_save',pages)
                    sim.click_at(position_close)
            sleep('time_normalwait',pages)
            shutil.copy(file_a,file_c_dir)
            sleep('time_normalwait',pages)
            if file_d_dir!='':
                open_in_bluebeam(file_a)
                sleep('time_open_FileB', pages)
                keyboard.press_and_release(keyboard_tofirstpage)
                sleep('time_change2fisrtpage', pages)
                sim.click_at(position_full_page)
                sleep('time_fullpage', pages)
                for i in range(pages):
                    sim.click_at(position_left)
                    sleep('time_normalwait', pages)
                    sim.click_at(position_piccolorchange)
                    sim.click_at(position_color_screen)
                    sim.click_at(position_grayscale)
                    sim.click_at(position_color_processimage)
                    keyboard.press_and_release(keyboard_enter)
                    sleep('time_changeopcacity',pages)
                    if i < pages - 1:
                        sim.click_at(position_left)
                        keyboard.press_and_release(keyboard_tonextpage)
                        sleep('time_nextpage',pages)
                    else:
                        sim.click_at(position_left)
                        keyboard.press_and_release(keyboard_save)
                        sleep('time_save',pages)
                        sim.click_at(position_close)
                sleep('time_normalwait',pages)
                shutil.copy(file_a,file_d_dir)
                sleep('time_normalwait',pages)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()

    @staticmethod
    def lumicolor(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_lumicolor.txt'), 'r', encoding='utf-8') as file:
                file_content = file.read()
                content = ast.literal_eval(file_content)
            lumionoff=content['lumionoff']
            changecoloronoff=content['changecoloronoff']
            sketch=content['file']
            color_change_small=content['color_change']
            lumi_set=content['lumi_set']
            folder_path = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_sketch_pc = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File A.pdf')
            shutil.copy(sketch, file_sketch_pc)
            pages = PDFTools_v2.page_count(file_sketch_pc)
            sleep('time_normalwait',pages)
            sleep('time_normalwait', pages)
            open_in_bluebeam(file_sketch_pc)
            sleep('time_open_FileB',pages)
            sim = Simulator()
            keyboard.press_and_release(keyboard_tofirstpage)
            sleep('time_change2fisrtpage',pages)
            print(color_change_small)
            color_change=[item.upper() for item in color_change_small]
            if changecoloronoff:
                for i in range(pages):
                    sim.click_at(position_colorprocess)
                    sleep('time_normalwait',pages)
                    sim.click_at(position_selectcolor)
                    color_dict=get_color()
                    color_dict_keys=list(color_dict.keys())
                    print(color_dict_keys)
                    sim.click_at(position_color_close)
                    sleep('time_normalwait',pages)
                    for j in range(len(color_change)):
                        if color_change[j] in color_dict_keys:
                            sim.click_at(position_colorprocess)
                            sleep('time_normalwait',pages)
                            sim.click_at(position_selectcolor)
                            color_dict2 = get_color()
                            sleep('time_normalwait',pages)
                            position_color=color_dict2[color_change[j]]
                            sim.click_at(position_color)
                            sleep('time_normalwait',pages)
                            sim.click_at(position_colorto)
                            sleep('time_normalwait',pages)
                            sim.click_at(position_colortotransparent)
                            sleep('time_normalwait',pages)
                            sim.click_at(position_color_ok)
                            sleep('time_changepiccolor',pages)
                    sim.click_at(position_color_ok)
                    if i < pages - 1:
                        keyboard.press_and_release(keyboard_tonextpage)
                        sleep('time_nextpage',pages)
                    else:
                        if lumionoff:
                            sim.click_at(position_colorprocess)
                            sim.click_at(position_color_screen)
                            sim.click_at(position_changeumi)
                            lumi_click=int(float(lumi_set)/0.05)
                            for k in range(lumi_click):
                                sim.click_at(position_addlumi)
                            sim.click_at(position_color_processimage)
                            sim.click_at(position_color_ok)
                            sleep('time_changecolor',pages)
                        sim.click_at(position_left)
                        keyboard.press_and_release(keyboard_save)
                        sleep('time_save',pages)
                        sim.click_at(position_close)
            else:
                if lumionoff:
                    sim.click_at(position_colorprocess)
                    sim.click_at(position_color_screen)
                    sim.click_at(position_changeumi)
                    lumi_click = int(float(lumi_set) / 0.05)
                    for k in range(lumi_click):
                        sim.click_at(position_addlumi)
                    sim.click_at(position_color_processimage)
                    sim.click_at(position_color_ok)
                    sleep('time_changecolor',pages)
                sim.click_at(position_left)
                keyboard.press_and_release(keyboard_save)
                sleep('time_save',pages)
                sim.click_at(position_close)
            sleep('time_normalwait',pages)
            shutil.copy(file_sketch_pc,sketch)
            sleep('time_normalwait',pages)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()

    @staticmethod
    def plot(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_plot.txt'), 'r', encoding='utf-8') as file:
                file_content = file.read()
                content = ast.literal_eval(file_content)
            sketch=content['File1']
            papersize=content['Papersize']
            folder_path = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_sketch_pc = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File A.pdf')
            shutil.copy(sketch, file_sketch_pc)
            pages = PDFTools_v2.page_count(file_sketch_pc)
            unflatten_pdf(file_sketch_pc,file_sketch_pc)
            sleep('time_normalwait', pages)
            open_in_bluebeam(file_sketch_pc)
            sleep('time_open_FileB',pages)
            sim = Simulator()
            keyboard.press_and_release(keyboard_tofirstpage)
            sleep('time_change2fisrtpage',pages)
            sim.click_at(position_full_page)
            sleep('time_fullpage', pages)
            sim.click_at(position_pagescale)
            sleep('time_normalwait', pages)
            if papersize=='A3':
                scale='25'
            elif papersize=='A1':
                scale='12.5'
            elif papersize=='A0':
                scale='8.84'
            else:
                scale='6.25'
            insert_word(scale)
            sleep('time_normalwait', pages)
            keyboard.press_and_release(keyboard_enter)
            sleep('time_fullpage', pages)
            sim.click_at(position_arrow)
            for i in range(pages):
                sim.drag(position_lefttoclear)
                keyboard.press_and_release(keyboard_delete)
                sleep('time_normalwait',pages)
                sim.drag(position_righttoclear)
                keyboard.press_and_release(keyboard_delete)
                sleep('time_normalwait',pages)
                if i < pages - 1:
                    keyboard.press_and_release(keyboard_tonextpage)
                    sleep('time_nextpage',pages)
                else:
                    sim.click_at(position_left)
                    keyboard.press_and_release(keyboard_save)
                    sleep('time_save',pages)
                    sim.click_at(position_close)
            sleep('time_normalwait',pages)

            cancel_task_txt = conf['cancel_task_txt']
            with open(cancel_task_txt, 'r', encoding='utf-8') as file:
                content = file.read()
                result_list = content.split(",")
                print(result_list)
                if not Path(folder_dir).name in result_list:
                    shutil.copy(file_sketch_pc,sketch)
            sleep('time_normalwait',pages)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()




    @staticmethod
    def remove(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_remove.txt'), 'r', encoding='utf-8') as file:
                file_content = file.read()
                content = ast.literal_eval(file_content)
            sketch=content['File1']
            open_in_bluebeam(sketch)
            sleep('time_normalwait', 1)
            sleep('time_normalwait', 1)
            sim = Simulator()
            sim.click_at(position_left)
            keyboard.press_and_release(keyboard_save)
            sleep('time_normalwait', 1)
            sleep('time_normalwait', 1)
            sim.click_at(position_close)
            sleep('time_normalwait',1)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()



    @staticmethod
    def stitch(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_stitch.txt'), 'r', encoding='utf-8') as file:
                file_content = file.read()
                content = ast.literal_eval(file_content)
            file_list=content['Files']
            file_o=file_list[0]
            '''加空白页'''
            file_other=file_list[1:]
            pages = PDFTools_v2.page_count(file_o)
            open_in_bluebeam(file_o)
            sleep('time_open_FileB',pages)
            for file_i in file_other:
                points_list = []
                for k in range(pages):
                    point_four = Simulator.get_rect(file_i, point_boundary_right, page_number=k + 1, color='#7C0000')
                    points_list.append(point_four)
                sleep('time_normalwait',pages)
                open_in_bluebeam(file_i)
                sleep('time_open_FileB',pages)
                sim = Simulator()
                keyboard.press_and_release(keyboard_division)
                sleep('division',pages)
                keyboard.press_and_release(keyboard_tofirstpage)
                sleep('time_change2fisrtpage',pages)
                sim.click_at(position_full_page)
                sleep('time_fullpage',pages)
                for i in range(pages):
                    points = points_list[i]
                    sim.click_at(position_right)
                    keyboard.press_and_release(keyboard_chooseall)
                    keyboard.press_and_release(keyboard_delete)
                    sleep('time_normalwait',pages)
                    keyboard.press_and_release(keyboard_snapshot)
                    sleep('time_normalwait',pages)
                    for position_snap in points:
                        sim.click_at(position_snap)
                    sleep('time_normalwait',pages)
                    keyboard.press_and_release(keyboard_enter)
                    sleep('time_screenshot',pages)
                    keyboard.press_and_release(keyboard_esc)
                    sleep('time_normalwait',pages)
                    sim.click_at(position_left)
                    keyboard.press_and_release(keyboard_past2sameloc)
                    sleep('time_paste2sameposition',pages)
                    if i < pages - 1:
                        keyboard.press_and_release(keyboard_tonextpage)
                        sleep('time_nextpage',pages)
                    else:
                        sim.click_at(position_right_close)
                        sleep('time_normalwait',pages)
                        sim.click_at(position_left2_close)
                        sleep('time_normalwait',pages)
                        keyboard.press_and_release('n')
                sim.click_at(position_left)
                keyboard.press_and_release(keyboard_save)
                sleep('time_save',pages)
                sim.click_at(position_close)
                sleep('time_normalwait',pages)
                os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()

    @staticmethod
    def erase_content(folder_dir):
        try:
            with open(os.path.join(folder_dir,'file_names_erase_content.json'), 'r') as f:
                data = json.load(f)
            input_dir = data["input_dir"]
            output_dir = data["output_dir"]
            folder_path = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_a = os.path.join(r'C:\Copilot_tmp', Path(folder_dir).name, 'File A.pdf')
            shutil.copy(input_dir, file_a)
            pages = PDFTools_v2.page_count(file_a)
            points_list=[]
            for i in range(pages):
                point_four = Simulator.get_rect(file_a, (position_left_top_corner, position_right_bottom_corner), page_number=i+1, color='#7C0000')
                points_list.append(point_four)
            print(points_list)
            sleep('time_normalwait',pages)
            open_in_bluebeam(file_a)
            sleep('time_open_FileA',pages)
            sim = Simulator()
            keyboard.press_and_release(keyboard_tofirstpage)
            sleep('time_change2FileA',pages)
            sim.click_at(position_full_page)
            sleep('time_fullpage',pages)
            for i in range(pages):
                points = points_list[i]
                keyboard.press_and_release(keyboard_chooseall)
                sleep('time_normalwait',pages)
                keyboard.press_and_release(keyboard_delete)
                sleep('time_normalwait',pages)
                sim.click_at(position_erase_color)
                sleep('time_erase_color', pages)
                erase_content_part_1 = [position_left_top_corner, (points[3][0], position_left_top_corner[1]), points[3],
                                        (position_right_bottom_corner[0], points[3][1]), position_right_bottom_corner,
                                        position_left_bottom_corner, position_left_top_corner]
                for position_erase in erase_content_part_1:
                    sim.click_at(position_erase)
                time.sleep(5)
                erase_content_part_2 = [(points[3][0], position_left_top_corner[1]), position_right_top_corner,
                                        (position_right_top_corner[0], points[2][1]), points[2], points[1],
                                        points[0],(points[3][0], position_left_top_corner[1])]
                for position_erase in erase_content_part_2:
                    sim.click_at(position_erase)
                sleep('time_paste2sameposition',pages)
                if i < pages - 1:
                    keyboard.press_and_release(keyboard_tonextpage)
                    sleep('time_nextpage',pages)
                else:
                    keyboard.press_and_release(keyboard_save)
                    sleep('time_save',pages)
                    sim.click_at(position_close)
                    sleep('time_normalwait',pages)
                    keyboard.press_and_release('n')

            shutil.copy(file_a, output_dir)
            sleep('time_normalwait',pages)
            os.rename(folder_dir, folder_dir+'-finished')
        except:
            traceback.print_exc()


def get_color():
    sleep('time_normalwait')
    data = pyautogui.screenshot()
    lines = 0
    while y_sampling_point[1] + step * lines < data.height:
        if data.getpixel((y_sampling_point[0], y_sampling_point[1] + step * lines)) != modify_color_background:
            break
        lines += 1
    lines -= 1
    res = {}
    for i in range(lines):
        for j in range(8):
            xy = (first_point[0] + step * j, first_point[1] + step * i)
            color = data.getpixel(xy)
            if color not in res:
                color_16=rgb_to_hex(color)
                res[color_16] = xy
    return res

def rgb_to_hex(rgb):
    try:
        return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'.upper()
    except:
        return '#000000'



def execute_executable():
    exe_path = r'B:\01.Bridge\asana_xero_daily_update_script.exe'
    while True:
        try:
            result = subprocess.run(exe_path, check=True)
            print("程序执行成功，返回码:", result.returncode)
        except subprocess.CalledProcessError as e:
            print("程序执行失败，错误信息:", e)
        time.sleep(86400)




folder_dir=conf["trans_dir"]
while True:
    try:
        folder_last=0
        folders=[f.name for f in Path(folder_dir).iterdir() if f.is_dir()]
        if len(folders)>0:
            for folder in folders:
                if is_int(folder):
                    print(folder)
                    if int(folder)!=int(folder_last):
                        time.sleep(1)
                        if os.path.exists(os.path.join(folder_dir,folder,'file_names_copymarkup.txt')):
                            Simulator.copy_markup(os.path.join(folder_dir,folder))
                        if os.path.exists(os.path.join(folder_dir,folder,'file_names_erase_content.json')):
                            Simulator.erase_content(os.path.join(folder_dir,folder))
                        elif os.path.exists(os.path.join(folder_dir,folder,'file_names_setupdrawing.txt')):
                            Simulator.setup_drawing(os.path.join(folder_dir,folder))
                        elif os.path.exists(os.path.join(folder_dir,folder,'file_names_overlay.txt')):
                            Simulator.overlay(os.path.join(folder_dir, folder))
                        elif os.path.exists(os.path.join(folder_dir,folder,'file_names_lumicolor.txt')):
                            Simulator.lumicolor(os.path.join(folder_dir, folder))
                        elif os.path.exists(os.path.join(folder_dir,folder,'file_names_stitch.txt')):
                            Simulator.stitch(os.path.join(folder_dir, folder))
                        elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_plot.txt')):
                            Simulator.plot(os.path.join(folder_dir, folder))
                        elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_remove.txt')):
                            Simulator.remove(os.path.join(folder_dir, folder))
                    else:
                        os.rename(os.path.join(folder_dir, folder), os.path.join(folder_dir, folder+ '-error'))
                    folder_last=folder

    except:
        traceback.print_exc()
    time.sleep(10)
# if __name__ == '__main__':
#     Simulator.copy_markup(r"B:\02.Copilot\Copilot_file_trans\20240906163226-finished")
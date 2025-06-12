from itertools import product
import zipfile

import PyPDF2
import fitz
from pymupdf import PDF_ENCRYPT_KEEP
from PIL import Image, ImageEnhance
import webbrowser
import _thread
from pypdf import Transformation
import pypdf
import time
import multiprocessing
import traceback
import shutil
import win32file
from pathlib import Path
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
import os
import subprocess
from xml.etree import ElementTree as ET
import json
from utility.pdf_tool import PDFTools as PDFTools_v2
# from conf.conf import CONFIGURATION as conf
from datetime import datetime

conf = {}

def init_environment(config):
    global conf
    conf = config
    PDFTools_v2.SetEnvironment(conf["bluebeam_dir"], conf["bluebeam_engine_dir"], r"C:\Progra~1\Inkscape\bin\inkscape.exe", conf["c_temp_dir"])


def read_json(json_dir):
    return json.load(open(json_dir))


def open_folder(dir):
    webbrowser.open(dir)


def get_file_name(file_path):
    return os.path.basename(file_path)

def move_file_to_ss(input_file, ss_folder):
    os.makedirs(ss_folder, exist_ok=True)
    if os.path.exists(input_file):
        shutil.move(input_file, os.path.join(ss_folder, f"{get_timestamp()}-{get_file_name(input_file)}"))

def open_link_with_edge(link):
    edge_address = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    def open_link():
        subprocess.call([edge_address, link])
    _thread.start_new_thread(open_link, ())


def file_exists(file_dir):
    return os.path.exists(file_dir)

def get_pdf_page_numbers(input_path):
    f = open(input_path, "rb")
    pdf = PyPDF2.PdfReader(f)
    pages = len(pdf.pages)
    f.close()
    return pages

def get_paper_sizes(paper_sizes, paper_types):
    return [(paper_sizes[paper]["width"], paper_sizes[paper]["height"]) for paper in paper_types]

def convert_mm_to_pixel(input_length):
    return round(input_length * conf["mm_to_pixel"], 2)

def get_timestamp():
    return datetime.now().strftime("%Y%m%d%H%M%S")


def is_float(input):
    try:
        if input == None:
            return False
        else:
            float(input)
            return True
    except ValueError:
        return False


def char2num(char):
    return ord(char) - ord('A')


def create_directory(directory):
    if os.path.exists(directory) == False:
        try:
            os.makedirs(directory)
        except Exception as e:
            print(f"Failed to create directory {directory}. Reason: {e}")


def markup_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def remove_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def plot_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def lumicolor_pc_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def setupdrawing_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def aligncombine_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def overlay_pc_wait(file_name):
    i = 0
    while True:
        i += 1
        print(i)
        if find_folder(Path(file_name).parent.absolute(), Path(file_name).name+'-finished'):
            break
        else:
            time.sleep(10)
        if i > 120:
            break


def get_file_bytes(file_path):
    file = Path(file_path)
    if file.exists():
        return file.stat().st_size
    else:
        return None


def getlogo(pdf_path, output_path, paper_size):
    page_number = 0
    img = pdf_page_to_image(pdf_path, page_number)
    if paper_size == 'A3':
        bbox = (660, 746, 763, 828)  # 根据需要调整坐标
    elif paper_size == 'A1':
        bbox = (960, 1527, 1235, 1645)   # 根据需要调整坐标
    else:
        bbox = (1258, 2225, 1870, 2348)  # 根据需要调整坐标
    cropped_img = crop_image(img, bbox)
    save_image(cropped_img, output_path)


def open_in_bluebeam(file_dir):
    try:
        _thread.start_new_thread(open_pdf_in_bluebeam, (file_dir, ))
    except:
        traceback.print_exc()


def is_int(name):
    try:
        int(name)
        return True
    except:
        return False


def resize_png(input_image, new_size):
    img = Image.open(input_image)
    img = img.resize(new_size)
    img.save(input_image)


def crop_image(img, bbox):
    return img.crop(bbox)


def save_image(img, path):
    img.save(path)


def list_all_files(directory):
    file_paths = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_paths.append(file_path)
    return file_paths


def f_waitinline(file_time):
    try:
        folder_dir = conf["trans_dir"]
        folders = [f.name for f in Path(folder_dir).iterdir() if f.is_dir()]
        folder_unfinished = []
        for folder_i in folders:
            if folder_i.isdigit():
                folder_unfinished.append(folder_i)
        folder_unfinished.sort()
        if file_time != '0000':
            if len(folder_unfinished) > 1:
                folder = folder_unfinished[0]
                if folder != file_time:
                    if os.path.exists(os.path.join(folder_dir, folder, 'file_names_copymarkup.txt')):
                        process_unfinished = 'copying markup'
                    elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_setupdrawing.txt')):
                        process_unfinished = 'setting up drawing'
                    elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_overlay.txt')):
                        process_unfinished = 'overlaying'
                    elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_lumicolor.txt')):
                        process_unfinished = 'changing color and luminosity'
                    elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_stitch.txt')):
                        process_unfinished = 'stitching drawings'
                    elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_plot.txt')):
                        process_unfinished = 'plotting'
                    else:
                        process_unfinished = 'doing something unknown'

                    return process_unfinished
                else:
                    return 'ok now'
            else:
                return 'ok now'
        else:
            if folder_unfinished != []:
                folder = folder_unfinished[0]
                if os.path.exists(os.path.join(folder_dir, folder, 'file_names_copymarkup.txt')):
                    process_unfinished = 'copying markup'
                elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_setupdrawing.txt')):
                    process_unfinished = 'setting up drawing'
                elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_overlay.txt')):
                    process_unfinished = 'overlaying'
                elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_lumicolor.txt')):
                    process_unfinished = 'changing color and luminosity'
                elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_stitch.txt')):
                    process_unfinished = 'stitching drawings'
                elif os.path.exists(os.path.join(folder_dir, folder, 'file_names_plot.txt')):
                    process_unfinished = 'plotting'
                else:
                    process_unfinished = 'doing something unknown'

                return process_unfinished
            else:
                return 'ok now'
    except:
        traceback.print_exc()


def get_temp_name(scale, drawing_type, paper_size):
    size_name = paper_size.lower()
    scale_name = scale.split(':')[1]
    folder_name = 'temp-'+size_name+'-'+drawing_type
    file_name1 = 'template 1-'+scale_name+'.pdf'
    file_name2 = 'template 2-'+scale_name+'.pdf'
    file_name3 = 'template 3-'+scale_name+'.pdf'
    temp1 = os.path.join(conf[folder_name], file_name1)
    temp2 = os.path.join(conf[folder_name], file_name2)
    temp3 = os.path.join(conf[folder_name], file_name3)
    return temp1, temp2, temp3


def is_pdf_open(file_path):
    try:
        if os.path.exists(file_path):
            handle = win32file.CreateFile(file_path, win32file.GENERIC_READ, 0, None,
                                          win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
            file_info = win32file.GetFileInformationByHandle(handle)
            win32file.CloseHandle(handle)
        return False
    except Exception as e:
        print(f"Exception occurred: {e}")
        return True


def create_or_clear_directory(directory):
    if os.path.exists(directory):
        # 如果目录存在，清空目录下的所有文件和子目录
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")
    else:
        # 如果目录不存在，创建目录
        try:
            os.makedirs(directory)
        except Exception as e:
            print(f"Failed to create directory {directory}. Reason: {e}")


def get_number_of_page(file_dir):
    try:
        reader = pypdf.PdfReader(file_dir)
        return len(reader.pages)
    except:
        return 0


def find_keyword(filename):
    keyword_list = []
    none_key_list = ['UPDATE', 'DRAWING', 'UPDATED']
    for key in filename.split():
        try:
            parts = key.split('-')
            for part in parts:
                if part.upper() not in none_key_list:
                    keyword_list.append(part)
        except:
            if key.upper() not in none_key_list:
                keyword_list.append(key)
    return keyword_list


def open_pdf_in_bluebeam(file_name):
    try:
        bluebeam_engine_dir = conf["bluebeam_engine_dir"]
        command = f"Open('{file_name}') View() Close()"
        subprocess.check_output([bluebeam_engine_dir, command])
    except:
        time.sleep(2)
        try:
            bluebeam_engine_dir = conf["bluebeam_engine_dir"]
            command = f"Open('{file_name}') View() Close()"
            subprocess.check_output([bluebeam_engine_dir, command])
        except:
            traceback.print_exc()


def pdf_page_to_image(pdf_path, page_number, zoom=1.0):
    pdf_document = fitz.open(pdf_path)
    page = pdf_document.load_page(page_number)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    pdf_document.close()
    return img


def insert_image_into_pdf(pdf_path, image_path, page_number, x, y, width, height, rotation):
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    rect = fitz.Rect(y, x, y+height, x+width)
    page.insert_image(rect, filename=image_path, rotate=rotation, keep_proportion=False)
    doc.save(pdf_path, incremental=True, encryption=PDF_ENCRYPT_KEEP)


def combine_pdf(input_file_list, output_file):
    if len(input_file_list) == 1:
        shutil.copy(input_file_list[0], output_file)
        return
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    combine_pdfs = []
    for file in input_file_list:
        combine_pdfs.append('\'' + file + '\'')
    command = f"Combine({', '.join(combine_pdfs)}) Save('{output_file}') Close()"
    subprocess.check_output([bluebeam_engine_dir, command])


def flatten_pdf(input_file, output_file):
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    command = f"Open('{input_file}') Flatten(false) Save('{output_file}') Close()"
    subprocess.check_output([bluebeam_engine_dir, command])


def unflatten_pdf(input_file, output_file):
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    command = f"Open('{input_file}') Unflatten() Save('{output_file}') Close()"
    subprocess.check_output([bluebeam_engine_dir, command])
def resize_pdf_multiprocessing(pool, input_file, input_scale, input_size_x, input_size_y,
                               output_scale, output_size_x, output_size_y, output_dir):
    template_folder = conf["c_temp_dir"]
    # template_folder =
    shutil.copy(input_file, template_folder)
    input_file2 = os.path.join(template_folder, Path(input_file).name)
    output_dir2 = os.path.join(template_folder, Path(output_dir).name)
    t = time.time()
    output_tmp_dir = os.path.join(os.path.dirname(input_file2), "resize_tmp")
    os.makedirs(output_tmp_dir, exist_ok=True)
    try:
        reader_pages = split_pdf(input_file2, output_tmp_dir)
        if reader_pages <= 1:
            resize_pdf(input_file2, input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y,
                       output_dir2)
        else:
            output_tmp_path = [os.path.join(output_tmp_dir, "{}.pdf".format(page)) for page in range(reader_pages)]
            args = [(path_k, input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y)
                    for path_k in output_tmp_path]
            for _ in pool.imap_unordered(parse_page, args):
                pass
            combine_pdf(output_tmp_path, output_dir2 + ".tmp.pdf")
            if os.path.exists(output_dir2):
                os.remove(output_dir2)
            os.rename(output_dir2 + ".tmp.pdf", output_dir2)
        shutil.copy(output_dir2, output_dir)
    except:
        traceback.print_exc()
    finally:
        shutil.rmtree(output_tmp_dir)
        print("Resize Success:", time.time() - t)

def luminocity_pdf_multiprocessing(pool, input_file, luminocity_value):
    template_folder = conf["c_temp_dir"]
    input_file2 = os.path.join(template_folder, Path(input_file).name)
    shutil.copy(input_file, input_file2)
    t = time.time()
    output_tmp_dir = os.path.join(os.path.dirname(input_file2), "lumi_tmp")
    os.makedirs(output_tmp_dir, exist_ok=True)
    try:
        reader_pages = split_pdf(input_file2, output_tmp_dir)
        output_tmp_path = [os.path.join(output_tmp_dir, "{}.pdf".format(page)) for page in range(reader_pages)]

        args = [(f, f, float(luminocity_value)) for f in output_tmp_path]
        for _ in pool.imap_unordered(parse_lumi_page, args):
            pass
        combine_pdf(output_tmp_path, input_file2 + ".tmp.pdf")
        if os.path.exists(input_file2):
            os.remove(input_file2)
        os.rename(input_file2 + ".tmp.pdf", input_file2)
        shutil.copy(input_file2, input_file)
    except:
        print("Luminocity Failed:", time.time() - t)
        traceback.print_exc()
    finally:
        print("Luminocity Success:", time.time() - t)
        shutil.rmtree(output_tmp_dir)

def grays_scale_pdf_multiprocessing(pool, input_file):
    template_folder = conf["c_temp_dir"]
    input_file2 = os.path.join(template_folder, Path(input_file).name)
    shutil.copy(input_file, input_file2)
    t = time.time()
    output_dir = os.path.join(os.path.dirname(input_file2), "tmp")
    os.makedirs(output_dir, exist_ok=True)

    reader_pages = split_pdf(input_file2, output_dir)
    output_dirs = [os.path.join(output_dir, "{}.pdf".format(page, )) for page in range(reader_pages)]
    if reader_pages <= 1:
        grays_scale_pdf(input_file2)
    else:
        for _ in pool.imap_unordered(grays_scale_pdf, output_dirs):
            pass
        combine_pdf(output_dirs, input_file2 + ".tmp.pdf")
        if os.path.exists(input_file2):
            os.remove(input_file2)
        os.rename(input_file2 + ".tmp.pdf", input_file2)
    shutil.copy(input_file2, input_file)

    print("Grayscale Success:", time.time() - t)
    shutil.rmtree(output_dir)


def align_multiprocessing(pool, sketch_dir):
    template_folder = conf["c_temp_dir"]
    shutil.copy(sketch_dir, template_folder)
    sketch_dir2 = os.path.join(template_folder, Path(sketch_dir).name)
    t = time.time()
    output_dir = os.path.join(os.path.dirname(sketch_dir2), "tmp_align")
    os.makedirs(output_dir, exist_ok=True)
    try:
        # total_pages = get_pdf_page_numbers(sketch_dir2)
        total_pages = get_number_of_page(sketch_dir2)
        if total_pages <= 1:
            return
        coordinate_list = []
        for i in range(total_pages):
            markups = return_makeup_by_page(sketch_dir2, i+1)
            markups = PDFTools_v2.filter_markup_by(markups, {"type": "PolyLine", "color": "#7C0000"})
            assert len(markups) > 0, "There is no polyline with rgb(124,0,0) in the set"
            assert len(markups) == 1, "There is more than one polyline with rgb(124, 0, 0) in the set"
            first_key = list(markups.keys())[0]
            coordinate_list.append({
                "markup_id": first_key,
                "x": float(markups[first_key]["x"]),
                "y": float(markups[first_key]["y"]),
                "width": float(markups[first_key]["width"]),
                "height": float(markups[first_key]["height"])
            }
        )
        first_coordinate = coordinate_list[0]
        reader_pages = split_pdf(sketch_dir2, output_dir)
        output_dirs = [os.path.join(output_dir, "{}.pdf".format(page, )) for page in range(reader_pages)]
        for _ in pool.imap_unordered(_align_page, list(zip(output_dirs, coordinate_list, [first_coordinate] * reader_pages))[1:]):
            pass
        combine_pdf(output_dirs, sketch_dir2 + ".tmp.pdf")
        if os.path.exists(sketch_dir2):
            os.remove(sketch_dir2)
        os.rename(sketch_dir2 + ".tmp.pdf", sketch_dir2)
    except Exception as e:
        print("Align Failed:", time.time() - t)
        traceback.print_exc()
        raise e
    finally:
        print("Align Success:", time.time() - t)
        shutil.copy(sketch_dir2,sketch_dir)
        shutil.rmtree(output_dir)


def add_tags_multiprocessing(pool, sketch_dir, markuptype, x, y, tag, level, width_coe):
    template_folder = conf["c_temp_dir"]
    shutil.copy(sketch_dir, template_folder)
    sketch_dir2 = os.path.join(template_folder, Path(sketch_dir).name)
    t = time.time()
    output_tmp_dir = os.path.join(os.path.dirname(sketch_dir2), "markup_tmp")
    os.makedirs(output_tmp_dir, exist_ok=True)
    try:
        reader_pages = split_pdf(sketch_dir2, output_tmp_dir)
        output_tmp_path = [os.path.join(output_tmp_dir, "{}.pdf".format(page)) for page in range(reader_pages)]
        args = [(i, (f, markuptype, x, y, tag, level, width_coe)) for i, f in enumerate(output_tmp_path)]
        for _ in pool.imap_unordered(_add_tags_single_thread, args):
            pass
        combine_pdf(output_tmp_path, sketch_dir2 + ".tmp.pdf")
        if os.path.exists(sketch_dir2):
            os.remove(sketch_dir2)
        os.rename(sketch_dir2 + ".tmp.pdf", sketch_dir2)
        shutil.copy(sketch_dir2, sketch_dir)
    except:
        print("AddTags Failed:", time.time() - t)
        traceback.print_exc()
    finally:
        print("AddTags Success:", time.time() - t)
        shutil.rmtree(output_tmp_dir)


def add_item_tags(sketch_dir, markuptype, tags):
    try:
        bluebeam_engine_dir = conf["bluebeam_engine_dir"]
        markup_dir = os.path.join(conf["database_dir"], "markup.json")
        markup_json = read_json(markup_dir)
        formated_xml = format_to_xml(markup_json[markuptype])
        total_pages = get_number_of_page(sketch_dir)
        reader = pypdf.PdfReader(sketch_dir)
        markup_paste_list = []
        i = 0
        for tag in tags.keys():
            label_x = tags[tag][0]
            label_y = tags[tag][1]
            markup_paste_list.append(f"MarkupPaste({1}, '{formated_xml}', {label_y}, {label_x})")
            i += 1
        command = f"Open('{sketch_dir}') " + " ".join(markup_paste_list) + " Save() Close()"
        new_markup_result = subprocess.check_output([bluebeam_engine_dir, command])
        new_markup_list = [markup_id.replace("1", "").replace("\r", "").replace("\n", "")
                           for markup_id in new_markup_result.decode("utf-8").split("\r\n1\r\n")]
        i = 0
        for tag in tags.keys():
            bluebeam_markup_set(sketch_dir, 1, new_markup_list[i], {"comment": tag, "height": "10", 'width': '100'})
            i += 1

    except:
        traceback.print_exc()


def find_word_coordinates_in_pdf(pdf_path, word):
    word_list = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            words = page.extract_words()
            for w in words:
                if w['text'].lower().find(word.lower()) != -1:
                    coordinates = {
                        'page': i + 1,
                        'x0': w['x0'],
                        'y0': w['top'],
                        'x1': w['x1'],
                        'y1': w['bottom']
                    }
                    word_list[w['text']] = [int(w["x0"]), int(w["top"])]
    return word_list


def blubeam_reduce_size(file, ratio):
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    command = f"Open('{file}') ReduceFileSize('{ratio}','150', 'true', 'true', 'true')  Save() Close()"
    subprocess.check_output([bluebeam_engine_dir, command])


def extract_first_page(input_pdf, output_pdf):
    try:
        with open(input_pdf, 'rb') as file:
            reader = PdfReader(file)
            writer = PdfWriter()
            page = reader.pages[0]
            writer.add_page(page)
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
    except:
        traceback.print_exc()


def remove_first_page(input_pdf, output_pdf):
    try:
        with open(input_pdf, 'rb') as file:
            reader = PdfReader(file)
            writer = PdfWriter()
            for page_num in range(1, len(reader.pages)):
                page = reader.pages[page_num]
                writer.add_page(page)
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
    except:
        traceback.print_exc()


def remove_pdf_page(input_pdf, output_pdf, page_list):
    try:
        with open(input_pdf, 'rb') as file:
            reader = PdfReader(file)
            writer = PdfWriter()
            print(page_list)
            for page_num in page_list:
                page = reader.pages[page_num-1]
                writer.add_page(page)
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
    except:
        traceback.print_exc()

def adjust_pdf_brightness(input_pdf, output_pdf, brightness_factor):
    # Open the input PDF
    doc = fitz.open(input_pdf)

    # Iterate through each page
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)

        # Iterate through each text instance on the page
        for text_instance in page.get_text("dict")["blocks"]:
            if text_instance["type"] == 0:  # This block contains text
                for line in text_instance["lines"]:
                    for span in line["spans"]:
                        span["color"] = adjust_color_brightness(span["color"], brightness_factor)

        # Iterate through each image on the page
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]

            # Adjust the brightness of the image
            adjusted_image = adjust_image_brightness(image_bytes, brightness_factor)

            # Replace the old image with the adjusted image
            page.insert_image(page.rect, stream=adjusted_image)

    # Save the modified PDF
    doc.save(output_pdf)

def int_to_rgb(color_int):
    r = (color_int >> 16) & 0xFF
    g = (color_int >> 8) & 0xFF
    b = color_int & 0xFF
    return r, g, b

def adjust_color_brightness(color, factor):
    # Extract RGB components
    r, g, b = int_to_rgb(color)
    # Adjust each component and ensure it stays within the 0-255 range
    r = min(255, max(0, int(r * factor)))
    g = min(255, max(0, int(g * factor)))
    b = min(255, max(0, int(b * factor)))
    return (r, g, b)

def adjust_image_brightness(image_bytes, factor):
    from PIL import Image
    from io import BytesIO

    # Open the image from bytes
    image = Image.open(BytesIO(image_bytes))

    # Convert the image to RGB if it's not
    image = image.convert("RGB")

    # Apply brightness adjustment
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(factor)

    # Save the adjusted image to bytes
    img_byte_arr = BytesIO()
    image.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def get_paper_size(filename, id):
    page_size = PDFTools_v2.page_size(filename, id)
    paper_standard = {"A0": {"width": 1189.0, "height": 841.0},
                      "A1": {"width": 841.0, "height": 594.0},
                      "A2": {"width": 594.0, "height": 420.0},
                      "A3": {"width": 420.0, "height": 297.0},
                      "A4": {"width": 297.0, "height": 210.0},
                      "A0-24": {"width": 1682.0, "height": 1189.0},
                      "A0-26": {"width": 2523.0, "height": 1189.0}}
    width = page_size[0]/2.83463544
    height = page_size[1]/2.83463544
    for k, v in paper_standard.items():
        if abs(width-v['width']) < 10 and abs(height-v['height']) < 10:
            return k
    return str(int(width))+':'+str(int(height))


def svg_set_opacity_multicolor(input_svg, pre_colors, output_path):
    et = PDFTools_v2.get_ET(input_svg, 0)
    for elem in et.iter():
        for pre_color in pre_colors:
            if 'fill' in elem.attrib and (pre_color.lower() in elem.attrib["fill"] or pre_color.upper() in elem.attrib["fill"]):
                elem.attrib["fill-opacity"] = '0'
            if 'stroke' in elem.attrib and (pre_color.lower() in elem.attrib["stroke"] or pre_color.upper() in elem.attrib["stroke"]):
                elem.attrib["stroke-opacity"] = '0'
    with open(output_path, 'wb') as f:
        root = ET.ElementTree(et)
        root.write(f)


def pdf_move_center_per_page(pdf_file, paper_size, output_page):
    if paper_size == 'A3':
        center_point = [13.91296, 14.40692, 1161.803, 724.2948]
    elif paper_size == 'A1':
        center_point = [51.49623, 39.79395, 2270.485, 1476.798]
    else:
        center_point = [50.58789, 39.68091, 3257.938, 2177.293]
    for i in range(1, PDFTools_v2.page_count(pdf_file)+1):
        markups = PDFTools_v2.return_markup_by_page(pdf_file, i)
        markups = PDFTools_v2.filter_markup_by(markups, {"subject": "Rectangle", "color": "#7C0000"})
        assert len(markups) > 0, f"Page {i} don't have any rectangular"
        assert len(markups) == 1, f"Page {i} have more than one rectangular"

    for i in range(1, PDFTools_v2.page_count(pdf_file) + 1):
        pass
def pdf_move_center(pdf_file, paper_size):
    if paper_size == 'A3':
        center_point = [13.91296, 14.40692, 1161.803, 724.2948]
    elif paper_size == 'A1':
        center_point = [51.49623, 39.79395, 2270.485, 1476.798]
    else:
        center_point = [50.58789, 39.68091, 3257.938, 2177.293]


    markups = PDFTools_v2.return_markup_by_page(pdf_file, 1)
    #retangular
    markups = PDFTools_v2.filter_markup_by(markups, {"subject":"Rectangle", "color": "#7C0000"})
    assert len(markups) == 1, "every pages should have only one rectangle"
    markup_rect = list(markups.items())[0]
    coordinate = (float(markup_rect[1]['x']) + float(markup_rect[1]['width']) / 2,
                  float(markup_rect[1]['y']) + float(markup_rect[1]['height']) / 2)
    center_coordinate = (center_point[0] + center_point[2] / 2, center_point[1] + center_point[3] / 2)
    offset = (center_coordinate[0] - coordinate[0], -center_coordinate[1] + coordinate[1])
    # move all markups
    for i in range(PDFTools_v2.page_count(pdf_file)):
        markups = PDFTools_v2.return_markup_by_page(pdf_file, i + 1)
        markups = PDFTools_v2.filter_markup_by(markups, {"color": "#7C0000"})
        for markup_name, markup_item in markups.items():
            PDFTools_v2.set_markup(
                pdf_file, i + 1, {markup_name: {"x": str(float(markup_item["x"]) + offset[0]), "y": str(float(markup_item["y"]) - offset[1])}})
    pool = multiprocessing.Pool(20)
    align_page_tmp = os.path.join(PDFTools_v2.TEMP_PATH, "align_page_tmp")
    os.makedirs(align_page_tmp, exist_ok=True)
    output_files = PDFTools_v2.split_pdf(pdf_file, align_page_tmp)
    output_files_2 = [item + ".output.pdf" for item in output_files]
    for _ in pool.starmap(PDFTools_v2.pdf_content_move, [(output_file, output_file_2, offset) for output_file, output_file_2 in zip(output_files, output_files_2)]):
        pass
    PDFTools_v2.combine_pdf(output_files_2, pdf_file)
    shutil.rmtree(align_page_tmp)


def insert_logo_into_pdf(pdf_path, image_path, page_number, paper_size):
    #TODO: need to keep the same rotation and make sure the template is no rotation for all
    if paper_size == 'A3':
        rec_x = 665
        rec_y = 16
        rec_width = 95
        rec_height = 75
    elif paper_size == 'A1':
        rec_x = 100
        rec_y = 0
        rec_width = 100
        rec_height = 200
    else:
        #A0
        rec_x = 1265
        rec_y = 45
        rec_width = 595
        rec_height = 110
    img = Image.open(image_path)
    img_width, img_height = img.size
    # img_x, img_y, img_width, img_height = get_logo_position(rec_x, rec_y, rec_width, rec_height, img_width, img_height)
    # insert_image_into_pdf(pdf_path, image_path, page_number, img_x, img_y, img_width, img_height, rotation=0)
    insert_image_into_pdf(pdf_path, image_path, page_number, rec_x, rec_y,rec_width, rec_height, rotation=270)

def remove_duplicates_keep_order(lst):
    seen = set()
    return [x for x in lst if not (x in seen or seen.add(x))]

def generate_possible_colors(colors_list):
    result = []
    for color in colors_list:
        result.append(color.upper())
        result+=adjust_hex_color(color, 1)
    return remove_duplicates_keep_order(result)

def version_setter(pdf_file, page, papersize):
    UPLIFT_SIZE = -11.6
    CONTEXT = {"A0": {"VERSION": {"$TYPE": "while-list", "$UPLIFT_SIZE": (0, UPLIFT_SIZE), "REV": [49.91895, 2318.24], "REVISION DESCRIPTION": [85.23486, 2318.24], "DRW": [227.1118, 2318.24], "CHK": [264.9587, 2318.24], "DATE Y.M.D": [304.6389, 2318.24]}, "PROJECT": [2272.446, 2232.563], "SCALE": [3053.2, 2235.87], "PROJECT No.": [3053.2, 2279.353], "DRAWING No.": [3053.2, 2320.773], "TYPE OF ISSUE": [3138.28, 2320.818], "DRAWN BY": [3138.28, 2236.499], "CHECKED BY": [3138.28, 2279.191]},
               "A1": {"VERSION": {"$TYPE": "while-list", "$UPLIFT_SIZE": (0, UPLIFT_SIZE), "REV": [51.47995, 1617.341], "REVISION DESCRIPTION": [86.79601, 1617.341], "DRW": [228.6728, 1617.341], "CHK": [266.5199, 1617.341], "DATE Y.M.D": [306.1994, 1617.341]}, "PROJECT": [1524.801, 1533.359], "SCALE": [2067.0, 1535.052], "PROJECT No.": [2067.0, 1578.535], "DRAWING No.": [2067.0, 1619.956], "TYPE OF ISSUE": [2152.08, 1620.0], "DRAWN BY": [2152.08, 1535.682], "CHECKED BY": [2152.08, 1578.373]},
               "A3": {"VERSION": {"$TYPE": "while-list", "$UPLIFT_SIZE": (0, UPLIFT_SIZE), "REV": [16.01697, 808.8484], "REVISION DESCRIPTION": [34.50598, 808.8484], "DRW": [142.6079, 808.8484], "CHK": [160.26, 808.8484], "DATE Y.M.D": [181.502, 808.8484]}, "PROJECT": [869.0056, 747.4893], "SCALE": [1075.596, 747.387], "PROJECT No.": [1079.934, 771.3923], "DRAWING No.": [1084.874, 794.3511], "TYPE OF ISSUE": [1084.786, 815.5005], "DRAWN BY": [1141.957, 748.9429], "CHECKED BY": [1143.425, 771.4201]}}
    TOL = 6
    markups = PDFTools_v2.return_markup_by_page(pdf_file, page)
    structured_context = PDFTools_v2.get_structured_markups_from(markups, CONTEXT[papersize], TOL)
    return structured_context


def get_frame_content(context, content1, content2, pos):
    try:
        if content1 != '':
            result = context[content1][pos][content2][1][2]
        else:
            result = context[content2][1][2]
    except:
        result = '/'
    return result


def getpage(pdf_path, output_path, page_number):
    img = pdf_page_to_image(pdf_path, page_number)
    save_image(img, output_path)


def resize_pdf(input_file, input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y, output_dir):
    reader = pypdf.PdfReader(input_file)
    writer = pypdf.PdfWriter()
    page_scale_x = float(output_size_x) / float(input_size_x)
    page_scale_y = float(output_size_y) / float(input_size_y)
    content_scale_x = (float(input_scale) / float(output_scale)) / page_scale_x
    content_scale_y = (float(input_scale) / float(output_scale)) / page_scale_y
    print(content_scale_x, content_scale_y)
    op = Transformation().scale(content_scale_x, content_scale_y)
    for page in reader.pages:
        if "/VP" in page.keys():
            page.pop("/VP")
        if page_scale_x != 1 or page_scale_y != 1:
            page.scale(page_scale_x, page_scale_y)

        if content_scale_x != 1 or content_scale_y != 1:
            page.add_transformation(op)
        writer.add_page(page)
    writer.write(output_dir)


def replace_string(string, replacement):
    for key, value in replacement.items():
        string = string.replace(key, value)
    return string


def return_makeup_by_page(file_dir, i):
    command = f"Open('{file_dir}')  MarkupGetExList({i}) Close()"
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]

    result_bytes = subprocess.check_output([bluebeam_engine_dir, command])
    result_text = result_bytes.decode("utf-8").split("\n")[1]
    replacement = {
        "|'": "'",
        "'{": "{",
        "}'": "}",
        "||": "\\",
        "\r":"",
        "'True'": "True",
        "'False'": "False",
        "'None'": "None"
    }
    string = replace_string(result_text, replacement)
    if len(string)==0:
        return {}
    result_json = eval(string)
    return result_json

def grays_scale_pdf(sketch_dir):
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    command = f"Open('{sketch_dir}')  ColorProcess('black','white') Close(true)"
    subprocess.check_output([bluebeam_engine_dir, command])


def _align_page(args):
    try:
        input_file, coordinate_i, first_coordinate = args
        writer = pypdf.PdfWriter()
        reader = pypdf.PdfReader(input_file)
        page = reader.pages[0]
        translation_x = first_coordinate["x"] - coordinate_i["x"]
        translation_y = - (first_coordinate["y"] - coordinate_i["y"])
        assert page.rotation in [0, 90, 180, 270]
        if page.rotation == 0:
            op = Transformation().translate(translation_x, translation_y)
        elif page.rotation == 90:
            op = Transformation().translate(- translation_y, translation_x)
        elif page.rotation == 180:
            op = Transformation().translate(- translation_x, - translation_y)
        else:
            op = Transformation().translate(translation_y, - translation_x)
        page.add_transformation(op)
        writer.add_page(page)
        writer.write(input_file)

        comments = '{"x":"%.2f","y":"%.2f"}' % (first_coordinate["x"], first_coordinate["y"])
        print(comments)
        commands = [
            "Open('{}')".format(input_file),
            "MarkupSet(1, '{}', '{}')".format(coordinate_i["markup_id"], comments),
            "Save()",
            "Close()",
        ]
        subprocess.check_output([conf["bluebeam_engine_dir"], " ".join(commands)])
    except:
        traceback.print_exc()


def split_pdf(input_pdf, output_dir):
    with open(input_pdf, 'rb') as file:
        reader = pypdf.PdfReader(file)
        if len(reader.pages) == 1:
            shutil.copy(input_pdf, os.path.join(output_dir, '0.pdf'))
            return len(reader.pages)
        for i in range(len(reader.pages)):
            writer = pypdf.PdfWriter()
            writer.add_page(reader.pages[i])
            output_pdf = os.path.join(output_dir, '{}.pdf'.format(i))
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
        return len(reader.pages)


def parse_page(r):
    input_file, input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y = r
    resize_pdf(input_file, input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y, input_file)

def parse_lumi_page(r):
    input_file, output_file, lumi_value = r
    PDFTools_v2.adjust_luminocity(input_file, output_file, lumi_value)

def format_to_xml(input_json):
    res = '<?xml version="1.0" encoding="us-ascii"?>'
    res += '<MarkupCopyItem>'
    res += '<Type>'
    res += input_json["Type"]
    res += '</Type>'
    res += '<Raw>'
    res += input_json["Raw"]
    res += '</Raw>'
    res += '</MarkupCopyItem>'
    return res


def format_markup_input(input_data):
    return "{" + ','.join([f'"{key}":"{value}"' for key, value in input_data.items()])+"}"


def bluebeam_markup_set(sketch_dir, page, markup_id, input_data):
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    input_data = format_markup_input(input_data)
    command = f"Open('{sketch_dir}') MarkupSet({page}, '{markup_id}', '{input_data}')  Save() Close()"
    subprocess.check_output([bluebeam_engine_dir, command])


def get_tag_width(tag, level, markup_width_coe):
    cha_num = max(len(tag), len(level))
    coe = 1
    return cha_num*float(markup_width_coe)*coe


def _add_tags_single_thread(args):
    i, (sketch_dir, markuptype, x, y, tag, level, width_coe) = args
    bluebeam_engine_dir = conf["bluebeam_engine_dir"]
    markup_dir = os.path.join(conf["database_dir"], "markup.json")
    markup_json = read_json(markup_dir)
    formated_xml = format_to_xml(markup_json[markuptype])
    reader = pypdf.PdfReader(sketch_dir)
    markup_paste_list = []
    page = reader.pages[0]
    label_x = page.mediabox.width * x
    label_y = page.mediabox.height * y
    assert page.rotation in [0, 90, 180, 270]
    if page.rotation == 0:
        markup_paste_list.append(f"MarkupPaste(1, '{formated_xml}', {label_x}, {label_y})")
    elif page.rotation == 90:
        markup_paste_list.append(f"MarkupPaste(1, '{formated_xml}', {page.mediabox.width-label_y}, {label_x})")
    elif page.rotation == 180:
        markup_paste_list.append(f"MarkupPaste(1, '{formated_xml}', {page.mediabox.width-label_x}, {page.mediabox.height-label_y})")
    else:
        markup_paste_list.append(f"MarkupPaste(1, '{formated_xml}', {label_y}, {page.mediabox.height-label_x})")
    command = f"Open('{sketch_dir}') " + " ".join(markup_paste_list) + " Save() Close()"
    new_markup_result = subprocess.check_output([bluebeam_engine_dir, command])
    new_markup_list = [markup_id.replace("1", "").replace("\r", "").replace("\n", "")
                       for markup_id in new_markup_result.decode("utf-8").split("\r\n1\r\n")]
    width = get_tag_width(tag, level, width_coe)
    bluebeam_markup_set(sketch_dir, 1, new_markup_list[0], {"comment": level[i].upper()+"\n"+tag, "height": "10", "width": width})


def get_pdf_page_sizes(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    size = (page.rect.width, page.rect.height)
    doc.close()
    return size


def find_folder(root_dir, target_folder_name):
    for root, dirs, files in os.walk(root_dir):
        if target_folder_name in dirs:
            return os.path.join(root, target_folder_name)
    return None


def get_logo_position(rec_x, rec_y, rec_width, rec_height, img_width, img_height):
    rec_ratio = rec_width / rec_height
    img_ratio = img_width / img_height
    if rec_ratio > img_ratio:
        img_width = int(img_width * (rec_height / img_height))
        img_height = rec_height
        img_x = int(rec_x + (rec_width - img_width) / 2)
        img_y = rec_y
    else:
        img_height = int(img_height * (rec_width / img_width))
        img_width = rec_width
        img_x = rec_x
        img_y = int(rec_y + (rec_height - img_height) / 2)
    return img_x, img_y, img_width, img_height


def parse_page_ranges(pages_cus):
    pages_list = []
    ranges = pages_cus.split(',')
    for r in ranges:
        r = r.strip()
        if '-' in r:
            start, end = r.split('-')
            start = int(start.strip())
            end = int(end.strip())
            pages_list.extend(range(start, end + 1))
        else:
            pages_list.append(int(r.strip()))

    return sorted(set(pages_list))


def compress_directory_to_zip(zip_filename, dir_to_zip):
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(dir_to_zip):
            for file in files:
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, dir_to_zip)
                zipf.write(file_path, arcname=relative_path)


def page_delete(input_pdf, output_pdf, n):
    try:
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        n = min(n, len(reader.pages))
        for i in range(n):
            writer.add_page(reader.pages[i])
        with open(output_pdf, 'wb') as f:
            writer.write(f)
    except:
        traceback.print_exc()


def generate_combinations(a, b, c, x):
    delta = list(range(-x, x + 1))  # 生成所有可能的增减值
    possible_a = [max(0, a + d) for d in delta]  # 计算每个数字的所有可能值
    possible_b = [max(0, b + d) for d in delta]
    possible_c = [max(0, c + d) for d in delta]
    possible_a = [min(255, val) for val in possible_a]  # 确保所有可能值都在0到255的范围内
    possible_b = [min(255, val) for val in possible_b]
    possible_c = [min(255, val) for val in possible_c]
    all_combinations = list(product(possible_a, possible_b, possible_c))  # 生成所有可能的组合
    return all_combinations


def adjust_hex_color(hex_color, tol):
    hex_color = hex_color.lstrip('#')  # 去掉前面的 #
    if len(hex_color) != 6:  # 确保颜色码是 6 位
        raise ValueError("Invalid hex color code")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    combinations = generate_combinations(r, g, b, tol)
    color_16_list = []
    for i in range(len(combinations)):
        ri = combinations[i][0]
        bi = combinations[i][1]
        gi = combinations[i][2]
        color_i = f'#{ri:02X}{gi:02X}{bi:02X}'
        color_16_list.append(color_i)
    return color_16_list

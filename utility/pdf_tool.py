import uuid
import os
import shutil
import subprocess
import logging
import re
import io
import fitz
import svgwrite
from xml.etree import ElementTree as ET
import pypdf
import json
import copy

from PIL import Image, ImageEnhance
import img2pdf

logger = logging.getLogger(__name__)
logging.basicConfig(filename='../myapp.log', level=logging.INFO)

class External(svgwrite.container.Group):
    def __init__(self, xml, **extra):
        self.xml = xml
        ns = u'{http://www.w3.org/2000/svg}'
        nsl = len(ns)
        for elem in self.xml.iter():
            if elem.tag.startswith(ns):
                elem.tag = elem.tag[nsl:]

        super(External, self).__init__(**extra)

    def get_xml(self):
        return self.xml

class PDFTools:
    BLUEBEAM_DIR = r"C:\Program Files\Bluebeam Software\Bluebeam Revu\2017\Revu\Revu.exe"
    BLUEBEAM_ENGINE_DIR = r"C:\Program Files\Bluebeam Software\Bluebeam Revu\2017\Script\ScriptEngine.exe"
    INKSCAPE_DIRECTORY = r"C:\Progra~1\Inkscape\bin\inkscape.exe"
    TEMP_PATH = r"D:\Workspace\repos\pdf_tools"
    @classmethod
    def SetEnvironment(cls, BLUEBEAM_DIR, BLUEBEAM_ENGINE_DIR, INKSCAPE_DIRECTORY, TEMP_PATH):
        cls.BLUEBEAM_DIR = BLUEBEAM_DIR
        cls.BLUEBEAM_ENGINE_DIR = BLUEBEAM_ENGINE_DIR
        cls.INKSCAPE_DIRECTORY = INKSCAPE_DIRECTORY
        cls.TEMP_PATH = TEMP_PATH

    @staticmethod
    def _rename_id(element, suffix=None):
        if suffix is None:
            suffix = str(uuid.uuid4()).split('-')[0]
        for elem in element.iter():
            for key in elem.keys():
                if key in ('id', ):
                    elem.set(key, elem.get(key) + '_' + suffix)
            for key in elem.attrib:
                if re.match(r".*url\(#.*\).*", elem.attrib[key]):
                    elem.attrib[key] = re.sub(r"\(#(.*?)\)", r"(#{}_{})".format(r"\1", suffix), elem.attrib[key])
                if key.endswith("href") and re.match(r"#.*", elem.attrib[key]):
                    elem.attrib[key] = re.sub(r"#(.*)", r"#{}_{}".format(r"\1", suffix), elem.attrib[key])

    @staticmethod
    def pdf_page_to_svg(pdf_path, page_number, svg_file):
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_number)
        svg_image = page.get_svg_image()
        svg_image = svg_image.replace("&", "&amp;")
        svg_file.write(svg_image.encode('ascii'))

    @staticmethod
    def page_count(pdf_path):
        doc = fitz.open(pdf_path)
        return doc.page_count

    @staticmethod
    def page_size(pdf_path, page_id):
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_id)
        return (page.rect.width, page.rect.height)

    @staticmethod
    def _get_svg_bytes_from_pdf(file_name, page_number):
        with io.BytesIO() as write_buffer:
            PDFTools.pdf_page_to_svg(file_name, page_number, write_buffer)
            write_buffer.seek(0)
            svg_file = write_buffer.read().decode('ascii')
        return svg_file

    @staticmethod
    def get_ET(file_name, page_number):
        ET.register_namespace('', "http://www.w3.org/2000/svg")
        if file_name.endswith(".pdf"):
            svg_file = PDFTools._get_svg_bytes_from_pdf(file_name, page_number)
        else:
            with open(file_name, 'r') as f:
                svg_file = f.read()
        root = ET.fromstring(svg_file)
        return root

    @staticmethod
    def ET_to_svg(root, svg_file):
        tree = ET.ElementTree(root)
        tree.write(svg_file)

    @staticmethod
    def overlay_page(file_name1, overlay, output_path, page_number=0, page_number_overlay=0, position=(10, 10)):
        background_root = PDFTools.get_ET(file_name1, page_number)
        overlay_root = PDFTools.get_ET(overlay, page_number_overlay)
        # avoid duplication of namespace
        PDFTools._rename_id(overlay_root)

        # same size as background
        background_width = background_root.attrib.get('width', '100%')
        background_height = background_root.attrib.get('height', '100%')
        dwg = svgwrite.Drawing(output_path, size=(background_width, background_height))

        background_group = dwg.g(id='background')
        background_group.add(External(background_root))
        dwg.add(background_group)

        overlay_group = dwg.g(id='overlay', transform='translate(%f, %f)' % (position[0], position[1]))
        overlay_group.add(External(overlay_root))
        dwg.add(overlay_group)

        dwg.save()

    @staticmethod
    def merge_page(et_list, coordinate_list, page_size, output_path):
        for item_root in et_list[1:]:
            PDFTools._rename_id(item_root)
        dwg = svgwrite.Drawing(output_path, size=page_size)
        for i, item in enumerate(et_list):
            group = dwg.g(id="svg_{}".format(i), transform='translate(%f, %f)' % (coordinate_list[i][0], coordinate_list[i][1]))
            group.add(External(item))
            dwg.add(group)
        dwg.save()

    @staticmethod
    def resize_pdf(input_file, input_scale, input_size_x, input_size_y, output_scale, output_size_x, output_size_y, output_dir):
        reader = pypdf.PdfReader(input_file)
        writer = pypdf.PdfWriter()
        page_scale_x = float(output_size_x) / float(input_size_x)
        page_scale_y = float(output_size_y) / float(input_size_y)
        content_scale_x = (float(input_scale) / float(output_scale)) / page_scale_x
        content_scale_y = (float(input_scale) / float(output_scale)) / page_scale_y
        op = pypdf.Transformation().scale(content_scale_x, content_scale_y)
        for page in reader.pages:
            if "/VP" in page.keys():
                page.pop("/VP")
            if page_scale_x != 1 or page_scale_y != 1:
                page.scale(page_scale_x, page_scale_y)

            if content_scale_x != 1 or content_scale_y != 1:
                page.add_transformation(op)
            writer.add_page(page)
        writer.write(output_dir)

    @staticmethod
    def svg_to_pdf(src, dest):
        command = [f'"{src}"', "--export-type=pdf", f'--export-filename="{dest}"']
        os.system(' '.join([f'{PDFTools.INKSCAPE_DIRECTORY}'] + command))

    # Bluebeam
    @staticmethod
    def combine_pdf(input_file_list, output_file):
        if len(input_file_list) == 1:
            shutil.copy(input_file_list[0], output_file)
            return
        combine_pdfs = []
        for file in input_file_list:
            combine_pdfs.append('\'' + file + '\'')

        command = f"Combine({', '.join(combine_pdfs)}) Save('{output_file}') Close()"
        subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, command])

    @staticmethod
    def _replace_control_chars(s):
        def replacer(match):
            char = match.group(0)
            return f'\\u{ord(char):04x}'
        return re.sub(r'[\x00-\x1F\x7F]', replacer, s)

    @staticmethod
    def return_markup_by_page(file_dir, i):
        command = f"Open('{file_dir}') MarkupGetExList({i}) Close()"

        result_bytes = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, command])
        result_text = result_bytes.decode("gbk").split("\r\n")[1]
        string = result_text.replace("|\"", "\"").replace("|'", "'").replace("'{", "{").replace("}'", "}").replace("||", "\\").replace("'True'", "true").replace("'False'", "false").replace("'None'", "null").replace("'", '"')
        if string.strip() == '':
            return {}
        string = PDFTools._replace_control_chars(string)
        result_json = json.loads(string)
        return result_json

    @staticmethod
    def get_markup_in_region(markups, region):
        # region should be in the format of [x, y, width, height]
        markups = {k: v for k, v in markups.items() if region[0] < float(v.get('x', -1)) < region[0] + region[2] and region[1] < float(v.get('y', -1)) < region[1] + region[3]}
        return markups

    @staticmethod
    def get_markup_without(markups, property_dict):
        for prop_k, prop_v in property_dict.items():
            if type(prop_v) != list:
                prop_v = [prop_v]
            markups = {k: v for k, v in markups.items() if not v.get(prop_k) in prop_v}
        return markups
    @staticmethod
    def get_markup_with_content(markups, contents):
        if type(contents) is str:
            contents = [contents]
        contents = [item.lower() for item in contents]
        markups = {k: v for k, v in markups.items() if v.get("comment", '').strip().lower() in contents}
        return markups

    @staticmethod
    def filter_markup_by(markups, property_dict):
        for prop_k, prop_v in property_dict.items():
            if type(prop_v) != list:
                prop_v = [prop_v]
            markups = {k: v for k, v in markups.items() if v.get(prop_k) in prop_v}
        return markups

    @staticmethod
    def copy_markup(file_dir, i, key):
        command = f"Open('{file_dir}') MarkupCopy({i}, '{key}') Close()"
        result_bytes = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, command])
        result_text = result_bytes.decode("utf-8").split("\n")[1].strip()
        return result_text

    @staticmethod
    def paste_markup_single(file_dir, i, format_xml, position):
        command = f"Open('{file_dir}') MarkupPaste({i}, '{format_xml}', {position[0]}, {position[1]}) Save() Close()"
        result_bytes = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, command])
        result_text = result_bytes.decode("utf-8").split("\n")[1].strip()
        return result_text

    @staticmethod
    def get_offset(standard_form, new_file, anchor_color="#7A0000", scale=(28.3463544, 28.3463544)):
        pages = PDFTools.page_count(standard_form)
        assert pages == PDFTools.page_count(new_file), "file page count mismatch"
        offsets = []
        for i in range(1, pages + 1):
            markups_1, markups_2 = PDFTools.return_markup_by_page(standard_form, i), PDFTools.return_markup_by_page(new_file, i)
            markups_1, markups_2 = [markup for markup in markups_1.values() if markup["color"].lower() == anchor_color.lower()], [markup for markup in markups_2.values() if markup["color"].lower() == anchor_color.lower()]
            assert len(markups_1) > 0 and len(markups_2) > 0, "expected markup with the color {}, but found ({}, {})".format(anchor_color, len(markups_1), len(markups_2))
            markup_1, markup_2 = markups_1[0], markups_2[0]
            offset = (float(markup_2['x']) - float(markup_1['x']), float(markup_2['y']) - float(markup_1['y']))
            offset = (offset[0] / scale[0], offset[1] / scale[1])
            offsets.append(offset)
        return offsets

    @staticmethod
    def copy_markup_batch(file_dir, i, dct):
        logger.info("page: {}, copy: {}".format(i, len(dct)))
        command = [f"Open('{file_dir}')", *[f"MarkupCopy({i}, '{key}')" for key in dct.keys()], "Close()"]
        # markup_bci_path = os.path.join(conf["c_temp"])
        with open("copy_markup.bci", 'w') as f:
            f.write('\n'.join(command))
        # result_bytes = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, "Script('copy_markup.bci')"])
        result_bytes = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR]+command)
        result_text = result_bytes.decode("utf-8").split("\r\n")[1::2]
        logger.info("page: {}, copy: {} result".format(i, len(result_text)))
        os.remove("copy_markup.bci")
        return result_text

    @staticmethod
    def paste_markup(file_dir, i, content, content_replace_dict=None):
        # paste
        logger.info("page: {}, paste: {}".format(i, len(content)))
        command = [f"Open('{file_dir}')"]
        for item in content:
            format_xml, position = item
            new_position = (position[1] + position[2], position[0]) if all([position[i] is not None for i in range(3)]) else (0, 0)
            command.append(f"MarkupPaste({i}, '{format_xml}', {new_position[0]}, {new_position[1]})")
        command.extend(["Save()", "Close()"])
        with open("paste_markup.bci", 'w') as f:
            f.write('\n'.join(command))
        markup_ids = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, "Script('paste_markup.bci')"])
        markup_ids_result = [markup_id.strip() for markup_id in markup_ids.decode('utf-8').strip().split('\r\n')]
        line_id, markup_ids = 0, []
        while line_id < len(markup_ids_result):
            k = int(markup_ids_result[line_id])
            markup_ids.append(markup_ids_result[line_id + 1: line_id + 1 + k])
            line_id += 1 + k
        # set
        logger.info("paste_markup: page: {}, set: {}".format(i, len(markup_ids)))
        markup_setter = {}
        for item, markup_id in zip(content, markup_ids):
            if any([item[1][i] is None for i in range(3)]):
                continue
            markup_id = markup_id[0]
            markup_setter[markup_id] = {"x": str(item[1][0]), "y": str(item[1][1])}
            if content_replace_dict is not None and item[1][3] in content_replace_dict:
                markup_setter[markup_id]["comment"] = content_replace_dict[item[1][3]]
        PDFTools.set_markup(file_dir, i, markup_setter)
        os.remove("paste_markup.bci")

    @staticmethod
    def set_markup(file_dir, i, markup_ids_dict):
        command = [f"Open('{file_dir}')"]
        for markup_id, cont in markup_ids_dict.items():
            command.append("MarkupSet({}, '{}', '{}')".format(i, markup_id, json.dumps(cont)))
        command.extend(["Save()", "Close()"])
        logger.info("paste_markup for page {} Complete.".format(i))
        with open("set_markup.bci", 'w') as f:
            f.write('\n'.join(command))
        _ = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, "Script('set_markup.bci')"])
        os.remove("set_markup.bci")

    @staticmethod
    def remove_color(file_dir, colors):
        command = [f"Open('{file_dir}')"]
        for color in colors:
            command.append(f"ColorProcess('{color}','#000000', true)")
        command.extend(["Save()", "Close()"])
        with open("remove_color.bci", 'w') as f:
            f.write('\n'.join(command))
        _ = subprocess.check_output([PDFTools.BLUEBEAM_ENGINE_DIR, "Script('remove_color.bci')"])
        os.remove("remove_color.bci")

    @staticmethod
    def insert_layer(file_dir, layer_name, i):
        doc = fitz.open(file_dir)
        page = doc[i]
        doc.add_ocg(layer_name, True)
        doc.save(file_dir, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

    @staticmethod
    def set_luminocity(file_dir, brightness_factor):
        doc = fitz.open(file_dir)
        temp_images = []
        for page_number in range(len(doc)):
            # Render page to image (as a PIL Image)
            pix = doc.load_page(page_number).get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Adjust brightness
            enhancer = ImageEnhance.Brightness(img)
            bright_img = enhancer.enhance(brightness_factor)

            # Save temporary image
            temp_img_path = f"temp_page_{page_number}.jpg"
            bright_img.save(temp_img_path, "JPEG")
            temp_images.append(temp_img_path)

        with open(file_dir, "wb") as f:
            f.write(img2pdf.convert(temp_images))

        # Cleanup temp images
        for img_path in temp_images:
            os.remove(img_path)
    @staticmethod
    def adjust_luminocity(input_pdf, output_pdf, luminocity=0.45):
        assert -1.0 <= luminocity <= 1.0, "Luminocity must be between -1.0 and 1.0"
        doc = fitz.open(input_pdf)

        for page in doc:
            rect = page.rect  # Get page size

            # Define a white semi-transparent overlay
            overlay = page.new_shape()
            overlay.draw_rect(rect)
            overlay.finish(width=0, color=(1, 1, 1), fill=(1, 1, 1), fill_opacity=luminocity)

            # Apply the overlay to the page
            overlay.commit()

        doc.save(output_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)



    @staticmethod
    def set_replace(file_dir, i, before, after):
        markups = PDFTools.return_markup_by_page(file_dir, i)
        markups = PDFTools.get_markup_with_content(markups, before)
        markup_ids_dict = {k: {"comment": after} for k, _ in markups.items()}
        PDFTools.set_markup(file_dir, i, markup_ids_dict)

    @staticmethod
    def paste_markup_to_file(standard_form, new_file, page_number=1, new_page_number=1, region=None, offset=(0, 0), content_replace_dict=None, exclude_markup_filter=None):
        logger.info("paste_markup_to_file for page {} Start.".format(page_number))
        markups = PDFTools.return_markup_by_page(standard_form, page_number)
        if region is not None:
            markups = PDFTools.get_markup_in_region(markups, region)
        if exclude_markup_filter is not None:
            markups = PDFTools.get_markup_without(markups, exclude_markup_filter)
        if len(markups) == 0:
            logger.error("paste_markup_to_file for page {}: no markup.".format(page_number))
            return
        logger.info("return_markup_by_page for page {} finished.".format(page_number))
        fn = lambda value: (float(value['x']) + offset[0] if 'x' in value else None, float(value['y']) + offset[1] if 'y' in value else None, float(value['height']) if 'height' in value else None, value["comment"] if 'comment' in value else None)

        content = list(zip(PDFTools.copy_markup_batch(standard_form, page_number, markups), [fn(value) for _, value in markups.items()]))
        logger.info("copy_markup for page {} finished.".format(page_number))
        PDFTools.paste_markup(new_file, new_page_number, content, content_replace_dict)

    @staticmethod
    def remove_colors_to_file(input_path, remove_color_list, page_number, output_path):
        logger.info(f"remove color for page {page_number} Start.")
        remove_color_tmp = os.path.join(PDFTools.TEMP_PATH, "remove_color_tmp")
        os.makedirs(remove_color_tmp, exist_ok=True)
        temp_file_path = PDFTools.split_pdf(input_path, remove_color_tmp)[page_number-1]
        PDFTools.remove_color(temp_file_path, remove_color_list)
        shutil.move(temp_file_path, output_path)
        shutil.rmtree(remove_color_tmp)

    @staticmethod
    def paste_all_markup_to_file_by_anchor(standard_form, new_file, anchor_color="#7C0000"):
        pages = PDFTools.page_count(standard_form)
        assert pages == PDFTools.page_count(new_file), "file page count mismatch"
        for i in range(1, pages + 1):
            markups_1, markups_2 = PDFTools.return_markup_by_page(standard_form, i), PDFTools.return_markup_by_page(new_file, i)
            markups_1, markups_2 = [markup for markup in markups_1.values() if markup["color"].lower() == anchor_color.lower()], [markup for markup in markups_2.values() if markup["color"].lower() == anchor_color.lower()]
            assert len(markups_1) > 0 and len(markups_2) > 0, "expected markup with the color {}, but found ({}, {})".format(anchor_color, len(markups_1), len(markups_2))
            markup_1, markup_2 = markups_1[0], markups_2[0]
            offset = (float(markup_2['x']) - float(markup_1['x']), float(markup_2['y']) - float(markup_1['y']))
            PDFTools.paste_markup_to_file(standard_form, new_file, i, i, offset)

    @staticmethod
    def paste_all_markup_to_file_by_base_point(standard_form, new_file, anchor_color="#7C0000"):
        # Assume both sketch is aligned and has at most one mark up on the first page
        pages = PDFTools.page_count(standard_form)
        assert pages == PDFTools.page_count(new_file), "file page count mismatch"
        for i in range(1, pages + 1):
            existing_markups = PDFTools.return_markup_by_page(standard_form, i)
            new_markups = PDFTools.return_markup_by_page(new_file, i)
            existing_markups = PDFTools.filter_markup_by(existing_markups, {"type": "PolyLine", "color": anchor_color})
            new_markups = PDFTools.filter_markup_by(new_markups, {"type": "PolyLine", "color": anchor_color})
            assert len(existing_markups) > 0, f"Page {i} of existing sketch dont have base point"
            assert len(new_markups) > 0, f"Page {i} of new sketch dont have base point"
            assert len(existing_markups) == 1, f"Page {i} of existing sketch have more than one base point"
            assert len(new_markups) == 1, f"Page {i} of existing sketch have more than one base point"

            markup_1 = list(existing_markups.values())[0]
            markup_2 = list(new_markups.values())[0]
            offset = (float(markup_2['x']) - float(markup_1['x']), float(markup_2['y']) - float(markup_1['y']))
            PDFTools.paste_markup_to_file(standard_form, new_file, i, i, offset=offset, exclude_markup_filter={"type": "PolyLine", "color": anchor_color})

    @staticmethod
    def mix_patch(standard_form, file_list, output_path):
        markups = PDFTools.return_markup_by_page(standard_form, 1)
        color_position = {item["color"].upper(): (float(item["x"]), float(item["y"])) for item in markups.values()}
        page_counts = [PDFTools.page_count(f) for f in file_list]
        for i in range(max(page_counts)):
            et_list, coord_list = [], []
            for j in range(len(file_list)):
                if i >= page_counts[j]:
                    continue
                markup_list = list(PDFTools.return_markup_by_page(file_list[j], i + 1).values())
                markup_item = [(markup["color"].upper(), (float(markup["x"]), float(markup["y"]))) for markup in markup_list if markup["color"].upper() in color_position]
                assert len(markup_item) == 1, f"Bad page in {file_list[j]}, page {i}: {markup_item}"
                markup_item = markup_item[0]
                coord_list.append((color_position[markup_item[0]][0] - markup_item[1][0], color_position[markup_item[0]][1] - markup_item[1][1]))
                et_list.append(PDFTools.get_ET(file_list[j], i))
            standard_form_size = PDFTools.page_size(standard_form, 0)
            PDFTools.merge_page(et_list, coord_list, standard_form_size, os.path.join(PDFTools.TEMP_PATH, f"{output_path}.{i}.svg"))
            pdf_path = os.path.join(PDFTools.TEMP_PATH, f"{output_path}.{i}.pdf")
            PDFTools.svg_to_pdf(os.path.join(PDFTools.TEMP_PATH, f"{output_path}.{i}.svg"), pdf_path)
            new_size = PDFTools.page_size(pdf_path, 0)
            # TODO: This should be removed if svg_to_pdf is fixed well
            PDFTools.resize_pdf(pdf_path, '100', new_size[0], new_size[1], 100 / standard_form_size[0] * new_size[0], standard_form_size[0], standard_form_size[1], pdf_path)
            PDFTools.paste_markup_to_file(standard_form, pdf_path)
            if os.path.exists(os.path.join(PDFTools.TEMP_PATH, f"{output_path}.{i}.svg")):
                os.remove(os.path.join(PDFTools.TEMP_PATH, f"{output_path}.{i}.svg"))
        pdf_list = [os.path.join(PDFTools.TEMP_PATH, f"{output_path}.{i}.pdf") for i in range(max(page_counts))]
        PDFTools.combine_pdf(pdf_list, output_path)
        for f in pdf_list:
            if os.path.exists(f):
                os.remove(f)

    @staticmethod
    def pdf_color(page, page_number):
        svg_file = PDFTools._get_svg_bytes_from_pdf(page, page_number)
        hex_color_pattern = re.compile(r'(#[A-Fa-f0-9]{6})')
        final_list =  hex_color_pattern.findall(svg_file)
        return set(final_list)

    @staticmethod
    def pdf_set_color(page, page_numbers, pre_color, post_color, output_path):
        page_count = PDFTools.page_count(page)
        output_file = os.path.basename(output_path)
        svg_list = [os.path.join(PDFTools.TEMP_PATH, f"{output_file}.{i}.svg") for i in range(page_count)]
        pdf_list = [os.path.join(PDFTools.TEMP_PATH, f"{output_file}.{i}.pdf") for i in range(page_count)]
        for page_number in range(page_count):
            svg_file = PDFTools._get_svg_bytes_from_pdf(page, page_number)
            if page_number in page_numbers:
                svg_file = svg_file.replace(pre_color.lower(), post_color).replace(pre_color.upper(), post_color)
            with open(svg_list[page_number], 'w') as f:
                f.write(svg_file)
            PDFTools.svg_to_pdf(svg_list[page_number], pdf_list[page_number])
        PDFTools.combine_pdf(pdf_list, output_path)
        for f in svg_list:
            os.remove(f)
        for f in pdf_list:
            os.remove(f)

    @staticmethod
    def pdf_color_v2(page, page_number):
        et = PDFTools.get_ET(page, page_number)
        result = set()
        for elem in et.iter():
            if 'fill' in elem.attrib:
                if re.match(r'(#[A-Fa-f0-9]{6})', elem.attrib["fill"]):
                    result.add(elem.attrib["fill"])
            if 'stroke' in elem.attrib:
                if re.match(r'(#[A-Fa-f0-9]{6})', elem.attrib["stroke"]):
                    result.add(elem.attrib["stroke"])
        return result

    @staticmethod
    def pdf_set_color_v2(page, page_numbers, pre_color, post_color, output_path, opacity=1.0):
        page_count = PDFTools.page_count(page)
        output_file = os.path.basename(output_path)
        svg_list = [os.path.join(PDFTools.TEMP_PATH, f"{output_file}.{i}.svg") for i in range(page_count)]
        pdf_list = [os.path.join(PDFTools.TEMP_PATH, f"{output_file}.{i}.pdf") for i in range(page_count)]
        for page_number in range(page_count):
            et = PDFTools.get_ET(page, page_number)
            if page_number in page_numbers:
                for elem in et.iter():
                    if 'fill' in elem.attrib and (pre_color.lower() in elem.attrib["fill"] or pre_color.upper() in elem.attrib["fill"]):
                        elem.attrib["fill"] = elem.attrib["fill"].replace(pre_color.lower(), post_color).replace(pre_color.upper(), post_color)
                        if float(opacity) < 1.0:
                            elem.attrib["fill-opacity"] = str(opacity)
                    if 'stroke' in elem.attrib and (pre_color.lower() in elem.attrib["stroke"] or pre_color.upper() in elem.attrib["stroke"]):
                        elem.attrib["stroke"] = elem.attrib["stroke"].replace(pre_color.lower(), post_color).replace(pre_color.upper(), post_color)
                        if float(opacity) < 1.0:
                            elem.attrib["stroke-opacity"] = str(opacity)
            with open(svg_list[page_number], 'wb') as f:
                root = ET.ElementTree(et)
                root.write(f)
            PDFTools.svg_to_pdf(svg_list[page_number], pdf_list[page_number])
        PDFTools.combine_pdf(pdf_list, output_path)
        for f in svg_list:
            os.remove(f)
        for f in pdf_list:
            os.remove(f)

    @staticmethod
    def split_pdf(input_pdf, output_dir):
        with open(input_pdf, 'rb') as file:
            reader = pypdf.PdfReader(file)
            if len(reader.pages) == 1:
                output_list = [os.path.join(output_dir, '0.pdf')]
                shutil.copy(input_pdf, output_list[0])
                return output_list
            for i in range(len(reader.pages)):
                writer = pypdf.PdfWriter()
                writer.add_page(reader.pages[i])
                output_pdf = os.path.join(output_dir, '{}.pdf'.format(i))
                with open(output_pdf, 'wb') as output_file:
                    writer.write(output_file)
            return [os.path.join(output_dir, '{}.pdf'.format(i)) for i in range(len(reader.pages))]

    @staticmethod
    def pdf_content_move(input_file, output_file, offset, pages="all"):
        reader = pypdf.PdfReader(input_file)
        writer = pypdf.PdfWriter()
        page = reader.pages[0]
        offset_x, offset_y = offset
        for page_id in range(len(reader.pages)):
            if pages != 'all' and (type(pages) == 'list' and page_id not in pages):
                writer.add_page(page)
                continue
            assert page.rotation in [0, 90, 180, 270]
            if page.rotation == 0:
                op = pypdf.Transformation().translate(offset_x, offset_y)
            elif page.rotation == 90:
                op = pypdf.Transformation().translate(-offset_y, offset_x)
            elif page.rotation == 180:
                op = pypdf.Transformation().translate(-offset_x, -offset_y)
            else:
                op = pypdf.Transformation().translate(offset_y, -offset_x)
            page.add_transformation(op)
            writer.add_page(page)
        writer.write(output_file)

    @staticmethod
    def get_structured_markups_from(markups, context, tol=6, offset=(0, 0)):
        def inside_with_tol(pos, target, tol, offset):
            pos = (pos[0] + offset[0], pos[1] + offset[1])
            return pos[0] - tol / 2 < target[0] < pos[0] + tol / 2 and pos[1] - tol / 2 < target[1] < pos[1] + tol / 2
        def get_markup_info_with_tol(markups, pos, tol, offset):
            result = []
            for ki, vi in markups.items():
                x, y, comment = float(vi.get('x', 0)), float(vi.get('y', 0)), vi.get('comment', 0)
                if inside_with_tol(pos, (x, y), tol, offset):
                    result.append([ki, [x, y, comment]])
            return result
        result = {}
        for item, v in copy.deepcopy(context).items():
            if item.startswith("$"):
                continue
            elif type(v) == dict and v.get("$TYPE") == "while-list":
                value = []
                up_lift_size = v.pop("$UPLIFT_SIZE")
                v["$TYPE"] = "dict"
                offset_i = offset[:]
                while True:
                    current_value = PDFTools.get_structured_markups_from(markups, v, tol, offset_i)
                    if all(map(lambda x: x is None, current_value.values())):
                        break
                    value.append(current_value)
                    offset_i = (offset_i[0] + up_lift_size[0], offset_i[1] + up_lift_size[1])
            elif type(v) == dict:
                value = PDFTools.get_structured_markups_from(markups, v, tol, offset)
            else:
                value = get_markup_info_with_tol(markups, v, tol, offset)
                value = None if len(value) == 0 else value[0]
            result[item] = value
        return result

    @staticmethod
    def set_structured_markups(file_dir, i, context, up_lift):
        context = copy.deepcopy(context)
        version = context.pop("VERSION")
        assert len(version) > 0, "At least one version Row in the page"
        PDFTools.set_markup(file_dir, i, {k: {"x": str(x), "y": str(y), "comment": comment} for _, (k, (x, y, comment)) in context.items()})
        for j, item in enumerate(version):
            PDFTools.set_markup(file_dir, i, {k: {"x": str(x), "y": str(y), "comment": comment} for _, (k, (x, y, comment)) in item.items() if k not in ("", None)})
            for name, (k, (_, _, comment)) in item.items():
                if k not in ("", None):
                    continue
                markup_content = PDFTools.copy_markup(file_dir, i, version[0][name][0])
                k = PDFTools.paste_markup_single(file_dir, i, markup_content, (0, 0))
                x, y = version[0][name][1][0], version[0][name][1][1]
                PDFTools.set_markup(file_dir, i, {k: {"x": str(x + j * up_lift[0]), "y": str(y + j * up_lift[1]), "comment": comment}})


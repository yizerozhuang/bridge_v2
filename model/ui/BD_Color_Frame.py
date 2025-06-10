from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QComboBox, QFrame, QToolButton, QLabel, QCheckBox, QLineEdit
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtGui import QPainter,QPixmap

import os
import shutil
from conf.conf import CONFIGURATION as conf
from utility.pdf_utility import PDFTools_v2, svg_set_opacity_multicolor

from functools import partial

class BD_Color_Frame(BD_Base_Frame):
    svg_width = 600
    svg_height = 450
    set_button_color = lambda color: f"background-color: {color};width:20;height:20;"

    def __init__(self, app, qt_combo_box_page: QComboBox, qt_frame_from: QFrame, qt_frame_to:QFrame,
                 qt_label_before:QLabel, qt_label_after:QLabel,
                 qt_check_box_luminocity: QCheckBox, qt_line_edit_luminocity: QLineEdit,
                 qt_check_box_grayscale: QCheckBox,qt_check_box_add_tags: QCheckBox):
        super().__init__(app)

        self.combo_box_page = qt_combo_box_page
        self.frame_from = qt_frame_from
        self.frame_to = qt_frame_to
        self.label_before = qt_label_before
        self.label_after = qt_label_after
        self.check_box_luminocity = qt_check_box_luminocity
        self.line_edit_luminocity = qt_line_edit_luminocity
        self.check_box_grayscale = qt_check_box_grayscale
        self.check_box_add_tags = qt_check_box_add_tags
        # self.tool_button_group = QButtonGroup()
        # self.tool_button_group.setExclusive(False)
        self.renderer_before = QSvgRenderer()
        self.renderer_after = QSvgRenderer()

        self.combo_box_page.currentIndexChanged.connect(self.on_index_changed)

    def get_selected_colors(self):
        colors = []
        for button in self.frame_to.findChildren(QToolButton):
            if button.text()=="/":
                colors.append(button.objectName())
        return colors

    def grayscale_is_checked(self):
        return self.check_box_grayscale.isChecked()
    def luminocity_is_checked(self):
        return self.check_box_luminocity.isChecked()
    def get_luminoicity(self):
        return self.line_edit_luminocity.text()
    def add_tags_is_checked(self):
        return self.check_box_add_tags.isChecked()
    def set_total_pages_number(self, pages_number):
        for i in range(1, pages_number+1):
            self.combo_box_page.addItem(f"Page {i}")

    def on_index_changed(self):
        page = int(self.combo_box_page.currentText().split(" ")[1]) - 1
        color_origin_pic = os.path.join(conf["c_temp_dir"], "color_origin_pic.svg")
        with open(color_origin_pic, 'wb') as f:
            PDFTools_v2.pdf_page_to_svg(self.app.sketch_tab.current_sketch_dir, page, f)
        self.load_before_svg(color_origin_pic)
        self.reset_frame(self.frame_from)
        self.reset_frame(self.frame_to)
        self.add_tool_buttons(list(PDFTools_v2.pdf_color(self.app.sketch_tab.current_sketch_dir, page)))
    def load_before_svg(self, svg):
        size = QSize(self.svg_width, self.svg_height)
        pixmap = QPixmap(size)
        pixmap.fill(Qt.white)
        painter = QPainter(pixmap)
        self.renderer_before.load(svg)
        self.renderer_before.render(painter)
        painter.end()
        self.label_before.setPixmap(pixmap)

    def load_after_svg(self, svg):
        size = QSize(self.svg_width, self.svg_height)
        pixmap = QPixmap(size)
        pixmap.fill(Qt.white)
        painter = QPainter(pixmap)
        self.renderer_after.load(svg)
        self.renderer_after.render(painter)
        painter.end()
        self.label_after.setPixmap(pixmap)

    def get_current_page(self):
        return int(self.combo_box_page.currentText().split(" ")[1]) - 1

    def add_tool_buttons(self, colors):
        for color in colors:
            tool_button_from = QToolButton()
            tool_button_from.setStyleSheet(BD_Color_Frame.set_button_color(color))
            self.frame_from.layout().addWidget(tool_button_from)
            tool_button_to = QToolButton()
            tool_button_to.setObjectName(color)
            tool_button_to.setStyleSheet(BD_Color_Frame.set_button_color(color))
            self.frame_to.layout().addWidget(tool_button_to)
            # self.tool_button_group.addButton(tool_button_to)
            tool_button_to.clicked.connect(partial(self.on_tool_button_clicked, tool_button_to))

    def on_tool_button_clicked(self, tool_button):
        if tool_button.text()=="":
            tool_button.setText("/")
            tool_button.setStyleSheet(BD_Color_Frame.set_button_color("rgb(255,255,255)"))
        else:
            tool_button.setText("")
            tool_button.setStyleSheet(BD_Color_Frame.set_button_color(tool_button.objectName()))
        remove_colors = self.get_selected_colors()
        # page = self.get_current_page()

        # PDFTools_v2.remove_colors_to_file(Sketch_Tab.current_sketch_dir, remove_colors, page, r"P:\000000AA-test_project\test.pdf")

        color_origin_pic = os.path.join(conf["c_temp_dir"], "color_origin_pic.svg")
        color_changed_pic = os.path.join(conf["c_temp_dir"], "color_changed_pic.svg")

        if len(remove_colors)!=0:
            shutil.copy(color_origin_pic, color_changed_pic)
            svg_set_opacity_multicolor(color_changed_pic, remove_colors, color_changed_pic)
        else:
            shutil.copy(color_origin_pic,color_changed_pic)
        self.load_after_svg(color_changed_pic)

    def reset_frame(self, frame):
        for button in frame.findChildren(QToolButton):
            button.deleteLater()
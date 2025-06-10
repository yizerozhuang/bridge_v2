from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QToolButton, QFrame, QCheckBox, QLineEdit
from PyQt5.QtGui import QPalette, QColor

from functools import partial

class BD_Floor_Frame(BD_Base_Frame):
    number_of_floor = 28

    def __init__(self, app, qt_line_edit_current_page: QLineEdit, qt_line_edit_total_page: QLineEdit,
                 qt_line_edit_scale: QLineEdit, qt_line_edit_paper_size: QLineEdit,
                 qt_tool_button_1: QToolButton, qt_frame_1: QToolButton,
                 qt_tool_button_2: QFrame, qt_frame_2: QFrame, button_floor_dict: {QCheckBox: QLineEdit}):
        super().__init__(app)

        self.line_edit_current_page = qt_line_edit_current_page
        self.line_edit_total_page = qt_line_edit_total_page
        self.line_edit_scale = qt_line_edit_scale
        self.line_edit_paper_size = qt_line_edit_paper_size
        self.tool_button_1 = qt_tool_button_1
        self.tool_button_2 = qt_tool_button_2
        self.frame_1 = qt_frame_1
        self.frame_2 = qt_frame_2
        self.button_floor_dict = button_floor_dict
        self.number_of_current_floors = 0
        for check_box, line_edit in self.button_floor_dict.items():
            check_box.stateChanged.connect(partial(self.on_state_changed, check_box))
            line_edit.textChanged.connect(partial(self.on_text_changed, line_edit))
        self.tool_button_1.clicked.connect(self.tool_button_1_toggle_visibility)
        self.tool_button_2.clicked.connect(self.tool_button_2_toggle_visibility)

        #why the function is not working
        self.frame_1.setVisible(False)
        self.tool_button_1.setText("+ Expand for more levels")
        self.frame_2.setVisible(False)
        self.tool_button_2.setText("+ Expand for more levels")

    def on_text_changed(self, line_edit):
        text = line_edit.text()
        if text != text.upper():
            cursor_pos = line_edit.cursorPosition()
            line_edit.blockSignals(True)
            line_edit.setText(text.upper())
            line_edit.setCursorPosition(cursor_pos)
            line_edit.blockSignals(False)

    def on_state_changed(self, check_box):
        if check_box.isChecked():
            self.number_of_current_floors += 1
        else:
            self.number_of_current_floors -= 1
        self.line_edit_current_page.setText(str(self.number_of_current_floors))
        palette = self.line_edit_current_page.palette()
        if self.number_of_current_floors==self.line_edit_total_page.text():
            palette.setColor(QPalette.Base, QColor(0, 255, 0))
        else:
            palette.setColor(QPalette.Base, QColor(255, 0, 0))
        self.line_edit_current_page.setPalette(palette)
        # self.line_edit_current_page.setPalette.show()
    def tool_button_1_toggle_visibility(self):
        if self.frame_1.isVisible():
            self.frame_1.setVisible(False)
            self.tool_button_1.setText("+ Expand for more levels")
        else:
            self.frame_1.setVisible(True)
            self.tool_button_1.setText("-")

    def tool_button_2_toggle_visibility(self):
        if self.frame_2.isVisible():
            self.frame_2.setVisible(False)
            self.tool_button_2.setText("+ Expand for more levels")
        else:
            self.frame_2.setVisible(True)
            self.tool_button_2.setText("-")

    def get_selected_floors(self):
        floors = []
        for button, floor in self.button_floor_dict.items():
            if button.isChecked():
                floors.append(floor.text())
        return floors

    def set_total_pages_number(self, pages_number):
        self.line_edit_total_page.setText(str(pages_number))

    def set_scale(self, scale):
        self.line_edit_scale.setText(str(scale))

    def set_paper_size(self, paper_size):
        self.line_edit_paper_size.setText(str(paper_size))
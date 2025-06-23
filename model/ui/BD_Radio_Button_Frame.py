from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QRadioButton, QButtonGroup, QLineEdit


class BD_Radio_Button_Frame(BD_Base_Frame):
    def __init__(self, app, qt_radiobutton_list: [QRadioButton], selections: list,
                 custom_radio_button: QRadioButton=None,custom_line_edit: QLineEdit or tuple=None):
        super().__init__(app)

        self.button_group = QButtonGroup()
        for i, radio_button in enumerate(qt_radiobutton_list):
            self.button_group.addButton(radio_button, i)
        if custom_radio_button:
            self.button_group.addButton(custom_radio_button, len(qt_radiobutton_list))
        self.selections = selections
        if custom_line_edit:
            self.selections.append(custom_line_edit)
        self.custom_radio_button = custom_radio_button
        self.custom_line_edit = custom_line_edit
        # self.button_group.buttonClicked.connect()

    def get_selection_text(self):
        if self.button_group.checkedButton() is None:
            return ""
        return self.button_group.checkedButton().text()
    def get_selection_value(self):
        selection = self.selections[self.button_group.checkedId()]
        if isinstance(selection, tuple):
            if isinstance(selection[0], QLineEdit):
                return (selection[0].text(), selection[1].text())
            return selection
        else:
            if isinstance(selection, QLineEdit):
                return selection.text()
            return selection


from model.ui.BD_Base_Frame import BD_Base_Frame

from ui.Info_Tab import Info_Tab
from ui.Sketch_Tab import Sketch_Tab
from ui.Drawing_Tab import Drawing_Tab
from utility.sql_function import format_output, get_value_from_table_with_filter

from PyQt5.QtWidgets import QMainWindow
from PyQt5 import QtGui

class Main_Window(QMainWindow):
    icon_dir = r"T:\00-Template-Do Not Modify\00-Bridge template\ico\test_logo.ico"
    working_dir = "P:\\"

    def __init__(self, ui):
        super().__init__()
        self.ui = ui

        self.initUI(ui)
        self.setup_tab()

    def initUI(self, ui):
        self.setWindowTitle("Copilot")
        self.setWindowIcon(QtGui.QIcon(Main_Window.icon_dir))
        self.central_widget = ui
        self.setCentralWidget(self.central_widget)
        self.show()
    def setup_tab(self):
        self.info_tab = Info_Tab(self)
        self.sketch_tab = Sketch_Tab(self)
        self.drawing_tab = Drawing_Tab(self)

    def load_project(self, quotation_number):
        project = format_output(get_value_from_table_with_filter("projects", "quotation_number", quotation_number))[quotation_number]
        for key, value in project.items():
            setattr(self, key, value)

        #TODO: need to fix this, use a static method rather then copy paste
        for attribute_name in dir(self):
            if attribute_name.startswith("__"):
                continue
            attribute = getattr(self, attribute_name)
            if isinstance(attribute, BD_Base_Frame):
                attribute.load()
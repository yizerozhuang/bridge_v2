from ui.Main_Window import Main_Window

from PyQt5 import uic
from PyQt5.QtWidgets import QApplication
from conf.config_copilot import CONFIGURATION as conf
from utility.pdf_utility import init_environment
from PyQt5.QtCore import QLoggingCategory
import sys
import multiprocessing
from utility.pdf_tool import PDFTools as PDFTools_v2
from exception import global_exception_hook
if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    init_environment(conf)
    PDFTools_v2.SetEnvironment(conf["bluebeam_dir"], conf["bluebeam_engine_dir"],
                               r"C:\Progra~1\Inkscape\bin\inkscape.exe", conf["c_temp_dir"])
    QLoggingCategory.setFilterRules("*.debug=false\n*.warning=false\n*.critical=false")
    multiprocessing.freeze_support()
    app = QApplication(sys.argv)

    # TODO: Try to use uic.loadUI(path, self)
    ui = uic.loadUi(r"T:\00-Template-Do Not Modify\00-Bridge template\ui\copilot.ui")
    main_window = Main_Window(ui)
    main_window.load_project("000000AA")
    sys.exit(app.exec_())
    # layout_copilot_pro_top = ui.findChild(QGridLayout, 'gridLayout_12')
    # tab_copilot = ui.findChild(QTabWidget, 'tabWidget')

    # ui2=uic.loadUi(r"T:\00-Template-Do Not Modify\00-Bridge template\ui\Copilot-Bridge.ui")
    # frame_bridge_top_1 = ui2.findChild(QFrame, 'frame_pro_stage')
    # frame_bridge_top_2 = ui2.findChild(QFrame, 'frame_pro_func')
    # frame_bridge_top_3 = ui2.findChild(QFrame, 'frame_pro_invoice')
    #
    # tab_bridge_1 = ui2.findChild(QWidget, 'tab_project_info')
    # tab_bridge_2 = ui2.findChild(QWidget, 'tab_fee')
    # tab_bridge_3 = ui2.findChild(QWidget, 'tab_finan')
    #
    # row = layout_copilot_pro_top.rowCount() - 1
    # layout_copilot_pro_top.addWidget(frame_bridge_top_1, row, layout_copilot_pro_top.columnCount())
    # layout_copilot_pro_top.addWidget(frame_bridge_top_2, row, layout_copilot_pro_top.columnCount())
    # layout_copilot_pro_top.addWidget(frame_bridge_top_3, row, layout_copilot_pro_top.columnCount())
    # tab_copilot.addTab(tab_bridge_1, "Project Info")
    # tab_copilot.addTab(tab_bridge_2, "Fee Details")
    # tab_copilot.addTab(tab_bridge_3, "Financial Panel")

    # stats2=Stats_bridge(ui2,main_window)
    # stats.set_quotation_number("000000AA")
    # stats.load_project("000000AA")


    # app.aboutToQuit.connect(main_window.on_close)
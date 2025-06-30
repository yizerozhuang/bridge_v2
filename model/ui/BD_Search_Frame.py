from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QTableWidget

from utility.sql_function import get_cursor_description, get_value_from_table
class BD_Search_Frame(BD_Base_Frame):
    def __init__(self, app, qt_table:QTableWidget):
        super().__init__(app)
        self.table = qt_table
        # self.table.customContextMenuRequested.connect(self.showContextMenu)
        #self.table.horizontalHeader().setFixedHeight(40)

    def load_projects(self):
        projects = get_value_from_table("projects")
        headers = get_cursor_description()
        # for i, project in enumerate(projects):
        #     for j, value in enumerate(project):



    # def show_content_menu(self, event):
    #     contextMenu = QMenu(self.ui)
    #     copyAction = QAction("Copy", self.ui)
    #     copyAction.triggered.connect(self.copy_selection)
    #     contextMenu.addAction(copyAction)
    #     contextMenu.exec_(self.table.viewport().mapToGlobal(event))

    # def copy_selection(self):
    #     try:
    #         selected_indexes = self.ui.table_search_1.selectedIndexes()
    #         if len(selected_indexes) > 0:
    #             selected_rows = list(set(index.row() for index in selected_indexes))
    #             rows_data = []
    #             for row in selected_rows:
    #                 row_data = [self.ui.table_search_1.item(row, col).text() for col in range(self.ui.table_search_1.columnCount())]
    #                 rows_data.append('\t'.join(row_data))
    #             clipboard = QApplication.clipboard()
    #             clipboard.setText('\n'.join(rows_data))
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import Qt


# class BD_Dialog(QDialog):
#     def __init__(self):
#         super().__init__()



class BD_Error_Message(QMessageBox):
    def __init__(self, error_message, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Error')
        self.setTextFormat(Qt.RichText)
        # self.setStyleSheet("QLabel { color:red; }")
        error_message = '<font color="black" size="4">' + error_message.replace('\n', '<br>') + '</font>'
        self.setText(error_message)
        self.setIcon(QMessageBox.Critical)
        self.exec_()

class BD_Info_Message(QMessageBox):
    def __init__(self, info_message, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Info')
        self.setTextFormat(Qt.RichText)
        # self.setStyleSheet("QLabel { color:red; }")
        info_message = '<font color="black" size="4">' + info_message.replace('\n', '<br>') + '</font>'
        self.setText(info_message)
        self.setIcon(QMessageBox.Information)
        self.exec_()

        # set_text = '<font color="black" size="4">' + error_message.replace('\n', '<br>') + '</font>'
        # self.exec_()

# class BD_Info_Message():
#     def __init__(self, info_message):
#         super.__init__()


from model.ui.BD_Base_Frame import BD_Base_Frame

from PyQt5.QtWidgets import QLabel
from PyQt5.QtGui import QGuiApplication, QPixmap

class BD_Clipboard_Frame(BD_Base_Frame):
    def __init__(self, app, qt_label:QLabel):
        super().__init__(app)
        self.label = qt_label

        self.clipboard = QGuiApplication.clipboard()
        self.clipboard.dataChanged.connect(self.on_clipboard_change)

    def on_clipboard_change(self):
        if self.is_image_included():
            image = self.clipboard.image()
            pixmap = QPixmap.fromImage(image)
            self.label.setPixmap(pixmap)
            self.label.setText("")  # Clear any previous text
        # else:
        #     self.label.setPixmap(QPixmap())  # Clear image
        #     self.label.setText("")

    def is_image_included(self):
        return self.clipboard.mimeData().hasImage()

    def save_image(self, file_path):
        if self.is_image_included():
            image = self.clipboard.image()
            image.save(file_path, "PNG")
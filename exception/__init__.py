import sys
import traceback

from exception.BD_Message_Box import BD_Error_Message

# from PyQt5.QtWidgets import QMessageBox
def global_exception_hook(exception_type, exception_value, exception_traceback):
    if issubclass(exception_type, KeyboardInterrupt):
        # Let the keyboard interrupt exit the app normally
        sys.__excepthook__(exception_type, exception_value, exception_traceback)
        return

    error_msg = "".join(traceback.format_exception(exception_type, exception_value, exception_traceback))

    BD_Error_Message(error_msg)


    # QMessageBox.critical(None, "Unexpected Error", error_msg)

# Apply global exception handler

from exception import global_exception_hook


class BD_Base_Frame():
    def __init__(self, app):
        self.app = app
        self.ui = app.ui

    #finishe this wrapper
    def wrapper(self, function):
        def wrapped(*args, **kwargs):
            pass

    def load(self):
        for attribute_name in dir(self):
            if attribute_name.startswith("__"):
                continue
            attribute = getattr(self, attribute_name)
            if isinstance(attribute, BD_Base_Frame):
                attribute.load()

    def save(self):
        pass
    def reset(self):
        pass

    def handle_thread_error(self, exception_info):
        global_exception_hook(*exception_info)
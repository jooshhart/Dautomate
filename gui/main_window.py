from gui.data_page import DataPage

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.page = DataPage(self.root)
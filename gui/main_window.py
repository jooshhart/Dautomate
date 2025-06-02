import tkinter as tk
from gui.pages.start_page import StartPage
from gui.pages.data_page import DataPage
from core.data_processor import DataProcessor
from gui.pages.column_selector_page import ColumnSelectorPage

class MainWindow(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.data_processor = DataProcessor()

        # Container for all pages
        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        # Make sure pages expand to fill the container
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Dictionary to hold references to page frames
        self.frames = {}
        for Page in (StartPage, DataPage, ColumnSelectorPage):
            page_name = Page.__name__
            frame = Page(container, self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Show the start page initially
        self.show_frame("StartPage")

    def show_frame(self, page_name):
        """Raise the frame with the given page_name."""
        frame = self.frames[page_name]
        frame.tkraise()

    def process_data(self, dtype, filename):
        self.data_processor.load_file(filename)
        if dtype == "test":
            self.data_processor.test()  # Replace with your curve logic
        elif dtype == "eff":
            columns = list(self.data_processor.df.columns)
            selector_page = self.frames["ColumnSelectorPage"]
            selector_page.set_columns(columns)
            self.show_frame("ColumnSelectorPage")

    def process_efficiency_with_columns(self, col_map):
        self.data_processor.col_map = col_map
        self.data_processor.process_efficiency()
        # Show DataPage with processed data
        df = self.data_processor.df
        data_page = self.frames["DataPage"]
        data_page.show_dataframe(df, col_map)
        self.show_frame("DataPage")
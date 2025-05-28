import tkinter as tk
from gui.pages.start_page import StartPage
from gui.pages.data_page import DataPage
from core.data_processor import DataProcessor

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
        for Page in (StartPage, DataPage):
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
        elif dtype == "stats":
            # Add stats logic
            pass
        elif dtype == "clean":
            # Add cleaning logic
            pass
        elif dtype == "viz":
            # Add visualization logic
            pass
        df = self.data_processor.df
        data_page = self.frames["DataPage"]
        data_page.show_dataframe(df)
        self.show_frame("DataPage")
        # Optionally, navigate to DataPage or show results
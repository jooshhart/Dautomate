import tkinter as tk
from tkinter import filedialog, messagebox
from core.data_processor import DataProcessor

class DataPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.processor = DataProcessor()

        tk.Button(self, text="Upload Excel", command=self.load_excel).pack(pady=10)
        tk.Button(self, text="Process & Export", command=self.process_excel).pack(pady=10)
        tk.Button(
            self,
            text="← Back to Start",
            command=lambda: controller.show_frame("StartPage")
        ).pack(pady=20)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.processor.load_file(file_path)
            messagebox.showinfo("File Loaded", "Excel file loaded successfully.")

    def process_excel(self):
        if self.processor.df is None:
            messagebox.showwarning("No Data", "Please upload an Excel file first.")
            return

        self.processor.process_data()
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if output_path:
            self.processor.save_file(output_path)
            messagebox.showinfo("Success", "File processed and saved successfully.")
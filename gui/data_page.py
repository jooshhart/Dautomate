import tkinter as tk
from tkinter import filedialog, messagebox
from core.data_processor import DataProcessor

class DataPage:
    def __init__(self, root):
        self.root = root
        self.frame = tk.Frame(root)
        self.frame.pack(padx=20, pady=20)

        self.upload_btn = tk.Button(self.frame, text="Upload Excel", command=self.load_excel)
        self.upload_btn.pack(pady=10)

        self.process_btn = tk.Button(self.frame, text="Process and Export", command=self.process_excel)
        self.process_btn.pack(pady=10)

        self.processor = DataProcessor()

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.processor.load_file(file_path)
            messagebox.showinfo("File Loaded", "Excel file loaded successfully.")

    def process_excel(self):
        if self.processor.df is not None:
            self.processor.process_data()
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            if output_path:
                self.processor.save_file(output_path)
                messagebox.showinfo("Success", "File processed and saved successfully.")
        else:
            messagebox.showwarning("No Data", "Please upload an Excel file first.")
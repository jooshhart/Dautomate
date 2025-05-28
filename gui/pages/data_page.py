import tkinter as tk
from tkinter import ttk

class DataPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f5f5f5")
        self.controller = controller

        self.label = tk.Label(self, text="Processed Data", font=("Helvetica", 18, "bold"), bg="#f5f5f5")
        self.label.pack(pady=10)

        self.tree = ttk.Treeview(self)
        self.tree.pack(expand=True, fill="both", padx=20, pady=20)

    def show_dataframe(self, df):
        # Clear previous content
        for col in self.tree.get_children():
            self.tree.delete(col)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))
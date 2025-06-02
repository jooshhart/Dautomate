from gui.widgets.scrollable_frame import ScrollableFrame
import tkinter as tk
from tkinter import ttk

class ColumnSelectorPage(ScrollableFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f5f5f5")
        self.controller = controller

        # Use self.scrollable_frame for all widgets!
        self.label = tk.Label(self.scrollable_frame, text="Select Columns for Efficiency Calculation", font=("Helvetica", 16, "bold"), bg="#f5f5f5")
        self.label.pack(pady=20)

        self.form_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.form_frame.pack(pady=10)

        self.selectors = {}
        self.column_vars = {}
        for idx, label in enumerate(["VIN", "VOUT", "IIN", "IOUT"]):
            tk.Label(self.form_frame, text=f"{label} Column:", font=("Helvetica", 12), bg="#f5f5f5").grid(row=idx, column=0, sticky="e", padx=10, pady=5)
            var = tk.StringVar()
            cb = ttk.Combobox(self.form_frame, textvariable=var, state="readonly")
            cb.grid(row=idx, column=1, padx=10, pady=5)
            self.selectors[label.lower()] = cb
            self.column_vars[label.lower()] = var

        self.confirm_btn = tk.Button(self.scrollable_frame, text="Confirm", font=("Helvetica", 12, "bold"), command=self.confirm_selection)
        self.confirm_btn.pack(pady=20)

    def set_columns(self, columns):
        for cb in self.selectors.values():
            cb["values"] = columns
            if columns:
                cb.current(0)

    def confirm_selection(self):
        col_map = {k: v.get() for k, v in self.column_vars.items()}
        if all(col_map.values()):
            self.controller.process_efficiency_with_columns(col_map)
        else:
            tk.messagebox.showerror("Selection Error", "Please select all columns.")
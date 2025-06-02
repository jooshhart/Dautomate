from gui.widgets.scrollable_frame import ScrollableFrame
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog
import os

class DataPage(ScrollableFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f5f5f5")
        self.controller = controller

        # Use self.scrollable_frame for all widgets
        self.prompt_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.prompt_frame.pack(pady=(10, 0))

        self.prompt_label = tk.Label(self.prompt_frame, text="Would you like to save the results?", font=("Helvetica", 12), bg="#f5f5f5")
        self.prompt_label.grid(row=0, column=0, padx=5)

        self.filename_var = tk.StringVar()
        self.filename_entry = tk.Entry(self.prompt_frame, textvariable=self.filename_var, width=30)
        self.filename_entry.grid(row=0, column=1, padx=5)
        self.filename_entry.insert(0, "results")

        self.save_btn = tk.Button(self.prompt_frame, text="Save", command=self.save_results, font=("Helvetica", 12, "bold"))
        self.save_btn.grid(row=0, column=2, padx=5)

        # FIX: Pack into self.scrollable_frame, not self
        self.status_label = tk.Label(self.scrollable_frame, text="", font=("Helvetica", 12), bg="#f5f5f5")
        self.status_label.pack(pady=(5, 0))

        self.label = tk.Label(self.scrollable_frame, text="Processed Data", font=("Helvetica", 18, "bold"), bg="#f5f5f5")
        self.label.pack(pady=10)

        self.tree = ttk.Treeview(self.scrollable_frame)
        self.tree.pack(expand=True, fill="both", padx=20, pady=20)

        self.plot_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.plot_frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.df = None
        self.col_map = None

    def show_dataframe(self, df, col_map=None):
        self.df = df
        self.col_map = col_map
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

        # Show the plot
        self.show_plot(df)

    def show_plot(self, df):
    # Clear previous plot
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        fig = Figure(figsize=(5, 4), dpi=100)
        ax = fig.add_subplot(111)
        # Use col_map if available, else fallback to default column names
        x_col = self.col_map["iout"] if self.col_map and "iout" in self.col_map else "IOUT_MEAS (I)"
        y_col = "EFF_CALC (%)"
        # Sort by x_col for line plot
        sorted_df = df.sort_values(by=x_col)
        ax.scatter(sorted_df[x_col], sorted_df[y_col], label="Data Points")
        ax.plot(sorted_df[x_col], sorted_df[y_col], color="orange", label="Trend Line")
        ax.set_xlabel("Load Current (A)")
        ax.set_ylabel("Efficiency (%)")
        ax.set_title("Efficiency Plot")
        ax.legend()

        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.current_fig = fig  # Save for later export

    def save_results(self):
        filename = self.filename_var.get().strip()
        if not filename:
            self.show_status("Please enter a filename.", success=False)
            return
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"

        # Ask user to select a folder
        folder_path = filedialog.askdirectory(title="Select Folder to Save Excel File")
        if not folder_path:
            self.show_status("Save cancelled.", success=False)
            return

        file_path = os.path.join(folder_path, filename)

        try:
            # Save DataFrame and plot to Excel
            import pandas as pd
            from openpyxl import load_workbook
            from openpyxl.drawing.image import Image as XLImage

            self.df.to_excel(file_path, index=False, sheet_name='Data')

            plot_img = "efficiency_plot_temp.png"
            self.current_fig.savefig(plot_img)

            wb = load_workbook(file_path)
            ws = wb['Data']
            img = XLImage(plot_img)
            ws.add_image(img, "H2")
            wb.save(file_path)
            os.remove(plot_img)

            self.show_status(f"File saved as {file_path}", success=True)
            self.after(2000, self.return_to_start)
        except Exception as e:
            self.show_status(f"Error saving file: {e}", success=False)

    def show_status(self, message, success=True):
        self.status_label.config(
            text=message,
            fg="green" if success else "red",
            bg="#eaffea" if success else "#ffeaea"
        )
        self.status_label.after(2000 if success else 4000, lambda: self.status_label.config(text="", bg="#f5f5f5"))

    def return_to_start(self):
        self.status_label.config(text="", bg="#f5f5f5")
        self.controller.show_frame("StartPage")
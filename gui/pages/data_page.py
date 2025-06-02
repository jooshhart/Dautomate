from gui.widgets.scrollable_frame import ScrollableFrame
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog
import xlsxwriter
import os

class DataPage(ScrollableFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f5f5f5")
        self.controller = controller

        # Save prompt
        self.prompt_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.prompt_frame.pack(pady=(10, 0))

        self.prompt_label = tk.Label(self.prompt_frame, text="Would you like to save the results?", font=("Helvetica", 12), bg="#f5f5f5")
        self.prompt_label.grid(row=0, column=0, padx=5)

        self.filename_var = tk.StringVar()
        self.filename_entry = tk.Entry(self.prompt_frame, textvariable=self.filename_var, width=30)
        self.filename_entry.grid(row=0, column=1, padx=5)
        self.filename_entry.insert(0, "results")

        self.save_eff_var = tk.BooleanVar(value=True)
        self.save_loss_var = tk.BooleanVar(value=True)
        self.save_combined_var = tk.BooleanVar(value=True)
        self.save_eff_cb = tk.Checkbutton(self.prompt_frame, text="Save Efficiency Graph", variable=self.save_eff_var, bg="#f5f5f5")
        self.save_loss_cb = tk.Checkbutton(self.prompt_frame, text="Save Power Loss Graph", variable=self.save_loss_var, bg="#f5f5f5")
        self.save_combined_cb = tk.Checkbutton(self.prompt_frame, text="Save Combined Graph", variable=self.save_combined_var, bg="#f5f5f5")
        self.save_eff_cb.grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.save_loss_cb.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        self.save_combined_cb.grid(row=1, column=2, padx=5, pady=2, sticky="w")

        self.save_btn = tk.Button(self.prompt_frame, text="Save", command=self.save_results, font=("Helvetica", 12, "bold"))
        self.save_btn.grid(row=1, column=3, padx=5)

        self.status_label = tk.Label(self.scrollable_frame, text="", font=("Helvetica", 12), bg="#f5f5f5")
        self.status_label.pack(pady=(5, 0))

        self.label = tk.Label(self.scrollable_frame, text="Processed Data", font=("Helvetica", 18, "bold"), bg="#f5f5f5")
        self.label.pack(pady=10)

        self.tree = ttk.Treeview(self.scrollable_frame)
        self.tree.pack(expand=True, fill="both", padx=20, pady=20)

        # --- Efficiency Graph Controls and Plot ---
        self.eff_axes_frame = tk.LabelFrame(self.scrollable_frame, text="Efficiency Graph Controls", bg="#f5f5f5")
        self.eff_axes_frame.pack(pady=(10, 0), fill="x", padx=10)
        tk.Label(self.eff_axes_frame, text="X min:", bg="#f5f5f5").grid(row=0, column=0)
        tk.Label(self.eff_axes_frame, text="X max:", bg="#f5f5f5").grid(row=0, column=2)
        tk.Label(self.eff_axes_frame, text="Y min:", bg="#f5f5f5").grid(row=0, column=4)
        tk.Label(self.eff_axes_frame, text="Y max:", bg="#f5f5f5").grid(row=0, column=6)
        self.eff_xmin_var = tk.StringVar()
        self.eff_xmax_var = tk.StringVar()
        self.eff_ymin_var = tk.StringVar()
        self.eff_ymax_var = tk.StringVar()
        eff_xmin_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_xmin_var, width=7)
        eff_xmax_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_xmax_var, width=7)
        eff_ymin_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_ymin_var, width=7)
        eff_ymax_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_ymax_var, width=7)
        eff_xmin_entry.grid(row=0, column=1)
        eff_xmax_entry.grid(row=0, column=3)
        eff_ymin_entry.grid(row=0, column=5)
        eff_ymax_entry.grid(row=0, column=7)
        tk.Label(self.eff_axes_frame, text="X label:", bg="#f5f5f5").grid(row=1, column=0)
        tk.Label(self.eff_axes_frame, text="Y label:", bg="#f5f5f5").grid(row=1, column=2)
        tk.Label(self.eff_axes_frame, text="Title:", bg="#f5f5f5").grid(row=1, column=4)
        self.eff_xlabel_var = tk.StringVar(value="Load Current (A)")
        self.eff_ylabel_var = tk.StringVar(value="Efficiency (%)")
        self.eff_title_var = tk.StringVar(value="Efficiency Plot")
        eff_xlabel_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_xlabel_var, width=15)
        eff_ylabel_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_ylabel_var, width=15)
        eff_title_entry = tk.Entry(self.eff_axes_frame, textvariable=self.eff_title_var, width=20)
        eff_xlabel_entry.grid(row=1, column=1)
        eff_ylabel_entry.grid(row=1, column=3)
        eff_title_entry.grid(row=1, column=5, columnspan=3, sticky="ew")
        tk.Button(self.eff_axes_frame, text="Update Efficiency Graph", command=lambda: self.show_eff_plot(self.df)).grid(row=2, column=0, columnspan=8, pady=5)

        for entry in (eff_xmin_entry, eff_xmax_entry, eff_ymin_entry, eff_ymax_entry, eff_xlabel_entry, eff_ylabel_entry, eff_title_entry):
            entry.bind("<Return>", lambda event: self.show_eff_plot(self.df))

        self.eff_plot_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.eff_plot_frame.pack(expand=True, fill="both", padx=20, pady=10)

        # --- Power Loss Graph Controls and Plot ---
        self.loss_axes_frame = tk.LabelFrame(self.scrollable_frame, text="Power Loss Graph Controls", bg="#f5f5f5")
        self.loss_axes_frame.pack(pady=(10, 0), fill="x", padx=10)
        tk.Label(self.loss_axes_frame, text="X min:", bg="#f5f5f5").grid(row=0, column=0)
        tk.Label(self.loss_axes_frame, text="X max:", bg="#f5f5f5").grid(row=0, column=2)
        tk.Label(self.loss_axes_frame, text="Y min:", bg="#f5f5f5").grid(row=0, column=4)
        tk.Label(self.loss_axes_frame, text="Y max:", bg="#f5f5f5").grid(row=0, column=6)
        self.loss_xmin_var = tk.StringVar()
        self.loss_xmax_var = tk.StringVar()
        self.loss_ymin_var = tk.StringVar()
        self.loss_ymax_var = tk.StringVar()
        loss_xmin_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_xmin_var, width=7)
        loss_xmax_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_xmax_var, width=7)
        loss_ymin_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_ymin_var, width=7)
        loss_ymax_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_ymax_var, width=7)
        loss_xmin_entry.grid(row=0, column=1)
        loss_xmax_entry.grid(row=0, column=3)
        loss_ymin_entry.grid(row=0, column=5)
        loss_ymax_entry.grid(row=0, column=7)
        tk.Label(self.loss_axes_frame, text="X label:", bg="#f5f5f5").grid(row=1, column=0)
        tk.Label(self.loss_axes_frame, text="Y label:", bg="#f5f5f5").grid(row=1, column=2)
        tk.Label(self.loss_axes_frame, text="Title:", bg="#f5f5f5").grid(row=1, column=4)
        self.loss_xlabel_var = tk.StringVar(value="Load Current (A)")
        self.loss_ylabel_var = tk.StringVar(value="Power Loss (W)")
        self.loss_title_var = tk.StringVar(value="Power Loss Plot")
        loss_xlabel_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_xlabel_var, width=15)
        loss_ylabel_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_ylabel_var, width=15)
        loss_title_entry = tk.Entry(self.loss_axes_frame, textvariable=self.loss_title_var, width=20)
        loss_xlabel_entry.grid(row=1, column=1)
        loss_ylabel_entry.grid(row=1, column=3)
        loss_title_entry.grid(row=1, column=5, columnspan=3, sticky="ew")
        tk.Button(self.loss_axes_frame, text="Update Power Loss Graph", command=lambda: self.show_loss_plot(self.df)).grid(row=2, column=0, columnspan=8, pady=5)

        for entry in (loss_xmin_entry, loss_xmax_entry, loss_ymin_entry, loss_ymax_entry, loss_xlabel_entry, loss_ylabel_entry, loss_title_entry):
            entry.bind("<Return>", lambda event: self.show_loss_plot(self.df))

        self.loss_plot_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.loss_plot_frame.pack(expand=True, fill="both", padx=20, pady=10)

        # --- Combined Graph Controls and Plot ---
        self.comb_axes_frame = tk.LabelFrame(self.scrollable_frame, text="Combined Graph Controls", bg="#f5f5f5")
        self.comb_axes_frame.pack(pady=(10, 0), fill="x", padx=10)
        # X axis
        tk.Label(self.comb_axes_frame, text="X min:", bg="#f5f5f5").grid(row=0, column=0)
        tk.Label(self.comb_axes_frame, text="X max:", bg="#f5f5f5").grid(row=0, column=2)
        self.comb_xmin_var = tk.StringVar()
        self.comb_xmax_var = tk.StringVar()
        comb_xmin_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_xmin_var, width=7)
        comb_xmax_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_xmax_var, width=7)
        comb_xmin_entry.grid(row=0, column=1)
        comb_xmax_entry.grid(row=0, column=3)
        # Left Y axis (Efficiency)
        tk.Label(self.comb_axes_frame, text="Eff Y min:", bg="#f5f5f5").grid(row=0, column=4)
        tk.Label(self.comb_axes_frame, text="Eff Y max:", bg="#f5f5f5").grid(row=0, column=6)
        self.comb_eff_ymin_var = tk.StringVar()
        self.comb_eff_ymax_var = tk.StringVar()
        comb_eff_ymin_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_eff_ymin_var, width=7)
        comb_eff_ymax_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_eff_ymax_var, width=7)
        comb_eff_ymin_entry.grid(row=0, column=5)
        comb_eff_ymax_entry.grid(row=0, column=7)
        # Right Y axis (Power Loss)
        tk.Label(self.comb_axes_frame, text="Loss Y min:", bg="#f5f5f5").grid(row=1, column=0)
        tk.Label(self.comb_axes_frame, text="Loss Y max:", bg="#f5f5f5").grid(row=1, column=2)
        self.comb_loss_ymin_var = tk.StringVar()
        self.comb_loss_ymax_var = tk.StringVar()
        comb_loss_ymin_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_loss_ymin_var, width=7)
        comb_loss_ymax_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_loss_ymax_var, width=7)
        comb_loss_ymin_entry.grid(row=1, column=1)
        comb_loss_ymax_entry.grid(row=1, column=3)
        # Labels and title
        tk.Label(self.comb_axes_frame, text="X label:", bg="#f5f5f5").grid(row=1, column=4)
        tk.Label(self.comb_axes_frame, text="Eff Y label:", bg="#f5f5f5").grid(row=1, column=6)
        tk.Label(self.comb_axes_frame, text="Loss Y label:", bg="#f5f5f5").grid(row=2, column=0)
        tk.Label(self.comb_axes_frame, text="Title:", bg="#f5f5f5").grid(row=2, column=2)
        self.comb_xlabel_var = tk.StringVar(value="Load Current (A)")
        self.comb_eff_ylabel_var = tk.StringVar(value="Efficiency (%)")
        self.comb_loss_ylabel_var = tk.StringVar(value="Power Loss (W)")
        self.comb_title_var = tk.StringVar(value="Efficiency & Power Loss Plot")
        comb_xlabel_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_xlabel_var, width=15)
        comb_eff_ylabel_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_eff_ylabel_var, width=15)
        comb_loss_ylabel_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_loss_ylabel_var, width=15)
        comb_title_entry = tk.Entry(self.comb_axes_frame, textvariable=self.comb_title_var, width=20)
        comb_xlabel_entry.grid(row=1, column=5)
        comb_eff_ylabel_entry.grid(row=1, column=7)
        comb_loss_ylabel_entry.grid(row=2, column=1)
        comb_title_entry.grid(row=2, column=3, columnspan=3, sticky="ew")
        tk.Button(self.comb_axes_frame, text="Update Combined Graph", command=lambda: self.show_combined_plot(self.df)).grid(row=3, column=0, columnspan=8, pady=5)

        for entry in (
            comb_xmin_entry, comb_xmax_entry, comb_eff_ymin_entry, comb_eff_ymax_entry,
            comb_loss_ymin_entry, comb_loss_ymax_entry, comb_xlabel_entry,
            comb_eff_ylabel_entry, comb_loss_ylabel_entry, comb_title_entry
        ):
            entry.bind("<Return>", lambda event: self.show_combined_plot(self.df))

        self.comb_plot_frame = tk.Frame(self.scrollable_frame, bg="#f5f5f5")
        self.comb_plot_frame.pack(expand=True, fill="both", padx=20, pady=10)

        self.df = None
        self.col_map = None
        self.current_eff_fig = None
        self.current_loss_fig = None
        self.current_comb_fig = None

    def show_dataframe(self, df, col_map=None):
        self.df = df
        self.col_map = col_map
        # Clear previous content
        for col in self.tree.get_children():
            self.tree.delete(col)
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def show_eff_plot(self, df):
        for widget in self.eff_plot_frame.winfo_children():
            widget.destroy()
        if df is None:
            return
        fig = Figure(figsize=(3.45, 4), dpi=100)
        ax = fig.add_subplot(111)
        x_col = self.col_map["iout"] if self.col_map and "iout" in self.col_map else "IOUT_MEAS (I)"
        y_col = "EFF_CALC (%)"
        sorted_df = df.sort_values(by=x_col)
        ax.scatter(sorted_df[x_col], sorted_df[y_col], label="Data Points")
        ax.plot(sorted_df[x_col], sorted_df[y_col], color="orange", label="Trend Line")
        ax.set_xlabel(self.eff_xlabel_var.get())
        ax.set_ylabel(self.eff_ylabel_var.get())
        ax.set_title(self.eff_title_var.get())
        ax.grid(True)
        ax.legend()
        # Axis limits
        try:
            xmin = float(self.eff_xmin_var.get()) if self.eff_xmin_var.get() else None
            xmax = float(self.eff_xmax_var.get()) if self.eff_xmax_var.get() else None
            ymin = float(self.eff_ymin_var.get()) if self.eff_ymin_var.get() else None
            ymax = float(self.eff_ymax_var.get()) if self.eff_ymax_var.get() else None
            if xmin is not None or xmax is not None:
                ax.set_xlim(left=xmin if xmin is not None else None, right=xmax if xmax is not None else None)
            if ymin is not None or ymax is not None:
                ax.set_ylim(bottom=ymin if ymin is not None else None, top=ymax if ymax is not None else None)
        except Exception:
            pass
        canvas = FigureCanvasTkAgg(fig, master=self.eff_plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.current_eff_fig = fig

    def show_loss_plot(self, df):
        for widget in self.loss_plot_frame.winfo_children():
            widget.destroy()
        if df is None:
            return
        fig = Figure(figsize=(3.45, 4), dpi=100)
        ax = fig.add_subplot(111)
        x_col = self.col_map["iout"] if self.col_map and "iout" in self.col_map else "IOUT_MEAS (I)"
        # Power Loss calculation
        if "VIN_MEAS (V)" in df.columns and "IIN_MEAS (I)" in df.columns and "VOUT_MEAS (V)" in df.columns and "IOUT_MEAS (I)" in df.columns:
            power_loss = (df["VIN_MEAS (V)"] * df["IIN_MEAS (I)"]) - (df["VOUT_MEAS (V)"] * df["IOUT_MEAS (I)"])
        else:
            power_loss = df.get("Power Loss (W)", None)
        y_col = "Power Loss (W)"
        sorted_df = df.copy()
        sorted_df[y_col] = power_loss
        sorted_df = sorted_df.sort_values(by=x_col)
        ax.scatter(sorted_df[x_col], sorted_df[y_col], label="Data Points")
        ax.plot(sorted_df[x_col], sorted_df[y_col], color="red", label="Trend Line")
        ax.set_xlabel(self.loss_xlabel_var.get())
        ax.set_ylabel(self.loss_ylabel_var.get())
        ax.set_title(self.loss_title_var.get())
        ax.grid(True)
        ax.legend()
        # Axis limits
        try:
            xmin = float(self.loss_xmin_var.get()) if self.loss_xmin_var.get() else None
            xmax = float(self.loss_xmax_var.get()) if self.loss_xmax_var.get() else None
            ymin = float(self.loss_ymin_var.get()) if self.loss_ymin_var.get() else None
            ymax = float(self.loss_ymax_var.get()) if self.loss_ymax_var.get() else None
            if xmin is not None or xmax is not None:
                ax.set_xlim(left=xmin if xmin is not None else None, right=xmax if xmax is not None else None)
            if ymin is not None or ymax is not None:
                ax.set_ylim(bottom=ymin if ymin is not None else None, top=ymax if ymax is not None else None)
        except Exception:
            pass
        canvas = FigureCanvasTkAgg(fig, master=self.loss_plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.current_loss_fig = fig

    def show_combined_plot(self, df):
        for widget in self.comb_plot_frame.winfo_children():
            widget.destroy()
        if df is None:
            return
        fig = Figure(figsize=(3.45, 4), dpi=100)
        ax1 = fig.add_subplot(111)
        x_col = self.col_map["iout"] if self.col_map and "iout" in self.col_map else "IOUT_MEAS (I)"
        eff_y_col = "EFF_CALC (%)"
        loss_y_col = "Power Loss (W)"
        # Power Loss calculation
        if "VIN_MEAS (V)" in df.columns and "IIN_MEAS (I)" in df.columns and "VOUT_MEAS (V)" in df.columns and "IOUT_MEAS (I)" in df.columns:
            power_loss = (df["VIN_MEAS (V)"] * df["IIN_MEAS (I)"]) - (df["VOUT_MEAS (V)"] * df["IOUT_MEAS (I)"])
        else:
            power_loss = df.get("Power Loss (W)", None)
        sorted_df = df.copy()
        sorted_df[loss_y_col] = power_loss
        sorted_df = sorted_df.sort_values(by=x_col)
        # Efficiency (left y)
        ax1.scatter(sorted_df[x_col], sorted_df[eff_y_col], label="Efficiency", color="orange")
        ax1.plot(sorted_df[x_col], sorted_df[eff_y_col], color="orange")
        ax1.set_xlabel(self.comb_xlabel_var.get())
        ax1.set_ylabel(self.comb_eff_ylabel_var.get(), color="orange")
        ax1.tick_params(axis='y', labelcolor="orange")
        # Axis limits for left y
        try:
            xmin = float(self.comb_xmin_var.get()) if self.comb_xmin_var.get() else None
            xmax = float(self.comb_xmax_var.get()) if self.comb_xmax_var.get() else None
            eff_ymin = float(self.comb_eff_ymin_var.get()) if self.comb_eff_ymin_var.get() else None
            eff_ymax = float(self.comb_eff_ymax_var.get()) if self.comb_eff_ymax_var.get() else None
            if xmin is not None or xmax is not None:
                ax1.set_xlim(left=xmin if xmin is not None else None, right=xmax if xmax is not None else None)
            if eff_ymin is not None or eff_ymax is not None:
                ax1.set_ylim(bottom=eff_ymin if eff_ymin is not None else None, top=eff_ymax if eff_ymax is not None else None)
        except Exception:
            pass
        ax1.grid(True)
        # Power Loss (right y)
        ax2 = ax1.twinx()
        ax2.scatter(sorted_df[x_col], sorted_df[loss_y_col], label="Power Loss", color="red")
        ax2.plot(sorted_df[x_col], sorted_df[loss_y_col], color="red")
        ax2.set_ylabel(self.comb_loss_ylabel_var.get(), color="red")
        ax2.tick_params(axis='y', labelcolor="red")
        # Axis limits for right y
        try:
            loss_ymin = float(self.comb_loss_ymin_var.get()) if self.comb_loss_ymin_var.get() else None
            loss_ymax = float(self.comb_loss_ymax_var.get()) if self.comb_loss_ymax_var.get() else None
            if loss_ymin is not None or loss_ymax is not None:
                ax2.set_ylim(bottom=loss_ymin if loss_ymin is not None else None, top=loss_ymax if loss_ymax is not None else None)
        except Exception:
            pass
        fig.suptitle(self.comb_title_var.get())
        canvas = FigureCanvasTkAgg(fig, master=self.comb_plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.current_comb_fig = fig

    def save_results(self):
        import pandas as pd
        import xlsxwriter

        filename = self.filename_var.get().strip()
        if not filename:
            self.show_status("Please enter a filename.", success=False)
            return
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"

        folder_path = filedialog.askdirectory(title="Select Folder to Save Excel File")
        if not folder_path:
            self.show_status("Save cancelled.", success=False)
            return

        file_path = os.path.join(folder_path, filename)

        try:
            # Save DataFrame to Excel using xlsxwriter
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                # Calculate Power Loss if not present
                if "Power Loss (W)" not in self.df.columns:
                    if all(col in self.df.columns for col in ["VIN_MEAS (V)", "IIN_MEAS (I)", "VOUT_MEAS (V)", "IOUT_MEAS (I)"]):
                        self.df["Power Loss (W)"] = (self.df["VIN_MEAS (V)"] * self.df["IIN_MEAS (I)"]) - (self.df["VOUT_MEAS (V)"] * self.df["IOUT_MEAS (I)"])
                self.df.to_excel(writer, index=False, sheet_name='Data')
                workbook = writer.book
                worksheet = writer.sheets['Data']

                columns = list(self.df.columns)
                x_col = self.col_map["iout"] if self.col_map and "iout" in self.col_map else "IOUT_MEAS (I)"
                eff_col = "EFF_CALC (%)"
                loss_col = "Power Loss (W)"
                x_idx = columns.index(x_col)
                eff_idx = columns.index(eff_col)
                loss_idx = columns.index(loss_col)
                nrows = len(self.df)

                # 1. Efficiency Chart
                if self.save_eff_var.get():
                    chart1 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                    chart1.add_series({
                        'name':       self.eff_title_var.get(),
                        'categories': ['Data', 1, x_idx, nrows, x_idx],
                        'values':     ['Data', 1, eff_idx, nrows, eff_idx],
                        'marker':     {'type': 'circle', 'size': 5, 'border': {'color': 'orange'}, 'fill': {'color': 'orange'}},
                        'line':       {'color': 'orange'},
                    })
                    chart1.set_title({'name': self.eff_title_var.get()})
                    chart1.set_x_axis({'name': self.eff_xlabel_var.get()})
                    chart1.set_y_axis({'name': self.eff_ylabel_var.get()})
                    worksheet.insert_chart('H2', chart1)

                # 2. Power Loss Chart
                if self.save_loss_var.get():
                    chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                    chart2.add_series({
                        'name':       self.loss_title_var.get(),
                        'categories': ['Data', 1, x_idx, nrows, x_idx],
                        'values':     ['Data', 1, loss_idx, nrows, loss_idx],
                        'marker':     {'type': 'circle', 'size': 5, 'border': {'color': 'red'}, 'fill': {'color': 'red'}},
                        'line':       {'color': 'red'},
                    })
                    chart2.set_title({'name': self.loss_title_var.get()})
                    chart2.set_x_axis({'name': self.loss_xlabel_var.get()})
                    chart2.set_y_axis({'name': self.loss_ylabel_var.get()})
                    worksheet.insert_chart('H20', chart2)

                # 3. Combined Chart (Efficiency + Power Loss, two y-axes)
                if self.save_combined_var.get():
                    chart3 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                    # Efficiency (primary y-axis)
                    chart3.add_series({
                        'name':       self.comb_eff_ylabel_var.get(),
                        'categories': ['Data', 1, x_idx, nrows, x_idx],
                        'values':     ['Data', 1, eff_idx, nrows, eff_idx],
                        'y2_axis':    False,
                        'marker':     {'type': 'circle', 'size': 5, 'border': {'color': 'orange'}, 'fill': {'color': 'orange'}},
                        'line':       {'color': 'orange'},
                    })
                    # Power Loss (secondary y-axis)
                    chart3.add_series({
                        'name':       self.comb_loss_ylabel_var.get(),
                        'categories': ['Data', 1, x_idx, nrows, x_idx],
                        'values':     ['Data', 1, loss_idx, nrows, loss_idx],
                        'y2_axis':    True,
                        'marker':     {'type': 'circle', 'size': 5, 'border': {'color': 'red'}, 'fill': {'color': 'red'}},
                        'line':       {'color': 'red'},
                    })
                    chart3.set_title({'name': self.comb_title_var.get()})
                    chart3.set_x_axis({'name': self.comb_xlabel_var.get()})
                    chart3.set_y_axis({'name': self.comb_eff_ylabel_var.get()})
                    chart3.set_y2_axis({'name': self.comb_loss_ylabel_var.get()})
                    worksheet.insert_chart('H38', chart3)

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
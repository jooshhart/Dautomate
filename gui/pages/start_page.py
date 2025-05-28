import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageTk, ImageSequence
import os

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f5f5f5")
        self.controller = controller

        self.gif_path = os.path.join("resources", "DA.gif")
        self.video_frame = tk.Frame(self, bg="#f5f5f5")
        self.video_frame.pack(expand=True, fill="both")

        # Load and play GIF
        self.gif = Image.open(self.gif_path)
        self.gif_frames = [ImageTk.PhotoImage(frame.copy().convert("RGBA")) for frame in ImageSequence.Iterator(self.gif)]
        self.gif_label = tk.Label(self.video_frame, bg="#f5f5f5")
        self.gif_label.pack(expand=True)
        self.current_frame = 0
        self.after_id = None
        self.play_gif()

    def play_gif(self):
        frame = self.gif_frames[self.current_frame]
        self.gif_label.config(image=frame)
        self.current_frame += 1
        if self.current_frame < len(self.gif_frames):
            self.after_id = self.after(20, self.play_gif)  # Adjust delay as needed
        else:
            self.after(100, self.fade_to_menu)  # Small pause before showing menu

    def fade_to_menu(self):
        # Create a white overlay label
        self.fade_overlay = tk.Label(self, bg="#f5f5f5")
        self.fade_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.fade_alpha = 0
        self._fade_step()

    def _fade_step(self):
        # Simulate fade by raising overlay and increasing its opacity
        # Tkinter doesn't support alpha, so we simulate with color steps
        steps = 10
        if self.fade_alpha < steps:
            base = 245
            val = base + int((255 - base) * self.fade_alpha / steps)
            color = f'#{val:02x}{val:02x}{val:02x}'  # e.g., #f5f5f5 to #ffffff
            self.fade_overlay.config(bg=color)
            self.fade_alpha += 1
            self.after(20, self._fade_step)
        else:
            self.fade_overlay.destroy()
            self.show_menu()

    def show_menu(self, event=None):
        if self.after_id:
            self.after_cancel(self.after_id)
        self.video_frame.destroy()

        # Title
        title = tk.Label(
            self,
            text="Welcome to Dautomate",
            font=("Helvetica", 24, "bold"),
            bg="#f5f5f5",
            fg="#333"
        )
        title.pack(pady=(20, 5))

        # Separator
        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill="x", padx=40, pady=10)

        # Prompt
        prompt = tk.Label(
            self,
            text="How would you like to process your data?",
            font=("Helvetica", 14),
            bg="#f5f5f5",
            fg="#666"
        )
        prompt.pack(pady=(0, 20))

        # Hub-style menu frame
        hub_frame = tk.Frame(self, bg="#e0e0e0", bd=3, relief="ridge")
        hub_frame.pack(pady=20, padx=60)

        # Example processing types (add more as needed)
        types = [
            ("Testing", "test"),
            ("Statistical Analysis", "stats"),
            ("Data Cleaning", "clean"),
            ("Visualization", "viz"),
        ]

        # Arrange buttons in a grid (2 columns)
        for idx, (text, dtype) in enumerate(types):
            row = idx // 2 + 1  # +1 because row 0 is the heading
            col = idx % 2
            btn = tk.Button(
                hub_frame,
                text=text,
                font=("Helvetica", 12),
                width=20,
                pady=12,
                bg="#fff",
                fg="#333",
                relief="raised",
                bd=2,
                cursor="hand2",
                command=lambda t=dtype: self.select_file(t)
            )
            btn.grid(row=row, column=col, padx=20, pady=10, sticky="nsew")

        # Make columns expand equally
        hub_frame.grid_columnconfigure(0, weight=1)
        hub_frame.grid_columnconfigure(1, weight=1)

    def select_file(self, dtype):
        filetypes = [("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        filename = filedialog.askopenfilename(
            title="Select data file",
            filetypes=filetypes
        )
        if filename:
            self.controller.process_data(dtype, filename)
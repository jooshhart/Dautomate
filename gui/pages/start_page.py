import tkinter as tk

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f5f5f5")
        self.controller = controller

        # Create a centered wrapper to hold content
        wrapper = tk.Frame(self, bg="#f5f5f5")
        wrapper.place(relx=0.5, rely=0.5, anchor="center")

        # Main title
        title = tk.Label(
            wrapper,
            text="Welcome to Dautomate",
            font=("Helvetica", 24, "bold"),
            bg="#f5f5f5",
            fg="#333"
        )
        title.pack(pady=(0, 10))

        # Subtitle explanation
        subtitle = tk.Label(
            wrapper,
            text="Automate your data processing with ease.",
            font=("Helvetica", 14),
            bg="#f5f5f5",
            fg="#666"
        )
        subtitle.pack(pady=(0, 30))

        # Styled start button
        start_btn = tk.Button(
            wrapper,
            text="Start Processing",
            font=("Helvetica", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
            bd=0,
            relief="ridge",
            activebackground="#45a049",
            cursor="hand2",
            command=lambda: controller.show_frame("DataPage")
        )
        start_btn.pack()
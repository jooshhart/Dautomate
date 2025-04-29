import tkinter as tk
from gui.main_window import MainWindow
from PIL import Image, ImageTk

if __name__ == "__main__":
    app = tk.Tk()
    app.title("Dautomate")
    app.geometry("600x400")

    # Load application icon
    icon_img = Image.open("resources/Dautomate.png")
    icon_photo = ImageTk.PhotoImage(icon_img)
    app.iconphoto(False, icon_photo)

    # Initialize and pack the main frame
    main_frame = MainWindow(app)
    main_frame.pack(fill="both", expand=True)

    app.mainloop()
from utils import resource_path
import tkinter as tk
from gui.main_window import MainWindow
from PIL import Image, ImageTk

if __name__ == "__main__":
    app = tk.Tk()
    app.title("Dautomate")
    app.resizable(True, True)
    # Set window size to full screen
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    app.geometry(f"{screen_width}x{screen_height}")

    # Load application icon for window and taskbar
    try:
        app.iconbitmap(resource_path("resources/Dautomate.ico"))
    except Exception as e:
        print("ICO icon not found or failed to load:", e)
        icon_img = Image.open(resource_path("resources/Dautomate.jpg"))
        icon_photo = ImageTk.PhotoImage(icon_img)
        app.iconphoto(False, icon_photo)

    # Initialize and pack the main frame
    main_frame = MainWindow(app)
    main_frame.pack(fill="both", expand=True)

    app.mainloop()
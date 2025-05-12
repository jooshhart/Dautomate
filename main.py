import tkinter as tk
from gui.main_window import MainWindow
from PIL import Image, ImageTk

if __name__ == "__main__":
    app = tk.Tk()
    app.title("Dautomate")
    app.geometry("600x400")

    # Load application icon for window and taskbar
    try:
        app.iconbitmap("resources/Dautomate.ico")  # Use .ico for taskbar icon
    except Exception as e:
        print("ICO icon not found or failed to load:", e)
        # Fallback to PNG for window icon
        icon_img = Image.open("resources/Dautomate.jpg")
        icon_photo = ImageTk.PhotoImage(icon_img)
        app.iconphoto(False, icon_photo)

    # Initialize and pack the main frame
    main_frame = MainWindow(app)
    main_frame.pack(fill="both", expand=True)

    app.mainloop()
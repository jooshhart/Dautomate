import tkinter as tk
from tkinter import ttk

class ScrollableFrame(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg=self["bg"])
        v_scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg=self["bg"])

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        self._window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")

        def _on_canvas_configure(event):
            bbox = canvas.bbox(self._window)
            if bbox is None:
                return
            frame_width = bbox[2] - bbox[0]
            canvas_width = event.width

            y = 0
            if frame_width < canvas_width:
                x = (canvas_width - frame_width) // 2
            else:
                x = 0
            canvas.coords(self._window, x, y)
            canvas.itemconfig(self._window, width=canvas_width)

        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.configure(yscrollcommand=v_scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")

        self._bind_mousewheel(canvas)

    def _bind_mousewheel(self, canvas):
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        def bind_to_children(widget):
            widget.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
            widget.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
            for child in widget.winfo_children():
                bind_to_children(child)
        bind_to_children(self.scrollable_frame)
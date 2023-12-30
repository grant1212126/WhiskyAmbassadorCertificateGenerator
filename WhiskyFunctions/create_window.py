import tkinter as tk


def create_window(window_name, window_title, root=None):

    if root is None:
        window = tk.TK()

    else:
        window = tk.Toplevel(root)
        window.grab_set()

    window.title(window_title)

    window.resizable(True, True)

    window_width = window_name.winfo_reqwidth()
    window_height = window_name.winfo_reqheight()

    position_right = int(window_name.winfo_screenwidth() / 2 - window_width / 2)
    position_down = int(window_name.winfo_screenheight() / 2 - window_height / 2)

    window_name.geometry("+{}+{}".format(position_right, position_down))

    window.update_idletasks()

    return window

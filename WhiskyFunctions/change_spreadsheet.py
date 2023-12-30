import os
import tkinter as tk

from tkinter import messagebox
from tkinter import filedialog

"""
@param SpreadSheetExists
@param
"""


def change_spreadsheet(certificate_template_list):
    tk.Tk().withdraw()
    new_spread_sheet = tk.filedialog.askopenfilename()

    new_spread_sheet = os.path.basename(new_spread_sheet)

    if ".xlsx" not in new_spread_sheet:
        tk.messagebox.showerror(title="Information",
                                message="Error, you have not selected a valid spreadsheet file, make sure you select "
                                        "a spreadsheet file with the extension '.xlsx'")
        return

    with open("config.txt", "w") as f:

        for entry in certificate_template_list:
            f.write("Certificate_template=" + entry)
            f.write("\n")

        f.write("Spreadsheet=" + new_spread_sheet + "\n")

    tk.messagebox.showinfo(title="Information", message="Spreadsheet file has been successfully selected and changed!")

    return True, new_spread_sheet

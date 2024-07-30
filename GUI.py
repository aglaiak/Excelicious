import importlib
import subprocess

def install_package(package):
    try:
        importlib.import_module(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_package('customtkinter')
install_package('pandas')
install_package('tkinter')

import pandas as pd
import subprocess
import customtkinter as ctk
import tkinter as tk 
from tkinter import filedialog, messagebox, Button, scrolledtext
from DuplicateBarcodes import find_unique_barcodes, find_edited_values, compare_values

def display_scrollable_message(title, message):
    """Generate a scrollable window"""
    root = tk.Tk()
    root.title(title)

    canvas = tk.Canvas(root, width=600, height=400)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(root, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor=tk.NW)

    label = tk.Label(frame, text=message, justify=tk.LEFT)
    label.pack()

    frame.bind("<Configure>", lambda event, canvas=canvas: on_frame_configure(canvas))

    canvas.configure(yscrollcommand=scrollbar.set)

    root.mainloop()

def on_frame_configure(canvas):
    """Update scroll region."""
    canvas.configure(scrollregion=canvas.bbox("all"))



ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (default), "green", "dark-blue"

root = ctk.CTk()
root.title("Excelicious")

frame = ctk.CTkFrame(root, width=400, height=200)
frame.pack(padx=30, pady=30)

label_title = ctk.CTkLabel(frame, text="Find Duplicate Barcodes", font=("Serif", 20))
label_title.pack(pady=7)


button_search = ctk.CTkButton(frame, text="Select Folder", command=DuplicateBarcodes)
button_search.pack(pady=10)
button = Button(root, text = 'Terminate all Excel processes', command = PythonExcelKiller)
button.pack()

root.mainloop()
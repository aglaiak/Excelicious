import importlib
import subprocess
import sys

def PythonExcelKiller():
    """ Function that closes all Excel Instances """
    import Scripts.CloseExcel  
    
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
from Scripts.DetectDuplicates import run_full_process
from Scripts.DetectInconsistencies import get_name_files, check_date
from Scripts.CloseExcel import close_excel

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

def fetch_folder_path():
    """Open a dialog to select a folder and return its path."""
    folder_path = filedialog.askdirectory()
    return folder_path

def handle_check_inconsistencies():
    """Handle the process of checking file name inconsistencies."""
    folder_path = fetch_folder_path()
    if not folder_path:
        messagebox.showwarning("No Folder Selected", "Please select a folder.")
        return

    print(f"Selected folder path: {folder_path}")  # Debugging print

    files = get_name_files(folder_path)
    print(f"Files found: {files}")  # Debugging print

    issues = check_date(files)
    print(f"Issues found: {issues}")  # Debugging print

    if issues:
        result_text = "Files with naming inconsistencies:\n\n"
        for file in issues:
            result_text += f"{file}\n"
        display_scrollable_message("Naming Inconsistencies", result_text)
    else:
        messagebox.showinfo("No Issues", "No naming inconsistencies found.")

def DuplicateBarcodes():
    import Scripts.DetectDuplicates

def on_frame_configure(canvas):
    """ Update scroll region."""
    canvas.configure(scrollregion=canvas.bbox("all"))

# Initialize CustomTkinter appearance

ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("dark-blue")  

root = ctk.CTk()
root.title("Excelicious")
root.geometry("300x300")  

frame = ctk.CTkFrame(root, width=700, height=700)
frame.pack(padx=15, pady=15)
frame1 = ctk.CTkFrame(root, width=600, height=800)
frame.pack(padx=15, pady=15)

label_title = ctk.CTkLabel(frame, text="Barcode Stuff", font=("Serif", 15))
label_title.pack(pady=10)

button_search = ctk.CTkButton(frame, text="Check for Duplicate Barcodes", command=run_full_process)
button_search.pack(pady=10)

button_inconsistencies = ctk.CTkButton(frame, text="Check for Naming Inconsistencies", command=handle_check_inconsistencies)
button_inconsistencies.pack(pady=10)

button_kill_excel = ctk.CTkButton(
    root, 
    text="Terminate all Excel processes",
    command=close_excel,
    fg_color="#FF0000",  
    hover_color="#CC0000"  
)
button_kill_excel.pack(pady=10)

root.mainloop()
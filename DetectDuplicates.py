import os
import sys
import subprocess
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
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, Button

def fetch_folder_path():
    """Open a dialog to select a folder and return its path."""
    folder_path = filedialog.askdirectory()
    return folder_path

def find_unique_barcodes(folder_path):
    """Find barcodes that are unique across all Excel files and sheets."""
    all_barcodes = set()
    unique_barcodes = set()

    # Dictionary to count occurrences of each barcode across all files
    barcode_count = {}

    # Traverse all files in the folder
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            if file_name.endswith('.xlsx'):
                file_path = os.path.join(root, file_name)
                try:
                    # Read all sheets from the excel file
                    xl = pd.ExcelFile(file_path)
                    for sheet_name in xl.sheet_names:
                        # Ensure we are reading from the 'reference' sheet
                        if sheet_name.lower() == 'reference':
                            df = xl.parse(sheet_name)
                            # Ensure column names are stripped of whitespace and converted to lowercase
                            df.columns = df.columns.str.strip().str.lower()
                            print(f"Processing file: {file_name}, Sheet: {sheet_name}")
                            print("Columns found:", df.columns)  # Debugging print
                            
                            # Track encountered barcodes in this sheet
                            encountered_barcodes = set()
                            # Iterate through each row
                            for idx, row in df.iterrows():
                                if row['status'] == 'edited':
                                    barcode = row['barcode']
                                    all_barcodes.add(barcode)
                                    encountered_barcodes.add(barcode)

                            # Update barcode_count for this file
                            for barcode in encountered_barcodes:
                                if barcode in barcode_count:
                                    barcode_count[barcode].add(file_name)
                                else:
                                    barcode_count[barcode] = {file_name}

                except KeyError as e:
                    messagebox.showerror("Error", f"Error processing {file_path}: Column not found - {e}")
                    print(f"Error processing {file_path}: Column not found - {e}")  # Debugging print
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading {file_path}: {e}")
                    print(f"Error reading {file_path}: {e}")  # Debugging print

    # Find unique barcodes
    for barcode, files in barcode_count.items():
        if len(files) == 1:  # Barcode found in only one file
            unique_barcodes.add(barcode)

    return unique_barcodes

def find_edited_values(folder_path, unique_barcodes):
    """Find the first 100 rows with 'edited' in column 'status' from each sheet 
    named 'reference' in each Excel file and extract columns 'number', 'barcode', and 'name'.
    Exclude unique barcodes that are found in other files."""
    edited_values = {}

    # Traverse all files in the folder
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            if file_name.endswith('.xlsx'):
                file_path = os.path.join(root, file_name)
                try:
                    # Read all sheets from the excel file
                    xl = pd.ExcelFile(file_path)
                    for sheet_name in xl.sheet_names:
                        # Ensure we are reading from the 'reference' sheet
                        if sheet_name.lower() == 'reference':
                            df = xl.parse(sheet_name)
                            # Ensure column names are stripped of whitespace and converted to lowercase
                            df.columns = df.columns.str.strip().str.lower()
                            print(f"Processing file: {file_name}, Sheet: {sheet_name}")
                            print("Columns found:", df.columns)  # Debugging print
                            
                            # List to hold concatenated values for this sheet with their original indexes
                            sheet_values = []
                            # Iterate through each row from bottom to top
                            for idx in range(len(df) - 1, -1, -1):
                                row = df.iloc[idx]
                                if row['status'] == 'edited':
                                    barcode = row['barcode']
                                    if barcode in unique_barcodes:
                                        continue  # Skip if barcode is unique across all files
                                    # Collect the row's columns 'number', 'barcode', and 'name' along with the original index
                                    row_values = f"{row['number']},{barcode},{row['name']}"
                                    sheet_values.append((idx, row_values))
                                    if len(sheet_values) >= 100:
                                        break
                            if sheet_values:
                                if file_name not in edited_values:
                                    edited_values[file_name] = {}
                                edited_values[file_name][sheet_name] = sheet_values

                except KeyError as e:
                    messagebox.showerror("Error", f"Error processing {file_path}: Column not found - {e}")
                    print(f"Error processing {file_path}: Column not found - {e}")  # Debugging print
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading {file_path}: {e}")
                    print(f"Error reading {file_path}: {e}")  # Debugging print

    return edited_values

def compare_values(edited_values):
    """Compare concatenated values across files to detect potential duplicates."""
    duplicates = {}  # Dictionary to store potential duplicates

    # Create a dictionary to store concatenated values for each file
    file_concatenated_values = {}

    # Populate file_concatenated_values dictionary
    for file_name, sheets_data in edited_values.items():
        file_values = {}
        for sheet_name, values in sheets_data.items():
            for idx, value in values:
                if idx not in file_values:
                    file_values[idx] = []
                file_values[idx].append(value)
        file_concatenated_values[file_name] = file_values

    # Compare values across different files based on their original indexes
    all_files = list(file_concatenated_values.keys())

    for i in range(len(all_files)):
        for j in range(i + 1, len(all_files)):
            file_name1 = all_files[i]
            file_name2 = all_files[j]
            values1 = file_concatenated_values[file_name1]
            values2 = file_concatenated_values[file_name2]

            # Find common indexes
            common_indexes = set(values1.keys()).intersection(set(values2.keys()))

            # Compare values at common indexes
            for idx in common_indexes:
                value1 = values1[idx][0]  # Assuming there's only one value per index per file
                value2 = values2[idx][0]  # Assuming there's only one value per index per file
                if value1 != value2:
                    if value1 not in duplicates:
                        duplicates[value1] = set()
                    duplicates[value1].add(file_name1)
                    duplicates[value1].add(file_name2)

                    if value2 not in duplicates:
                        duplicates[value2] = set()
                    duplicates[value2].add(file_name1)
                    duplicates[value2].add(file_name2)

    # Prepare results for notification
    if duplicates:
        result_text = "Potential duplicates found:\n\n"
        duplicate_count = len(duplicates)
        for value, files in duplicates.items():
            result_text += f"Value: {value}\nFiles: {', '.join(files)}\n\n"
       
        result_text += f"Total number of duplicate barcodes: {duplicate_count}\n"
        # Display results in a scrollable messagebox
        display_scrollable_message("Potential Duplicates", result_text)
    else:
        messagebox.showinfo("No Duplicates", "No potential duplicates found.")

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

def search_edited_barcodes():
    """ Go through each barcode of an excel file, look for values that are flagged as 'edited'
    and save the first 100 by moving upstream"""
    folder_path = fetch_folder_path()
    if not folder_path:
        messagebox.showwarning("No Folder Selected", "Please select a folder.")
        return
    
    unique_barcodes = find_unique_barcodes(folder_path)
    if unique_barcodes:
        print("Unique Barcodes:")
        print(unique_barcodes)  # Debugging print to show unique barcodes
        
    edited_values = find_edited_values(folder_path, unique_barcodes)
    
    if edited_values:
        compare_values(edited_values)


ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (default), "green", "dark-blue"

root = ctk.CTk()
root.title("Histo Tool")

frame = ctk.CTkFrame(root, width=400, height=200)
frame.pack(padx=30, pady=30)

label_title = ctk.CTkLabel(frame, text="Find Duplicate Barcodes", font=("Serif", 20))
label_title.pack(pady=7)


button_search = ctk.CTkButton(frame, text="Select Folder", command=search_edited_barcodes)
button_search.pack(pady=10)

root.mainloop()

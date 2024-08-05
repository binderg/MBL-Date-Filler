import os
import re
from docx import Document
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import sys

# Global variables for placeholder strings
DATE_1_PATTERN = re.compile(r'_{1,}, 20_{1,}')
DATE_2_PATTERN = re.compile(r'_{1,} day of _{1,}, 20_{1,}')
DATE_3_PATTERN = re.compile(r'_{1,} day of _{1,}, in the year 20_{1,}')

def replace_dates_in_docs(src_folder, dest_folder, date1, date2, date3):
    for filename in os.listdir(src_folder):
        if filename.endswith(".docx"):
            src_path = os.path.join(src_folder, filename)
            dest_path = os.path.join(dest_folder, filename)
            doc = Document(src_path)
            for paragraph in doc.paragraphs:
                paragraph_text = paragraph.text

                # Replace using regex patterns, applied in specific order (because regex 1 is included in 2)
                if DATE_2_PATTERN.search(paragraph_text):
                    paragraph_text = DATE_2_PATTERN.sub(date2, paragraph_text)
                elif DATE_3_PATTERN.search(paragraph_text):
                    paragraph_text = DATE_3_PATTERN.sub(date3, paragraph_text)
                else:
                    paragraph_text = DATE_1_PATTERN.sub(date1, paragraph_text)

                # Update the paragraph text
                paragraph.text = paragraph_text
            doc.save(dest_path)


def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_var.set(folder_selected)

def start_replacement():
    src_folder = folder_var.get()
    if not src_folder:
        messagebox.showerror("Error", "Please select a folder.")
        return

    selected_date_str = date_entry.entry.get()  # Get the selected date as string
    if not selected_date_str:
        messagebox.showerror("Error", "Please select a date.")
        return

    try:
        selected_date = datetime.strptime(selected_date_str, '%m/%d/%Y')  # Convert string to datetime
    except ValueError:
        messagebox.showerror("Error", "Invalid date format.")
        return

    # Format dates
    date1 = selected_date.strftime("%B %d, %Y")
    day_suffix = 'th' if 11 <= selected_date.day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(selected_date.day % 10, 'th')
    date2 = f"{selected_date.day}{day_suffix} day of {selected_date.strftime('%B, %Y')}"
    date3 = f"{selected_date.day}{day_suffix} day of {selected_date.strftime('%B, in the year %Y')}"

    # Display variables
    date1_var.set(f"{DATE_1_PATTERN.pattern}: {date1}")
    date2_var.set(f"{DATE_2_PATTERN.pattern}: {date2}")
    date3_var.set(f"{DATE_3_PATTERN.pattern}: {date3}")

    # Create a new folder named with the current date
    date_folder_name = selected_date.strftime('%m-%d-%Y')
    dest_folder = os.path.join(src_folder, date_folder_name)
    os.makedirs(dest_folder, exist_ok=True)

    replace_dates_in_docs(src_folder, dest_folder, date1, date2, date3)
    messagebox.showinfo("Success", "Dates replaced and files saved.")

def on_closing():
    # Deactivate the virtual environment if needed
    # This is usually done via command-line, but here we just exit the script
    # Depending on your needs, you might run a shell command here to deactivate the venv
    print("Closing program...")
    sys.exit()

# Create GUI
root = ttk.Window(themename="cosmo")
root.title("Date Replacer")

folder_var = StringVar()
date1_var = StringVar()
date2_var = StringVar()
date3_var = StringVar()

# Display placeholders and example dates at the top
ttk.Label(root, text=f"{DATE_1_PATTERN.pattern}: Example: {datetime.now().strftime('%B %d, %Y')}", font=("Helvetica", 12, "bold")).grid(row=0, column=0, columnspan=3, pady=10)
ttk.Label(root, text=f"{DATE_2_PATTERN.pattern}: Example: {datetime.now().day}th day of {datetime.now().strftime('%B, %Y')}", font=("Helvetica", 12, "bold")).grid(row=1, column=0, columnspan=3, pady=5)
ttk.Label(root, text=f"{DATE_3_PATTERN.pattern}: Example: {datetime.now().day}th day of {datetime.now().strftime('%B, in the year %Y')}", font=("Helvetica", 12, "bold")).grid(row=2, column=0, columnspan=3, pady=5)

ttk.Label(root, text="Select the folder containing .docx files:", font=("Helvetica", 12)).grid(row=3, column=0, columnspan=3, pady=10)
ttk.Entry(root, textvariable=folder_var, width=50).grid(row=4, column=0, columnspan=2, padx=10)
ttk.Button(root, text="Browse", command=browse_folder).grid(row=4, column=2, padx=10)

ttk.Label(root, text="Select the date:", font=("Helvetica", 12)).grid(row=5, column=0, columnspan=3, pady=10)
date_entry = ttk.DateEntry(root, bootstyle='success')
date_entry.grid(row=6, column=0, columnspan=3, padx=10)

ttk.Button(root, text="Add Dates", command=start_replacement).grid(row=10, column=0, columnspan=3, pady=20)

# Bind the closing event to the on_closing function
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()

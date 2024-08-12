import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import time

CHUNK_SIZE = 5  # Reduced chunk size to process fewer lines at a time
DELAY = 0.1  # Delay in seconds between processing chunks

def browse_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, filepath)

def add_default_replacements():
    entry_replacements.delete("1.0", tk.END)
    for i in range(1, 91):
        old_text = f"SAW{i:02}"
        entry_replacements.insert(tk.END, f"{old_text} -> \n")

def load_presentation():
    ppt_path = entry_file_path.get()
    if not ppt_path:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return None
    return Presentation(ppt_path)

def apply_saw_replacements(prs):
    try:
        replacement_lines = entry_replacements.get("1.0", tk.END).strip().splitlines()
        replacements.clear()

        for i, line in enumerate(replacement_lines):
            old_text = f"SAW{i+1:02}"
            if '->' in line:
                _, new_text = line.split('->')
                new_text = new_text.strip()
                replacements[old_text] = new_text
            else:
                replacements[old_text] = line.strip()

        for slide in prs.slides:
            for shape in slide.shapes:
                process_shape(shape)
        messagebox.showinfo("Success", "SAW replacements applied.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def process_shape(shape):
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            process_shape(s)  # Recursive call to handle nested groups
    elif shape.has_text_frame:
        text_frame = shape.text_frame
        replace_text_in_text_frame(text_frame)

    if shape.has_table:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                text_fra

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
import pandas as pd
import os
import threading
import time

CHUNK_SIZE = 5  # Adjusted chunk size
DELAY = 0.1  # Delay between processing chunks

def browse_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, filepath)

def browse_excel_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")]
    )
    entry_excel_path.delete(0, tk.END)
    entry_excel_path.insert(0, filepath)

def load_presentation():
    ppt_path = entry_file_path.get()
    if not ppt_path:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return None
    return Presentation(ppt_path)

def load_excel_data():
    excel_path = entry_excel_path.get()
    if not excel_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return None, None
    try:
        df_servers = pd.read_excel(excel_path, sheet_name='Servers Part of Report Cycle')
        df_format = pd.read_excel(excel_path, sheet_name='Format Box')
        return df_servers, df_format
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the Excel file: {e}")
        return None, None

def extract_and_apply_replacements():
    try:
        prs = load_presentation()
        if not prs:
            return
        
        df_servers, df_format = load_excel_data()
        if df_servers is None or df_format is None:
            return

        for _, server_row in df_servers.iterrows():
            report_name = server_row['Report Name']
            hostname = server_row['Hostname']

            combined_replacements = {
                '31.77%': str(df_format.loc[0, 'CPU Utilization']),
                '53.07%': str(df_format.loc[0, 'Memory Utilization']),
                '83.07%': str(df_format.loc[0, 'Disk Utilization'])
            }
            
            # Filter data for the current report_name
            filtered_df = df_servers[df_servers['Report Name'] == report_name]
            
            for _, row in filtered_df.iterrows():
                if not row['Report Name'] or not row['Hostname']:
                    continue

                new_ppt = prs
                apply_draft_replacement(new_ppt, report_name)
                apply_saw_replacements(new_ppt, hostname)
                apply_combined_replacements(new_ppt, combined_replacements)
                save_path = f"{os.path.splitext(entry_file_path.get())[0]}_{report_name}.pptx"
                new_ppt.save(save_path)

        messagebox.showinfo("Success", "All replacements applied and files saved.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def apply_draft_replacement(prs, new_text):
    try:
        first_slide = prs.slides[0]  # Get the first slide
        for shape in first_slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "Draft Template" in run.text:
                            run.text = run.text.replace("Draft Template", new_text)
                            run.font.color.rgb = RGBColor(0, 0, 0)  # Set text color to black
        messagebox.showinfo("Success", f"'Draft Template' replaced with '{new_text}' on the first slide.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def apply_saw_replacements(prs, hostname):
    try:
        replacements = {"SAW01": hostname}
        for slide in prs.slides:
            for shape in slide.shapes:
                process_shape(shape, replacements)
        messagebox.showinfo("Success", "Hostnames->SAW replacements applied.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def process_shape(shape, replacements):
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            process_shape(s, replacements)  # Recursive call to handle nested groups
    elif shape.has_text_frame:
        text_frame = shape.text_frame
        replace_text_in_text_frame(text_frame, replacements)

    if shape.has_table:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                text_frame = cell.text_frame
                replace_text_in_text_frame(text_frame, replacements)

def replace_text_in_text_frame(text_frame, replacements):
    if text_frame is not None:
        for paragraph in text_frame.paragraphs:
            full_text = ''.join([run.text for run in paragraph.runs])  # Combine all runs' text
            for old_text, new_text in replacements.items():
                if old_text in full_text:
                    full_text = full_text.replace(old_text, new_text)
                    for run in paragraph.runs:
                        run.text = ''  # Clear existing text
                    paragraph.runs[0].text = full_text  # Set the first run to the new text

def apply_combined_replacements(prs, replacements):
    try:
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    for row in table.rows:
                        for cell in row.cells:
                            text_frame = cell.text_frame
                            if text_frame is not None:
                                for paragraph in text_frame.paragraphs:
                                    full_text = ''.join([run.text for run in paragraph.runs])
                                    for search_value, replacement_value in replacements.items():
                                        if search_value in full_text:
                                            full_text = full_text.replace(search_value, replacement_value)
                                            for run in paragraph.runs:
                                                run.text = ''  # Clear existing text
                                            paragraph.runs[0].text = full_text  # Set the first run to the new text

        messagebox.showinfo("Success", "Combined replacements applied.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def start_threaded_processing():
    threading.Thread(target=extract_and_apply_replacements).start()

# Set up the main window
root = tk.Tk()
root.title("PowerPoint Report Text Replacer")

# Set up the notebook (tabs)
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, padx=10, pady=10)

# First tab for SAW replacements
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Hostnames->SAW Replacements")

# File selection for SAW Replacements
tk.Label(tab1, text="Select PowerPoint File:").grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(tab1, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
tk.Button(tab1, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

tk.Label(tab1, text="Select Excel File:").grid(row=1, column=0, padx=10, pady=5)
entry_excel_path = tk.Entry(tab1, width=50)
entry_excel_path.grid(row=1, column=1, padx=10, pady=5)
tk.Button(tab1, text="Browse", command=browse_excel_file).grid(row=1, column=2, padx=10, pady=5)

# Remove the input fields and buttons for individual replacements
tk.Label(tab1, text="Replacement Values for\n CPU, Memory and Disk Utilization\n (three per line, separated by spaces):").grid(row=2, column=0, padx=10, pady=5)
entry_combined = tk.Text(tab1, width=50, height=10)
entry_combined.grid(row=2, column=1, padx=10, pady=5)

tk.Button(tab1, text="Apply Combined Replacements", command=start_threaded_processing).grid(row=3, column=1, padx=10, pady=20)

# Progress label for feedback
progress_label = tk.Label(root, text="")
progress_label.grid(row=4, column=0, padx=10, pady=5)

# Start the Tkinter main loop
root.mainloop()

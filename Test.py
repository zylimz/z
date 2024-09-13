import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
import pandas as pd

# Helper function to run a function in a separate thread
def run_in_background(func):
    threading.Thread(target=func).start()

# Load the PowerPoint presentation
def load_presentation():
    file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return None
    try:
        return Presentation(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load presentation: {e}")
        return None

# Save the modified PowerPoint presentation
def save_changes(prs):
    save_path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])
    if not save_path:
        messagebox.showerror("Error", "No save location selected.")
        return
    try:
        prs.save(save_path)
        messagebox.showinfo("Success", "Presentation saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save presentation: {e}")

# Function to replace "Draft Template" text on the first slide
def replace_draft_template(prs, new_text):
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

# Apply the replacement for "Draft Template"
def apply_draft_replacement():
    prs = load_presentation()
    if prs:
        excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not excel_path:
            messagebox.showerror("Error", "No Excel file selected.")
            return
        
        df = pd.read_excel(excel_path, sheet_name=0)
        for name in df['name'].unique():
            replace_draft_template(prs, name)
            save_changes(prs)

# Function to apply SAW replacements
def apply_saw_replacements(prs):
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        messagebox.showerror("Error", "No Excel file selected.")
        return
    
    df = pd.read_excel(excel_path, sheet_name=0)
    for name in df['name'].unique():
        saw_replacements = df[df['name'] == name]['hostname'].tolist()
        replacements = {f"SAW{str(i+1).zfill(2)}": saw_replacements[i] for i in range(len(saw_replacements))}

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    replace_text_in_text_frame(shape.text_frame, replacements)
                elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    for grouped_shape in shape.shapes:
                        if grouped_shape.has_text_frame:
                            replace_text_in_text_frame(grouped_shape.text_frame, replacements)

# Function to replace text in a text frame
def replace_text_in_text_frame(text_frame, replacements):
    if text_frame is not None:
        for paragraph in text_frame.paragraphs:
            full_text = ''.join([run.text for run in paragraph.runs])  # Combine all runs' text
            for old_text, new_text in replacements.items():
                if old_text in full_text:
                    # Replace the text in the combined string
                    full_text = full_text.replace(old_text, new_text)

                    # Clear the paragraph runs and create a single run with the replaced text
                    for run in paragraph.runs:
                        run.text = ''  # Clear existing text
                    paragraph.runs[0].text = full_text  # Set the first run to the new text

# Function to apply combined replacements
def apply_combined_replacements(prs):
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        messagebox.showerror("Error", "No Excel file selected.")
        return
    
    df = pd.read_excel(excel_path, sheet_name=1)
    for name in df['name'].unique():
        combined_replacements = df[df['name'] == name].iloc[0].to_dict()
        replacements = {
            "31.77%": combined_replacements['memory utilisation'],
            "53.07%": combined_replacements['CPU utilisation'],
            "83.07%": combined_replacements['disk utilisation'],
        }

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    replace_text_in_text_frame(shape.text_frame, replacements)
                elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    for grouped_shape in shape.shapes:
                        if grouped_shape.has_text_frame:
                            replace_text_in_text_frame(grouped_shape.text_frame, replacements)

# Apply all replacements in a separate thread
def apply_all_replacements():
    prs = load_presentation()
    if prs:
        apply_draft_replacement()
        apply_saw_replacements(prs)
        apply_combined_replacements(prs)
        save_changes(prs)
        messagebox.showinfo("Success", "All replacements have been applied and saved successfully.")

def apply_all_replacements_threaded():
    run_in_background(apply_all_replacements)

# Set up the main window
root = tk.Tk()
root.title("PowerPoint Text Replacer")

# Set up the notebook (tabs)
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, padx=10, pady=10)

# New tab for "Draft Template" replacement
tab_draft = ttk.Frame(notebook)
notebook.add(tab_draft, text="Replace 'Draft Template'")

tk.Label(tab_draft, text="Replacement Text:").grid(row=0, column=0, padx=10, pady=5)
entry_draft = tk.Entry(tab_draft, width=50)
entry_draft.grid(row=0, column=1, padx=10, pady=5)

tk.Button(tab_draft, text="Replace and Save", command=lambda: run_in_background(apply_draft_replacement)).grid(row=1, column=1, padx=10, pady=20)

# First tab for SAW replacements
tab_saw = ttk.Frame(notebook)
notebook.add(tab_saw, text="SAW Replacements")

tk.Label(tab_saw, text="SAW Replacements (One per line):").grid(row=0, column=0, padx=10, pady=5)
text_saw = tk.Text(tab_saw, height=10, width=50)
text_saw.grid(row=1, column=0, padx=10, pady=5)

# Second tab for Combined Replacements
tab_combined = ttk.Frame(notebook)
notebook.add(tab_combined, text="Combined Replacements")

tk.Label(tab_combined, text="Combined Replacements (Three values per line):").grid(row=0, column=0, padx=10, pady=5)
text_combined = tk.Text(tab_combined, height=10, width=50)
text_combined.grid(row=1, column=0, padx=10, pady=5)

# Apply replacements button for all replacements
tk.Button(root, text="Apply All Replacements and Save", command=apply_all_replacements_threaded).grid(row=3, column=0, padx=10, pady=20)

# Start the Tkinter main loop
root.mainloop()

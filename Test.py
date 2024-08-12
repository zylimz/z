import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

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

def apply_saw_replacements():
    ppt_path = entry_file_path.get()
    if not ppt_path:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return

    try:
        global prs
        prs = Presentation(ppt_path)
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

        messagebox.showinfo("Success", "SAW replacements applied. Don't forget to save your changes!")
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
                text_frame = cell.text_frame
                replace_text_in_text_frame(text_frame)

def replace_text_in_text_frame(text_frame):
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

def search_and_replace_value(search_value, replacement_value):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        if search_value in cell.text:
                            text_frame = cell.text_frame
                            for paragraph in text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if search_value in run.text:
                                        # Get the index of the text to replace
                                        start = run.text.find(search_value)
                                        end = start + len(search_value)
                                        # Replace the text while preserving formatting
                                        run.text = run.text[:start] + replacement_value + run.text[end:]
                                        return  # Exit after the first match per slide

def apply_combined_replacements():
    ppt_path = entry_file_path.get()
    if not ppt_path:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return

    try:
        global prs
        prs = Presentation(ppt_path)
        replacement_lines = entry_combined.get("1.0", tk.END).strip().splitlines()

        for line in replacement_lines:
            try:
                value_31, value_53, value_83 = line.split()
                # Perform the replacements one at a time
                search_and_replace_value("31.77%", value_31)
                search_and_replace_value("53.07%", value_53)
                search_and_replace_value("83.07%", value_83)
            except ValueError:
                messagebox.showerror("Error", "Each line must contain exactly three values separated by spaces.")
                return

        messagebox.showinfo("Success", "Combined replacements applied. Don't forget to save your changes!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def save_changes():
    if 'prs' in globals():
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if save_path:
            prs.save(save_path)
            messagebox.showinfo("Success", f"Changes saved to {save_path}")
    else:
        messagebox.showerror("Error", "No changes to save. Please apply replacements first.")

# Initialize the replacements dictionary
replacements = {}

# Set up the main window
root = tk.Tk()
root.title("PowerPoint Text Replacer")

# Set up the notebook (tabs)
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, padx=10, pady=10)

# First tab for SAW replacements
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="SAW Replacements")

# File selection for SAW Replacements
tk.Label(tab1, text="Select PowerPoint File:").grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(tab1, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
tk.Button(tab1, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

# Default replacement input for SAW Replacements
tk.Button(tab1, text="Load SAW01 to SAW90", command=add_default_replacements).grid(row=1, column=1, padx=10, pady=5)

# Replacement input area for SAW Replacements
tk.Label(tab1, text="Replacement Pairs (one per line):").grid(row=2, column=0, padx=10, pady=5)
entry_replacements = tk.Text(tab1, width=50, height=20)
entry_replacements.grid(row=2, column=1, padx=10, pady=5)

# Apply replacements button for SAW Replacements
tk.Button(tab1, text="Apply SAW Replacements", command=apply_saw_replacements).grid(row=3, column=1, padx=10, pady=20)

# Save changes button
tk.Button(tab1, text="Save Changes", command=save_changes).grid(row=4, column=1, padx=10, pady=20)

# Second tab for Combined Replacements
tab_combined = ttk.Frame(notebook)
notebook.add(tab_combined, text="Combined Replacements")

# Combined replacement input
tk.Label(tab_combined, text="Replacement Values (three per line, separated by spaces):").grid(row=0, column=0, padx=10, pady=5)
entry_combined = tk.Text(tab_combined, width=50, height=20)
entry_combined.grid(row=0, column=1, padx=10, pady=5)

# Apply combined replacements button
tk.Button(tab_combined, text="Apply Combined Replacements", command=apply_combined_replacements).grid(row=1, column=1, padx=10, pady=20)

# Save changes button
tk.Button(tab_combined, text="Save Changes", command=save_changes).grid(row=2, column=1, padx=10, pady=20)

# Start the Tkinter main loop
root.mainloop()

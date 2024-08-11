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

def apply_replacements():
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

        messagebox.showinfo("Success", "Replacements applied. Don't forget to save your changes!")
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

def search_and_replace_value(search_value, entry_widget):
    ppt_path = entry_file_path.get()
    if not ppt_path:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return

    try:
        global prs
        replacement_values = entry_widget.get("1.0", tk.END).strip().splitlines()

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
                                            run.text = run.text[:start] + replacement_values.pop(0) if replacement_values else run.text[:start]
                                            run.text += run.text[end:]
                                            # Update remaining replacements in the input field
                                            entry_widget.delete("1.0", tk.END)
                                            entry_widget.insert(tk.END, "\n".join(replacement_values))
                                            break
                                    else:
                                        continue
                                    break
                            if not replacement_values:
                                break
                        else:
                            continue
                        break
                if not replacement_values:
                    break

        messagebox.showinfo("Success", "Search and replacement applied. Don't forget to save your changes!")
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
tk.Button(tab1, text="Apply Replacements", command=apply_replacements).grid(row=3, column=1, padx=10, pady=20)

# Save changes button
tk.Button(tab1, text="Save Changes", command=save_changes).grid(row=4, column=1, padx=10, pady=20)

# Second tab for Search and Replace "31.77%"
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Replace 31.77%")

# Search and replace input
tk.Label(tab2, text="Replacement Values (one per line):").grid(row=0, column=0, padx=10, pady=5)
entry_search_replace_31 = tk.Text(tab2, width=50, height=20)
entry_search_replace_31.grid(row=0, column=1, padx=10, pady=5)

# Apply search and replace button
tk.Button(tab2, text="Apply Search and Replace", command=lambda: search_and_replace_value("31.77%", entry_search_replace_31)).grid(row=1, column=1, padx=10, pady=20)

# Save changes button
tk.Button(tab2, text="Save Changes", command=save_changes).grid(row=2, column=1, padx=10, pady=20)

# Third tab for Search and Replace "53.07%"
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="Replace 53.07%")

# Search and replace input
tk.Label(tab3, text="Replacement Values (one per line):").grid(row=0, column=0, padx=10, pady=5)
entry_search_replace_53 = tk.Text(tab3, width=50, height=20)
entry_search_replace_53.grid(row=0, column=1, padx=10, pady=5)

# Apply search and replace button
tk.Button(tab3, text="Apply Search and Replace", command=lambda: search_and_replace_value("53.07%", entry_search_replace_53)).grid(row=1, column=1, padx=10, pady=20)

# Save changes button
tk.Button(tab3, text="Save Changes", command=save_changes).grid(row=2, column=1, padx=10, pady=20)

# Fourth tab for Search and Replace "83.07%"
tab4 = ttk.Frame(notebook)
notebook.add(tab4, text="Replace 83.07%")

# Search and replace input
tk.Label(tab4, text="Replacement Values (one per line):").grid(row=0, column=0, padx=10, pady=5)
entry_search_replace_83 = tk.Text(tab4, width=50, height=20)
entry_search_replace_83.grid(row=0, column=1, padx=10, pady=5)

# Apply search and replace button
tk.Button(tab4, text="Apply Search and Replace", command=lambda: search_and_replace_value("83.07%", entry_search_replace_83)).grid(row=1, column=1, padx=10, pady=20)

# Save changes button
tk.Button(tab4, text="Save Changes", command=save_changes).grid(row=2, column=1, padx=10, pady=20)

# Start the Tkinter main loop
root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from tkinter import ttk

# Global variable to store the PowerPoint presentation
prs = None

def browse_file():
    global prs
    filepath = filedialog.askopenfilename(
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, filepath)
    if filepath:
        prs = Presentation(filepath)

def add_default_replacements():
    entry_replacements.delete("1.0", tk.END)
    for i in range(1, 91):
        old_text = f"SAW{i:02}"
        entry_replacements.insert(tk.END, f"{old_text} -> \n")

def apply_saw_replacements():
    if prs is None:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return

    try:
        replacement_lines = entry_replacements.get("1.0", tk.END).strip().splitlines()
        replacements.clear()

        for i, line in enumerate(replacement_lines):
            old_text = f"SAW{i+1:02}"
            if '->' in line:
                _, new_text = line.split('->', 1)  # Limit split to 1 occurrence
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
            process_shape(s)
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
            full_text = ''.join([run.text for run in paragraph.runs])
            for old_text, new_text in replacements.items():
                if old_text in full_text:
                    full_text = full_text.replace(old_text, new_text)
                    for run in paragraph.runs:
                        run.text = ''
                    paragraph.runs[0].text = full_text

def apply_three_value_replacements():
    if prs is None:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return

    try:
        replacement_lines = entry_three_value_replacements.get("1.0", tk.END).strip().splitlines()
        replacement_pairs = []

        for line in replacement_lines:
            values = line.split()
            if len(values) == 3:
                replacement_pairs.append(values)
            else:
                messagebox.showerror("Error", "Each line must contain exactly 3 values.")
                return

        # Ensure we don't go out of bounds if there are more slides than replacement pairs
        slide_index = 0
        for slide in prs.slides:
            if slide_index < len(replacement_pairs):
                for shape in slide.shapes:
                    process_three_value_shape(shape, replacement_pairs[slide_index])
                slide_index += 1

        messagebox.showinfo("Success", "Three-value replacements applied.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def process_three_value_shape(shape, replacement_values):
    placeholders = ["30.02%", "15.34%", "83.46%"]  # The actual placeholders to be replaced
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            process_three_value_shape(s, replacement_values)
    elif shape.has_text_frame:
        text_frame = shape.text_frame
        replace_three_values_in_text_frame(text_frame, placeholders, replacement_values)

def replace_three_values_in_text_frame(text_frame, placeholders, replacements):
    if text_frame is not None:
        for paragraph in text_frame.paragraphs:
            full_text = ''.join([run.text for run in paragraph.runs])
            for i, placeholder in enumerate(placeholders):
                if i < len(replacements) and placeholder in full_text:
                    full_text = full_text.replace(placeholder, replacements[i])
            for run in paragraph.runs:
                run.text = ''
            paragraph.runs[0].text = full_text

def save_presentation():
    if prs is None:
        messagebox.showerror("Error", "No presentation loaded to save.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")]
    )
    if save_path:
        try:
            prs.save(save_path)
            messagebox.showinfo("Success", f"Presentation saved to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving: {e}")

# Initialize the replacements dictionary
replacements = {}

# Set up the main window
root = tk.Tk()
root.title("PowerPoint Text Replacer")

# Create a tabbed interface
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill='both')

# Tab 1: SAW replacements
frame_saw = ttk.Frame(notebook)
notebook.add(frame_saw, text="SAW Replacements")

# File selection
tk.Label(frame_saw, text="Select PowerPoint File:").grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(frame_saw, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
tk.Button(frame_saw, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

# Default replacement input
tk.Button(frame_saw, text="Load SAW01 to SAW90", command=add_default_replacements).grid(row=1, column=1, padx=10, pady=5)

# Replacement input area
tk.Label(frame_saw, text="Replacement Pairs (one per line):").grid(row=2, column=0, padx=10, pady=5)
entry_replacements = tk.Text(frame_saw, width=50, height=20)
entry_replacements.grid(row=2, column=1, padx=10, pady=5)

# Apply replacements button
tk.Button(frame_saw, text="Apply SAW Replacements", command=apply_saw_replacements).grid(row=3, column=1, padx=10, pady=20)

# Tab 2: Three-value replacements
frame_three_values = ttk.Frame(notebook)
notebook.add(frame_three_values, text="Three-Value Replacements")

# Replacement input area for three-value replacements
tk.Label(frame_three_values, text="Replacement Values (3 per line, separated by space):").grid(row=0, column=0, padx=10, pady=5)
entry_three_value_replacements = tk.Text(frame_three_values, width=50, height=20)
entry_three_value_replacements.grid(row=1, column=0, padx=10, pady=5)

# Apply three-value replacements button
tk.Button(frame_three_values, text="Apply Three-Value Replacements", command=apply_three_value_replacements).grid(row=2, column=0, padx=10, pady=20)

# Separate Save button
tk.Button(root, text="Save Presentation", command=save_presentation).pack(pady=20)

# Start the GUI loop
root.mainloop()

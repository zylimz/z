import tkinter as tk
from tkinter import filedialog, messagebox
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
                # Handle the case where input line is just the replacement value
                replacements[old_text] = line.strip()

        for slide in prs.slides:
            for shape in slide.shapes:
                process_shape(shape)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if save_path:
            prs.save(save_path)
            messagebox.showinfo("Success", f"Replacements applied and saved to {save_path}")
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
            for run in paragraph.runs:
                for old_text, new_text in replacements.items():
                    if old_text in run.text:
                        print(f"Replacing '{old_text}' with '{new_text}' in: {run.text}")  # Debug output
                        run.text = run.text.replace(old_text, new_text)

# Initialize the replacements dictionary
replacements = {}

# Set up the main window
root = tk.Tk()
root.title("PowerPoint Text Replacer")

# File selection
tk.Label(root, text="Select PowerPoint File:").grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

# Default replacement input
tk.Button(root, text="Load SAW01 to SAW90", command=add_default_replacements).grid(row=1, column=1, padx=10, pady=5)

# Replacement input area
tk.Label(root, text="Replacement Pairs (one per line):").grid(row=2, column=0, padx=10, pady=5)
entry_replacements = tk.Text(root, width=50, height=20)
entry_replacements.grid(row=2, column=1, padx=10, pady=5)

# Apply replacements button
tk.Button(root, text="Apply Replacements", command=apply_replacements).grid(row=3, column=1, padx=10, pady=20)

# Start the GUI loop
root.mainloop()

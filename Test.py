import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation

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
        replacements = {}
        for line in replacement_lines:
            if '->' in line:
                old_text, new_text = line.split('->')
                old_text, new_text = old_text.strip(), new_text.strip()
                replacements[old_text] = new_text

        def replace_text_in_shape(shape):
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        for old_text, new_text in replacements.items():
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
                for shape in text_frame.shapes:
                    replace_text_in_shape(shape)  # Recursive call to handle nested shapes

        def replace_text_in_table(table):
            for row in table.rows:
                for cell in row.cells:
                    text_frame = cell.text_frame
                    replace_text_in_shape(cell)  # Call to handle text in nested shapes
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            for old_text, new_text in replacements.items():
                                if old_text in run.text:
                                    run.text = run.text.replace(old_text, new_text)

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    replace_text_in_shape(shape)
                if shape.has_table:
                    replace_text_in_table(shape.table)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if save_path:
            prs.save(save_path)
            messagebox.showinfo("Success", f"Replacements applied and saved to {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

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
tk.Label(root, text="Replacement Pairs (modify as needed):").grid(row=2, column=0, padx=10, pady=5)
entry_replacements = tk.Text(root, width=50, height=20)
entry_replacements.grid(row=2, column=1, padx=10, pady=5)

# Apply replacements button
tk.Button(root, text="Apply Replacements", command=apply_replacements).grid(row=3, column=1, padx=10, pady=20)

# Start the GUI loop
root.mainloop()

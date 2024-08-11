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
        prs = Presentation(ppt_path)
        replacement_lines = entry_replacements.get("1.0", tk.END).strip().splitlines()
        replacements.clear()

        for line in replacement_lines:
            if '->' in line:
                old_text, new_text = map(str.strip, line.split('->'))
                replacements[old_text] = new_text
            else:
                old_text = f"SAW{len(replacements) + 1:02}"
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
            process_shape(s)
    elif shape.has_text_frame:
        text_frame = shape.text_frame
        replace_text_in_text_frame(text_frame)

    if shape.has_table:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text in values_to_replace:
                    index = values_to_replace.index(cell_text)
                    cell.text = current_replacements[index]

def replace_text_in_text_frame(text_frame):
    if text_frame is not None:
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                for old_text, new_text in replacements.items():
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)

def apply_table_replacements():
    ppt_path = entry_file_path.get()
    if not ppt_path:
        messagebox.showerror("Error", "Please select a PowerPoint file.")
        return

    try:
        prs = Presentation(ppt_path)
        replacement_lines = entry_table_replacements.get("1.0", tk.END).strip().splitlines()
        replacement_values = [line.split() for line in replacement_lines if line.strip()]

        if not replacement_values:
            messagebox.showerror("Error", "No replacement values provided.")
            return

        global values_to_replace
        global current_replacements

        values_to_replace = ["31.77%", "53.07%", "83.07%"]
        replacement_index = 0

        for slide in prs.slides:
            if replacement_index >= len(replacement_values):
                break
            current_replacements = replacement_values[replacement_index]

            # Process shapes on the current slide
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    for row in table.rows:
                        for cell in row.cells:
                            cell_text = cell.text.strip()
                            if cell_text in values_to_replace:
                                index = values_to_replace.index(cell_text)
                                cell.text = current_replacements[index]
            
            replacement_index += 1

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if save_path:
            prs.save(save_path)
            messagebox.showinfo("Success", f"Table replacements applied and saved to {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Initialize the replacements dictionary
replacements = {}

# Set up the main window
root = tk.Tk()
root.title("PowerPoint Text Replacer")

# Create the Notebook (tabs)
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, padx=10, pady=10)

# First tab - Text Replacement
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text='Text Replacement')

tk.Label(tab1, text="Select PowerPoint File:").grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(tab1, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
tk.Button(tab1, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

tk.Button(tab1, text="Load SAW01 to SAW90", command=add_default_replacements).grid(row=1, column=1, padx=10, pady=5)

tk.Label(tab1, text="Replacement Pairs (one per line):").grid(row=2, column=0, padx=10, pady=5)
entry_replacements = tk.Text(tab1, width=50, height=20)
entry_replacements.grid(row=2, column=1, padx=10, pady=5)

tk.Button(tab1, text="Apply Replacements", command=apply_replacements).grid(row=3, column=1, padx=10, pady=20)

# Second tab - Table Replacement
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text='Table Replacement')

tk.Label(tab2, text="Select PowerPoint File:").grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(tab2, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
tk.Button(tab2, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

tk.Label(tab2, text="Replacement Values (one line per set of slides):").grid(row=1, column=0, padx=10, pady=5)
entry_table_replacements = tk.Text(tab2, width=50, height=20)
entry_table_replacements.grid(row=1, column=1, padx=10, pady=5)

tk.Button(tab2, text="Apply Table Replacements", command=apply_table_replacements).grid(row=2, column=1, padx=10, pady=20)

# Start the GUI loop
root.mainloop()

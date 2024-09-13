import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor

class PowerPointReplacer:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Report Text Replacer")

        # Set up the notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, padx=10, pady=10)

        # First tab for file selection and processing
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="File Selection")

        # File selection
        tk.Label(self.tab1, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=5)
        self.entry_file_path = tk.Entry(self.tab1, width=50)
        self.entry_file_path.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(self.tab1, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=10, pady=5)

        # Progress label for feedback
        self.progress_label = tk.Label(root, text="")
        self.progress_label.grid(row=2, column=0, padx=10, pady=5)

        # Apply replacements button
        tk.Button(root, text="Apply Replacements and Save", command=self.start_processing).grid(row=3, column=0, padx=10, pady=20)

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.entry_file_path.delete(0, tk.END)
        self.entry_file_path.insert(0, filepath)

    def extract_data(self, df_report_cycle, df_format_box):
        report_names = df_report_cycle['Report Name'].unique()
        data_by_report = {}
        for report_name in report_names:
            report_data = df_report_cycle[df_report_cycle['Report Name'] == report_name]
            saw_values = report_data['Hostname'].astype(str).tolist()  # Ensure values are strings
            format_data = df_format_box[df_format_box['Report Name'] == report_name].iloc[0]

            # Convert values to percentage strings
            cpu_utilization = f"{float(format_data['CPU Utilization']) * 100:.2f}%"  # Convert to percentage
            memory_utilization = f"{float(format_data['Memory Utilization']) * 100:.2f}%"
            disk_utilization = f"{float(format_data['Disk Utilization']) * 100:.2f}%"

            data_by_report[report_name] = {
                'saw_values': saw_values,
                'cpu_utilization': cpu_utilization,
                'memory_utilization': memory_utilization,
                'disk_utilization': disk_utilization
            }
        return data_by_report

    def search_and_replace_values(self, prs, search_values, replacement_values):
        occurrence_index = 0  # Track which set of replacement values to use

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    for row in table.rows:
                        for cell in row.cells:
                            text_frame = cell.text_frame
                            for paragraph in text_frame.paragraphs:
                                for run in paragraph.runs:
                                    for i, search_value in enumerate(search_values):
                                        # Check if the search value is in the run's text
                                        if search_value in run.text and occurrence_index < len(replacement_values):
                                            # Get the corresponding replacement value for the current occurrence
                                            replacement_set = replacement_values[occurrence_index]
                                            replacement_value = replacement_set[i]

                                            # Replace the text
                                            start = run.text.find(search_value)
                                            end = start + len(search_value)
                                            run.text = run.text[:start] + replacement_value + run.text[end:]

                                            # Optional: Set color to red if above 85%
                                            try:
                                                if float(replacement_value.strip('%')) > 85:
                                                    self.set_text_color(run, RGBColor(255, 0, 0))  # Red color
                                            except ValueError:
                                                pass  # In case the replacement value is not a number

                                    # Move to the next set of replacements after each full set
                                    occurrence_index += 1
                                    if occurrence_index >= len(replacement_values):
                                        occurrence_index = 0  # Reset to avoid index error if replacements are exhausted

    def set_text_color(self, run, rgb_color):
        run.font.color.rgb = rgb_color

    def process_presentation(self, excel_file):
        try:
            # Load the Excel file
            df_report_cycle = pd.read_excel(excel_file, sheet_name='Servers Part of Report Cycle')
            df_format_box = pd.read_excel(excel_file, sheet_name='Format Box')

            data_by_report = self.extract_data(df_report_cycle, df_format_box)

            for report_name, data in data_by_report.items():
                # Load the presentation
                prs = Presentation()  # Initialize a new presentation

                # Apply the replacements
                search_values = ["31.77%", "53.07%", "83.07%"]
                replacement_values = [[data['cpu_utilization'], data['memory_utilization'], data['disk_utilization']]]

                self.search_and_replace_values(prs, search_values, replacement_values)

                # Save the presentation
                save_path = f"{report_name}_modified.pptx"
                prs.save(save_path)
                messagebox.showinfo("Success", f"Changes saved to {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def start_processing(self):
        excel_file = self.entry_file_path.get()
        if not excel_file:
            messagebox.showerror("Error", "Please select an Excel file.")
            return
        self.process_presentation(excel_file)

# Set up the main window
root = tk.Tk()
app = PowerPointReplacer(root)
root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
import pandas as pd

class PowerPointProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Report Processor")
        
        # File selection for Excel and PowerPoint
        tk.Label(root, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=5)
        self.entry_excel_path = tk.Entry(root, width=50)
        self.entry_excel_path.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.browse_excel_file).grid(row=0, column=2, padx=10, pady=5)

        tk.Label(root, text="Select PowerPoint Template:").grid(row=1, column=0, padx=10, pady=5)
        self.entry_ppt_path = tk.Entry(root, width=50)
        self.entry_ppt_path.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.browse_ppt_file).grid(row=1, column=2, padx=10, pady=5)
        
        # Progress label
        self.progress_label = tk.Label(root, text="")
        self.progress_label.grid(row=3, column=0, columnspan=3, padx=10, pady=5)

        # Process button
        tk.Button(root, text="Process Reports", command=self.process_reports).grid(row=4, column=0, columnspan=3, padx=10, pady=20)

    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.entry_excel_path.delete(0, tk.END)
        self.entry_excel_path.insert(0, filepath)

    def browse_ppt_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx")])
        self.entry_ppt_path.delete(0, tk.END)
        self.entry_ppt_path.insert(0, filepath)

    def load_data(self):
        excel_file = self.entry_excel_path.get()
        df_report_cycle = pd.read_excel(excel_file, sheet_name='Servers Part of Report Cycle')
        df_format_box = pd.read_excel(excel_file, sheet_name='Format Box')
        return df_report_cycle, df_format_box

    def extract_data(self, df_report_cycle, df_format_box):
        report_names = df_report_cycle['Report Name'].unique()
        data_by_report = {}
        for report_name in report_names:
            report_data = df_report_cycle[df_report_cycle['Report Name'] == report_name]
            saw_values = report_data['Hostname'].astype(str).tolist()  # Ensure values are strings
            format_data = df_format_box[df_format_box['Report Name'] == report_name].iloc[0]
            data_by_report[report_name] = {
                'saw_values': saw_values,
                'cpu_utilization': str(format_data['CPU Utilization']),
                'memory_utilization': str(format_data['Memory Utilization']),
                'disk_utilization': str(format_data['Disk Utilization'])
            }
        return data_by_report

    def apply_saw_replacements(self, prs, saw_values):
        replacements = {f"SAW{i+1:02}": value for i, value in enumerate(saw_values)}
        for slide in prs.slides:
            for shape in slide.shapes:
                self.process_shape(shape, replacements)

    def process_shape(self, shape, replacements):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for s in shape.shapes:
                self.process_shape(s, replacements)
        elif shape.has_text_frame:
            text_frame = shape.text_frame
            self.replace_text_in_text_frame(text_frame, replacements)

        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    text_frame = cell.text_frame
                    self.replace_text_in_text_frame(text_frame, replacements)

    def replace_text_in_text_frame(self, text_frame, replacements):
        if text_frame is not None:
            for paragraph in text_frame.paragraphs:
                full_text = ''.join([run.text for run in paragraph.runs])
                for old_text, new_text in replacements.items():
                    if old_text in full_text:
                        full_text = full_text.replace(old_text, new_text, 1)  # Replace only the first occurrence
                        for run in paragraph.runs:
                            run.text = ''
                        paragraph.runs[0].text = full_text

    def apply_combined_replacements(self, prs, values_list):
        for i, (cpu_utilization, memory_utilization, disk_utilization) in enumerate(values_list):
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_table:
                        self.search_and_replace_value(shape, '31.77%', cpu_utilization)
                        self.search_and_replace_value(shape, '53.07%', memory_utilization)
                        self.search_and_replace_value(shape, '83.07%', disk_utilization)
                        break  # Move to the next set of values

    def search_and_replace_value(self, shape, search_value, replacement_value):
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    if search_value in cell.text:
                        text_frame = cell.text_frame
                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                if search_value in run.text:
                                    start = run.text.find(search_value)
                                    end = start + len(search_value)
                                    run.text = run.text[:start] + replacement_value + run.text[end:]

                                    # Check if the replacement value is above 85% and set color to red
                                    try:
                                        if float(replacement_value.strip('%')) > 85:
                                            self.set_text_color(run, RGBColor(255, 0, 0))  # Red color
                                    except ValueError:
                                        pass  # In case the replacement value is not a number
                                    return  # Exit after the first match per shape

    def set_text_color(self, run, color):
        run.font.color.rgb = color

    def process_reports(self):
        try:
            df_report_cycle, df_format_box = self.load_data()
            data_by_report = self.extract_data(df_report_cycle, df_format_box)

            ppt_template = self.entry_ppt_path.get()
            if not ppt_template:
                messagebox.showerror("Error", "Please select a PowerPoint template file.")
                return

            for report_name, data in data_by_report.items():
                prs = Presentation(ppt_template)
                self.apply_saw_replacements(prs, data['saw_values'])
                
                # Extract combined replacement values and convert to list of tuples
                combined_values = [
                    (row['CPU Utilization'], row['Memory Utilization'], row['Disk Utilization'])
                    for _, row in df_format_box[df_format_box['Report Name'] == report_name].iterrows()
                ]
                
                self.apply_combined_replacements(prs, combined_values)
                
                save_path = f'{report_name}_updated.pptx'
                prs.save(save_path)
                print(f'Saved {save_path}')
            
            messagebox.showinfo("Success", "All reports processed and saved.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerPointProcessorApp(root)
    root.mainloop()
    

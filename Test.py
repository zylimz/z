import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from pptx.dml.color import RGBColor
import pandas as pd

class PowerPointProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Report Processor")

        # Set up the notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, padx=10, pady=10)

        # First tab for selecting files and processing
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Process Reports")

        tk.Label(self.tab1, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=5)
        self.entry_excel_path = tk.Entry(self.tab1, width=50)
        self.entry_excel_path.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(self.tab1, text="Browse", command=self.browse_excel_file).grid(row=0, column=2, padx=10, pady=5)

        tk.Label(self.tab1, text="Select PowerPoint Template File:").grid(row=1, column=0, padx=10, pady=5)
        self.entry_ppt_path = tk.Entry(self.tab1, width=50)
        self.entry_ppt_path.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(self.tab1, text="Browse", command=self.browse_ppt_file).grid(row=1, column=2, padx=10, pady=5)

        tk.Button(self.tab1, text="Process Reports", command=self.process_reports).grid(row=2, column=1, padx=10, pady=20)

    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        self.entry_excel_path.delete(0, tk.END)
        self.entry_excel_path.insert(0, filepath)

    def browse_ppt_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("PowerPoint Files", "*.pptx")]
        )
        self.entry_ppt_path.delete(0, tk.END)
        self.entry_ppt_path.insert(0, filepath)

    def load_data(self):
        excel_path = self.entry_excel_path.get()
        if not excel_path:
            messagebox.showerror("Error", "Please select an Excel file.")
            return None, None
        
        df_report_cycle = pd.read_excel(excel_path, sheet_name='Servers Part of Report Cycle')
        df_format_box = pd.read_excel(excel_path, sheet_name='Format Box')
        
        return df_report_cycle, df_format_box

    def extract_data(self, df_report_cycle, df_format_box):
        data_by_report = {}

        # Extract percentage values from the "Format Box" sheet
        combined_replacements = []
        for _, row in df_format_box.iterrows():
            cpu_values = row['CPU Utilization']
            mem_values = row['Memory Utilization']
            disk_values = row['Disk Utilization']

            # Convert values to percentages and format them
            formatted_cpu = f"{float(cpu_values) * 100:.2f}%"
            formatted_mem = f"{float(mem_values) * 100:.2f}%"
            formatted_disk = f"{float(disk_values) * 100:.2f}%"
            
            combined_replacements.append([formatted_cpu, formatted_mem, formatted_disk])
        
        # Extract values from the "Servers Part of Report Cycle" sheet
        for _, row in df_report_cycle.iterrows():
            report_name = row['Report Name']
            hostname = row['Hostname']
            
            if report_name not in data_by_report:
                data_by_report[report_name] = {'saw_values': [], 'combined_replacements': []}
            
            data_by_report[report_name]['saw_values'].append(hostname)
            data_by_report[report_name]['combined_replacements'] = combined_replacements
        
        return data_by_report

    def search_and_replace_value(self, prs, search_value, replacement_value):
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
                                            start = run.text.find(search_value)
                                            end = start + len(search_value)
                                            run.text = run.text[:start] + replacement_value + run.text[end:]

                                            # Check if the replacement value is above 85% and set color to red
                                            try:
                                                if float(replacement_value.strip('%')) > 85:
                                                    self.set_text_color(run, RGBColor(255, 0, 0))  # Red color
                                            except ValueError:
                                                pass  # In case the replacement value is not a number
                                            return  # Exit after the first match per slide

    def set_text_color(self, run, rgb_color):
        run.font.color.rgb = rgb_color

    def apply_saw_replacements(self, prs, saw_values):
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    for row in table.rows:
                        for cell in row.cells:
                            for saw_value in saw_values:
                                if saw_value in cell.text:
                                    self.search_and_replace_value(prs, saw_value, saw_value)

    def apply_combined_replacements(self, prs, combined_replacements):
        for replacement_set in combined_replacements:
            for slide in prs.slides:
                tables = [shape.table for shape in slide.shapes if shape.has_table]
                for table in tables:
                    for placeholder, replacement_value in zip(['31.77%', '53.07%', '83.07%'], replacement_set):
                        self.search_and_replace_value(prs, placeholder, replacement_value)
                        break  # Exit after processing one placeholder per slide

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
                combined_values = data['combined_replacements']
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

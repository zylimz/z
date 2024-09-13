import pandas as pd
from pptx import Presentation
from tkinter import Tk, filedialog, StringVar, Text, Button, Label, ttk
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

def load_excel_data(file_path):
    # Load Excel data
    excel_data = pd.ExcelFile(file_path)
    
    # Extract data from relevant sheets
    draft_data = pd.read_excel(excel_data, sheet_name='Sheet1')  # Modify with actual sheet name if different
    combined_data = pd.read_excel(excel_data, sheet_name='Sheet2')  # Modify with actual sheet name if different
    report_cycle_data = pd.read_excel(excel_data, sheet_name='Servers Part of Report Cycle')
    
    # Extract unique report names
    unique_names = report_cycle_data['Report Name'].unique()
    
    return draft_data, combined_data, unique_names

def replace_draft_application(prs, replacement_text):
    # Replace "Draft Template" with replacement_text on the first slide
    slide = prs.slides[0]  # First slide
    for shape in slide.shapes:
        if shape.has_text_frame and "Draft Template" in shape.text:
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.text = run.text.replace("Draft Template", replacement_text)
                    run.font.color.rgb = (0, 0, 0)  # Set color to black

def replace_saw_values(prs, saw_replacements):
    # Replace SAW values throughout the presentation
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                replace_text_in_text_frame(shape.text_frame, saw_replacements)

def replace_combined_values(prs, combined_replacements):
    # Replace values based on the combined replacements
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        if "31.77%" in cell.text:
                            cell.text = cell.text.replace("31.77%", combined_replacements.get("31.77%", ""))
                            # Retain original formatting
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    run.font.size = Pt(12)  # Example size, adjust as needed
                                    run.font.color.rgb = (0, 0, 0)  # Set color to black
                        elif "53.07%" in cell.text:
                            cell.text = cell.text.replace("53.07%", combined_replacements.get("53.07%", ""))
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    run.font.size = Pt(12)
                                    run.font.color.rgb = (0, 0, 0)
                        elif "83.07%" in cell.text:
                            cell.text = cell.text.replace("83.07%", combined_replacements.get("83.07%", ""))
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    run.font.size = Pt(12)
                                    run.font.color.rgb = (0, 0, 0)

def run_operations_for_each_name(excel_path, pptx_template):
    draft_data, combined_data, unique_names = load_excel_data(excel_path)
    
    for name in unique_names:
        # Filter data based on the current name
        draft_row = draft_data[draft_data['name'] == name]
        combined_rows = combined_data[combined_data['name'] == name]

        # Prepare replacement values
        draft_replacement = draft_row['hostname'].iloc[0] if not draft_row.empty else ""
        saw_replacements = {f"SAW{str(i).zfill(2)}": row['hostname'] for i, row in enumerate(draft_row.itertuples(), 1)}
        combined_replacements = {
            "31.77%": combined_rows['memory utilisation'].iloc[0] if not combined_rows.empty else "",
            "53.07%": combined_rows['CPU utilisation'].iloc[0] if not combined_rows.empty else "",
            "83.07%": combined_rows['disk utilisation'].iloc[0] if not combined_rows.empty else ""
        }

        # Load PowerPoint template
        prs = Presentation(pptx_template)

        # Run replacements
        replace_draft_application(prs, draft_replacement)
        replace_saw_values(prs, saw_replacements)
        replace_combined_values(prs, combined_replacements)

        # Save the modified PowerPoint file for each unique name
        prs.save(f"{name}_Modified.pptx")

def select_files_and_run():
    # GUI for selecting Excel and PowerPoint files
    root = Tk()
    root.withdraw()  # Hide the root window
    excel_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    pptx_template = filedialog.askopenfilename(title="Select PowerPoint Template", filetypes=[("PowerPoint files", "*.pptx")])
    root.destroy()
    
    if excel_path and pptx_template:
        run_operations_for_each_name(excel_path, pptx_template)
        print("Operations completed for all unique names.")

# GUI Button to start the file selection and process
if __name__ == "__main__":
    select_files_and_run()

import os
from pathlib import Path
from datetime import datetime
import pandas as pd
import argparse

# Import the logic from our other scripts
from mapper import process_timesheet
from pdf2excel import convert as pdf_to_excel

def main():
    parser = argparse.ArgumentParser(description="Process a folder of PDF timesheets and compile them into templates.")
    parser.add_argument("input_folder", help="Path to the folder containing PDF or Excel timesheets.")
    args = parser.parse_args()

    input_dir = Path(args.input_folder)
    
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Error: Directory '{input_dir}' does not exist.")
        return

    # Create the Outputs directory inside the target folder
    output_dir = input_dir / "Outputs"
    output_dir.mkdir(exist_ok=True)

    # 1. Load the empty templates
    # Templates live in the repo's own templates/ folder
    base_path = Path(__file__).resolve().parent / "templates"
    template_admin_path = base_path / "Admin.xlsx"
    template_absence_path = base_path / "Absense.xlsx"
    template_overtime_path = base_path / "overtime.xlsx"

    print("Loading templates...")
    df_admin = pd.read_excel(template_admin_path)
    df_absence = pd.read_excel(template_absence_path)
    df_overtime = pd.read_excel(template_overtime_path)

    # 2. Loop through the directory
    # We will process both PDFs (by converting them) and direct .xlsx files if they exist
    pdf_files = list(input_dir.glob("*.pdf"))
    xlsx_files = list(input_dir.glob("*.xlsx"))

    files_to_process = pdf_files + [f for f in xlsx_files if "Final_Timesheet" not in f.name and "Admin" not in f.name and "Absense" not in f.name and "overtime" not in f.name]

    if not files_to_process:
        print(f"No PDF or Excel files found in {input_dir}.")
        return

    print(f"Found {len(files_to_process)} file(s) to process.")

    for file_path in files_to_process:
        try:
            print(f"\nProcessing: {file_path.name}")
            
            # If it's a PDF, convert it to Excel first
            if file_path.suffix.lower() == ".pdf":
                print("  -> Converting PDF to Excel...")
                # The converted Excel is saved alongside the PDF
                excel_path = pdf_to_excel(str(file_path))
                print(f"  -> Converted successfully: {excel_path}")
            else:
                # If it's already an Excel file, use it directly
                excel_path = str(file_path)

            print("  -> Extracting and mapping data...")
            # Run the mapper script which updates the DataFrames
            df_admin, df_absence, df_overtime = process_timesheet(
                excel_path, df_admin, df_absence, df_overtime
            )
            print("  -> Mapping successful.")
            
        except Exception as e:
            print(f"  -> Error processing {file_path.name}: {e}")

    # 3. Save all updated DataFrames to a single final Excel file
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    final_output_path = output_dir / f"Final_Timesheet_Report_{current_time}.xlsx"

    print(f"\nSaving final compiled report to: {final_output_path}")
    with pd.ExcelWriter(final_output_path, engine='openpyxl') as writer:
        df_admin.to_excel(writer, sheet_name="Admin", index=False)
        df_absence.to_excel(writer, sheet_name="Absence", index=False)
        df_overtime.to_excel(writer, sheet_name="Overtime", index=False)

    print("✅ Pipeline complete! All files processed and appended.")

if __name__ == "__main__":
    main()

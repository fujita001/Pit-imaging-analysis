import os
import pandas as pd
from tkinter import filedialog
import tkinter as tk

def select_folder():
    """
    Displays a dialog to select a folder and returns the path to the selected folder.

    Returns:
    - folder_path (str): Path of the selected folder
    """
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Please select the folder where the CSV file is stored")
    return folder_path

def extract_sheet_name(file_name):
    """
    Extract the sheet name from the file name.

    Args:
    - file_name (str): File name (e.g., “prefix_part1_part2_ROI123.csv”)

    Returns:
    - sheet_name (str): Sheet name (e.g., “part1_part2ROI123”)
    """
    try:
        # Split file names with underscores
        parts = file_name.split("_")
        if len(parts) < 4:
            raise ValueError(f"File name ‘{file_name}’ is not in the expected format.")

        # Extract the portion between the first and third underscores
        middle_part = "_".join(parts[1:3])
        
        # Extract strings after "ROI
        roi_part = file_name.split("ROI")[-1].split(".")[0]

        # Assemble sheet name
        sheet_name = f"{middle_part}ROI{roi_part}"
        
        # Adjust sheet name length (31 character limit in Excel)
        return sheet_name[:31]
    except Exception as e:
        print(f"Sheet name generation error: {str(e)}")
        return "InvalidSheetName"

def merge_all_csv_to_excel(folder_path):
    """
    Combine all CSV files in the specified folder into one Excel file.

    Args:
    - folder_path (str): Path of the root folder where the CSV file is stored

    Returns:
    - None
    """
    # Destination Excel file
    save_path = os.path.join(folder_path, "CombinedWorkbook.xlsx")
    
    # Create an ExcelWriter object
    writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    
    for root, _, files in os.walk(folder_path):
        csv_files = [f for f in files if f.endswith('.csv')]
        
        for csv_file in csv_files:
            csv_path = os.path.join(root, csv_file)
            
            try:
                # Extract sheet name
                sheet_name = extract_sheet_name(csv_file)
                
                # Import CSV files as DataFrames
                df = pd.read_csv(csv_path, encoding='utf-8')
                
                # Write DataFrame to a new sheet in Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            except Exception as e:
                print(f"Error while processing CSV file ‘{csv_file}’: {str(e)}")
    
    # Save Excel file
    writer._save()
    writer.close()
    print(f"The CSV file has been compiled into an Excel file ({save_path})!")

if __name__ == "__main__":
    # Select Folder
    root_folder = select_folder()
    if not root_folder:
        print("No folder was selected. Abort the process.")
    else:
        merge_all_csv_to_excel(root_folder)

import pandas as pd
from openpyxl import load_workbook

def load_and_analyze_xlsx(file_path, sheets_to_analyze):
    """
    Load an .xlsx file with multiple sheets and analyze specific sheets.

    Parameters:
        file_path (str): Path to the .xlsx file.
        sheets_to_analyze (list): List of sheet names to analyze.

    Returns:
        None: Prints the comment at cell C5 for each sheet.
    """
    try:
        # Load the workbook
        workbook = load_workbook(filename=file_path, data_only=False)

        # Iterate through the specified sheets
        for sheet in sheets_to_analyze:
            if sheet in workbook.sheetnames:
                worksheet = workbook[sheet]

                # Get the comment at cell C5
                cell_c5 = worksheet['C5']
                comment_at_c5 = cell_c5.comment.text if cell_c5.comment else "No comment"

                print(f"Comment at C5 in sheet '{sheet}': {comment_at_c5}")
            else:
                print(f"Sheet '{sheet}' not found in the file.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
if __name__ == "__main__":
    file_path = "Worcester Final week Schedule 2024.xlsx"  # Replace with your .xlsx file name
    sheets_to_analyze = ["Line", "Dish"]  # Replace with sheet names to analyze
    load_and_analyze_xlsx(file_path, sheets_to_analyze)

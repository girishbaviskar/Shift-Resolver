from openpyxl import load_workbook
from employee import Employee
def parse_comments(raw_comment):
    """
    Parses a raw comment string into a list of (comment, commenter) tuples.

    Parameters:
        raw_comment (str): The raw comment text from a cell.

    Returns:
        list: A list of tuples, each containing (comment, commenter).
    """
    if not raw_comment:
        return []

    # Split the comments into a list by the delimiter '\n----\n'
    comment_list = raw_comment.split('\n----\n')
    processed_comments = []

    for item in comment_list:
        # Further split into comment and commenter using '\n\t-' as the delimiter
        parts = item.split('\n\t-')
        if len(parts) == 2:
            comment, commenter = parts
            processed_comments.append((comment.strip(), commenter.strip()))
        else:
            # If no valid structure, add as "Unknown"
            processed_comments.append((parts[0].strip(), "Unknown"))

    return processed_comments


def load_and_assign_shift_xlsx(file_path, sheets_to_analyze):
    """
    Load an .xlsx file, analyze specific sheets, and assign shifts based on the last commenter.
    Updates the cell value with the last commenter's name.

    Parameters:
        file_path (str): Path to the .xlsx file.
        sheets_to_analyze (list): List of sheet names to analyze.

    Returns:
        dict: A dictionary with sheet names as keys and the last person who commented assigned to each shift.
    """
    # Initialize result dictionary
    results = {}

    try:
        # Load the workbook
        workbook = load_workbook(filename=file_path, data_only=False)

        # Iterate through the specified sheets
        for sheet in sheets_to_analyze:
            if sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
                for row in worksheet.iter_rows():
                    for cell in row:
                        # Skip cells that already have a value or are part of a merged range
                        if not cell.value and cell.comment:
                            # Get the raw comment
                            raw_comment = cell.comment.text if cell.comment else None

                            # Parse the comment into tuples
                            processed_comments = parse_comments(raw_comment)

                            # Assign the last commenter to the shift
                            last_commenter_tuple = processed_comments[-1] if processed_comments else ("No comment", "Unknown")
                            last_commenter_comment = last_commenter_tuple[0]
                            last_commenter_name = last_commenter_tuple[1]
                            #check if atleast one dish shift is taken
                            
                        
                            # Update the cell with the last commenter's name
                            cell.value = last_commenter_comment
                            cell.comment = None

                            # Store the result for the current sheet
                            results[sheet] = last_commenter_comment
            else:
                results[sheet] = "Sheet not found"
        
        # Save changes to the workbook
        workbook.save(filename=file_path)

    except Exception as e:
        print(f"An error occurred: {e}")
        for sheet in sheets_to_analyze:
            results[sheet] = f"Error: {str(e)}"

    return results

# Example function usage
file_path = "Worcester Final week Schedule 2024.xlsx"  # Replace with your .xlsx file name
sheets_to_analyze = ["Dish"]  # Replace with sheet names to analyze
shift_assignments = load_and_assign_shift_xlsx(file_path, sheets_to_analyze)
print(str(shift_assignments))
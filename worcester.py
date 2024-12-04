import pandas as pd

# Load and preprocess the data
def load_data(file_path):
    data = pd.read_csv(file_path)

    # Extract Dates and Shifts
    data['Date'] = data['Grab & Go'].where(
        data['Grab & Go'].str.contains(r'\w+day \d{1,2}/\d{1,2}', na=False)
    ).ffill()

    data['Shift'] = data['Grab & Go'].where(
        data['Grab & Go'].str.contains(r'\d{1,2}:\d{2}[AP]M - \d{1,2}:\d{2}[AP]M', na=False)
    ).ffill()

    # Clean up irrelevant rows
    data_cleaned = data.dropna(subset=['Unnamed: 2', 'Unnamed: 3'], how='all')

    # Rename columns for clarity
    data_cleaned = data_cleaned[['Date', 'Shift', 'Unnamed: 2', 'Unnamed: 3']]
    data_cleaned.columns = ['Date', 'Shift', 'Comment_1', 'Comment_2']
    return data_cleaned

# Identify and resolve conflicting shifts
def resolve_conflicts(data):
    results = []
    assignments = {}

    # Function to determine if a person has a dish room shift
    def has_dish_room_shift(person):
        return person in assignments and any('Dish Room' in shift for shift in assignments[person])

    for _, row in data.iterrows():
        comments = [row['Comment_1'], row['Comment_2']]
        comments = [c for c in comments if pd.notna(c)]  # Filter out empty comments

        if len(comments) > 1:
            # Conflict detected
            results.append({
                'Date': row['Date'],
                'Shift': row['Shift'],
                'Conflict': comments,
                'Resolution_Suggestions': {
                    'First_Comment_Wins': comments[0],
                    'Manual_Resolution': comments
                }
            })
        elif len(comments) == 1:
            person = comments[0]
            if has_dish_room_shift(person):
                # Resolve the shift assignment
                if person not in assignments:
                    assignments[person] = []
                assignments[person].append(row['Shift'])
            else:
                # Unresolved due to lack of dish room shift
                results.append({
                    'Date': row['Date'],
                    'Shift': row['Shift'],
                    'Conflict': None,
                    'Resolution_Suggestions': {
                        'Reason': f"{person} does not have a Dish Room shift."
                    }
                })

    return results, assignments

# Main function to load, process, and resolve
if __name__ == "__main__":
    file_path = 'Worcester Final week Schedule 2024 - Grab & Go.csv'

    data = load_data(file_path)
    conflicts, resolved_assignments = resolve_conflicts(data)

    # Output conflicts and resolutions
    conflicts_df = pd.DataFrame(conflicts)
    assignments_df = pd.DataFrame({
        'Person': list(resolved_assignments.keys()),
        'Shifts': [', '.join(shifts) for shifts in resolved_assignments.values()]
    })

    # Save results to CSV files for review
    conflicts_df.to_csv('shift_conflicts.csv', index=False)
    assignments_df.to_csv('shift_assignments.csv', index=False)


    print("Conflict report saved as 'shift_conflicts.csv'")
    print("Resolved assignments saved as 'shift_assignments.csv'")

from datetime import datetime

class Employee:
    def __init__(self, name):
        """
        Initialize a person object to store and manage shift details.

        Parameters:
            name (str): Full name of the person.
        """
        self.name = name
        self.shifts = []  # List of shifts: [{"location": ..., "date": ..., "time": ...}]
        self.dish_room_shift_taken = False  # Boolean flag for Dish Room shift
        self.total_shift_count = 0  # Total number of shifts
        self.total_hours = 0.0  # Total hours assigned

    def add_shift(self, location, date, time, hours):
        """
        Add a shift to the person's schedule.

        Parameters:
            location (str): Location of the shift (e.g., Dish Room).
            date (str): Date of the shift.
            time (str): Time range of the shift (e.g., "8:30AM - 12:00PM").
            hours (float): Duration of the shift in hours.
        """
        time = time.replace(" ", "")
        self.shifts.append({"location": location, "date": date, "time": time})
        self.total_shift_count += 1
        self.total_hours += hours
        if location == "Dish":
            self.dish_room_shift_taken = True

    def has_conflict(self, date_or_day, time_cell):
        """
        Check if the person has a conflict with a new shift.
        Overlaps are not allowed, but shifts can be back-to-back.

        Parameters: 
            date_or_day (str): Date of the new shift.
            time (str): Time range of the new shift (e.g., "8:30AM - 12:00PM").

        Returns:
            bool: True if there is a conflict, False otherwise.
        """
        # Parse the new shift's start and end times
        time = time_cell.value
        time = time.replace(" ", "")
        try:
            new_start, new_end = map(
                lambda t: datetime.strptime(t.strip(), "%I:%M%p"),
                time.split('-')
            )
        except ValueError:
            raise ValueError(f"{time_cell}Time should be in the format '8:30AM - 12:00PM'")
        
        for shift in self.shifts:
            if shift["date"] == date_or_day:
                # Parse the existing shift's start and end times
                existing_start, existing_end = map(
                    lambda t: datetime.strptime(t.strip(), "%I:%M%p"),
                    shift["time"].split('-')
                )
                # Check for overlap
                if not (new_end <= existing_start or new_start >= existing_end):
                    return True
        
        return False


    def get_summary(self):
        """
        Get a summary of the person's assigned shifts and statistics.

        Returns:
            dict: Summary of the person's shifts, counts, and hours.
        """
        return {
            "name": self.name,
            "shifts": self.shifts,
            "dish_room_shift_taken": self.dish_room_shift_taken,
            "total_shift_count": self.total_shift_count,
            "total_hours": self.total_hours,
        }

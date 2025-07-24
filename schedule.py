import pandas as pd
from datetime import datetime, timedelta, time
import re

def parse_time_string(time_str):
    """
    Converts a time string in 'HH:MM' format to a Python time object.
    Returns None if the format is incorrect.
    """
    try:
        if not isinstance(time_str, str):
            time_str = str(time_str)
        return datetime.strptime(time_str, '%H:%M').time()
    except (ValueError, TypeError):
        return None

def calculate_duration(in_time_str, out_time_str):
    """
    Calculates the time difference between an 'in' and 'out' punch.
    Handles overnight shifts correctly.
    """
    in_time = parse_time_string(in_time_str)
    out_time = parse_time_string(out_time_str)

    if in_time and out_time:
        dummy_date = datetime(2000, 1, 1)
        dt_in = datetime.combine(dummy_date, in_time)
        dt_out = datetime.combine(dummy_date, out_time)

        if dt_out < dt_in:
            dt_out += timedelta(days=1)
        
        return dt_out - dt_in
    return timedelta(0)

def format_timedelta(td):
    """
    Formats a timedelta object into a readable HH:MM:SS string.
    """
    td_in_seconds = int(td.total_seconds())
    hours, remainder = divmod(td_in_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def process_punch_records(punch_records_str):
    """
    Processes the 'Punch Records' string to calculate net work hours,
    extracts shift timings and their durations, and accounts for a 30-minute break.
    Returns a tuple: (net_work_duration, list_of_shifts_with_duration)
    """
    if not isinstance(punch_records_str, str) or not punch_records_str.strip():
        return timedelta(0), []

    # Find all 'HH:MM:in' or 'HH:MM:out' patterns
    pattern = re.compile(r'(\d{2}:\d{2}):(in|out)')
    punches = pattern.findall(punch_records_str)
    
    if not punches:
        return timedelta(0), []

    parsed_punches = []
    for time_str, punch_type in punches:
        parsed_punches.append({'time_str': time_str, 'type': punch_type})
    
    parsed_punches.sort(key=lambda x: parse_time_string(x['time_str']))

    total_work_duration = timedelta(0)
    shifts = []
    in_time = None
    for punch in parsed_punches:
        if punch['type'] == 'in':
            in_time = punch['time_str']
        elif punch['type'] == 'out' and in_time:
            duration = calculate_duration(in_time, punch['time_str'])
            total_work_duration += duration
            # Store shift timing and its formatted duration
            shift_info = {
                "timing": f"{in_time} - {punch['time_str']}",
                "duration": format_timedelta(duration)
            }
            shifts.append(shift_info)
            in_time = None

    # If there's only one shift (one in/out pair), deduct the 30-minute break.
    if len(shifts) == 1:
        net_work_duration = total_work_duration - timedelta(minutes=30)
        return max(net_work_duration, timedelta(0)), shifts
    
    # If there are multiple shifts, the break is the time between them, so no deduction is needed.
    return total_work_duration, shifts


def get_first_in_time(punch_records_str):
    """Helper function to get the earliest 'in' punch time from the record string."""
    if not isinstance(punch_records_str, str) or not punch_records_str.strip():
        return None

    pattern = re.compile(r'(\d{2}:\d{2}):(in|out)')
    punches = pattern.findall(punch_records_str)
    
    in_punch_times = [parse_time_string(p[0]) for p in punches if p[1] == 'in' and parse_time_string(p[0]) is not None]
    
    if not in_punch_times:
        return None

    return min(in_punch_times)

def determine_status_and_ot(row):
    """
    Determines attendance status and calculates overtime based on net work duration and arrival time.
    Returns a tuple: (status, overtime_duration)
    """
    net_work_duration = row['Calculated Hours']
    punch_records = row['Punch Records']

    # --- ATTENDANCE & OT RULES ---
    LATE_ARRIVAL_START = time(10, 30)
    LATE_ARRIVAL_END = time(13, 30)
    HALF_DAY_MIN_DURATION = timedelta(hours=4)
    FULL_DAY_MIN_DURATION = timedelta(hours=8) # 8 hours of net work
    STANDARD_WORKDAY = timedelta(hours=8)
    # --------------------------------

    first_in_time = get_first_in_time(punch_records)
    overtime = timedelta(0)
    status = "Absent" # Default status

    # Rule 1: No OT for late arrivals
    if first_in_time and first_in_time > LATE_ARRIVAL_START:
        overtime = timedelta(0)
        # Determine status for late arrivals
        if first_in_time > LATE_ARRIVAL_END:
            status = "Absent"
        else:
            status = "Half Day"
    else: # On-time arrivals
        # Calculate OT for on-time arrivals
        if net_work_duration > STANDARD_WORKDAY:
            overtime = net_work_duration - STANDARD_WORKDAY

        # Determine status for on-time arrivals
        if net_work_duration < HALF_DAY_MIN_DURATION:
            status = "Absent"
        elif net_work_duration >= FULL_DAY_MIN_DURATION:
            status = "Full Day"
        else:
            status = "Half Day"
            
    return status, overtime


def analyze_attendance_report(file_path):
    """
    Main function to read the Excel file, clean the data,
    and calculate the work hours and status.
    """
    print(f"\nProcessing file: {file_path}")
    try:
        df = pd.read_excel(file_path, header=9, sheet_name=0)

        print("Columns found in Excel file:", df.columns.tolist())
        
        df.columns = df.columns.str.strip()
        df.dropna(subset=['Name'], inplace=True)

        if 'Punch Records' not in df.columns:
            print("Error: 'Punch Records' column not found.")
            return pd.DataFrame()

        # Apply the function to get both duration and shifts
        punch_data = df['Punch Records'].apply(process_punch_records)
        
        # Separate the results into new columns
        df['Calculated Hours'] = punch_data.apply(lambda x: x[0])
        df['Shifts'] = punch_data.apply(lambda x: x[1])
        df['Shift 1'] = df['Shifts'].apply(lambda x: x[0]['timing'] if len(x) > 0 else '')
        df['Shift 1 Duration'] = df['Shifts'].apply(lambda x: x[0]['duration'] if len(x) > 0 else '')
        df['Shift 2'] = df['Shifts'].apply(lambda x: x[1]['timing'] if len(x) > 1 else '')
        df['Shift 2 Duration'] = df['Shifts'].apply(lambda x: x[1]['duration'] if len(x) > 1 else '')

        # Apply function to get status and OT
        status_ot_data = df.apply(determine_status_and_ot, axis=1)
        df['Calculated Status'] = status_ot_data.apply(lambda x: x[0])
        df['Calculated OT'] = status_ot_data.apply(lambda x: x[1])

        df['Calculated Hours (HH:MM:SS)'] = df['Calculated Hours'].apply(format_timedelta)
        df['Calculated OT (HH:MM:SS)'] = df['Calculated OT'].apply(format_timedelta)


        # Define the final columns for the output file
        final_cols = ['Name', 'Calculated Status', 'Calculated Hours (HH:MM:SS)', 'Calculated OT (HH:MM:SS)', 'Shift 1', 'Shift 1 Duration', 'Shift 2', 'Shift 2 Duration']
        
        # Create a temporary full dataframe for console printing
        console_cols = ['Name', 'Status', 'Calculated Status', 'Work Dur.', 'OT', 'Calculated OT (HH:MM:SS)', 'Tot. Dur.', 'Calculated Hours (HH:MM:SS)', 'Shift 1', 'Shift 1 Duration', 'Shift 2', 'Shift 2 Duration']
        console_display_df = df[[col for col in console_cols if col in df.columns]]
        
        # Return both the final dataframe for the file and the console dataframe
        return df[final_cols], console_display_df

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        return None, None
    except KeyError as e:
        print(f"Error: A required column was not found. Please check your Excel file for column: {e}")
        return None, None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None, None

# --- Main Execution Block ---
if __name__ == "__main__":
    input_file = "Daily Attendance Report.xlsx"
    output_file = "Calculated Attendance Report.xlsx"

    # The function now returns two dataframes
    final_df, console_df = analyze_attendance_report(input_file)
    
    if final_df is not None and console_df is not None:
        print(f"\n--- Full Results for Console ---")
        print(console_df.to_string(index=False))

        # --- CONFUSION MATRIX ---
        if 'Status' in console_df.columns and 'Calculated Status' in console_df.columns:
            print("\n\n--- Confusion Matrix ---")
            print("Compares original report status vs. new calculated status.")
            
            confusion_matrix = pd.crosstab(
                console_df['Status'].str.strip(),
                console_df['Calculated Status'],
                rownames=['Original Status'],
                colnames=['Calculated Status']
            )
            print(confusion_matrix)
        # -----------------------------

        try:
            # Save the dataframe with only the new columns to the Excel file
            final_df.to_excel(output_file, index=False)
            print(f"\nSuccessfully created/updated the report at: {output_file}")
            print("The saved file now contains the 'Name', 'Calculated Status', 'Calculated Hours', 'Calculated OT', 'Shift 1', 'Shift 1 Duration', 'Shift 2', and 'Shift 2 Duration' columns.")
        except Exception as e:
            print(f"\nError saving file: {e}")

    print("\n--- Script Finished ---")

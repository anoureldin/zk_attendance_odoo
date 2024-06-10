import pandas as pd
from datetime import datetime, timedelta
from dateutil.parser import parse
import os


def preprocess_data(file_path):
    # Read the file based on its extension
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    # Convert all column names to lowercase for consistency
    df.columns = df.columns.str.lower()

    # Check if the required columns are present
    if 'no' not in df.columns or 'time' not in df.columns:
        raise ValueError("Input file must contain 'No' for employee number and 'Time' for attendance date/time.")

    # Rename the identified columns to standard names
    df = df.rename(columns={'no': 'Employee', 'time': 'Check in'})

    # Convert Arabic AM/PM and other incorrect characters to English
    df['Check in'] = df['Check in'].str.replace('ص', 'AM').str.replace('م', 'PM').str.replace('Õ', 'AM').str.replace(
        'ã', 'PM')

    # Parse the Date/Time column to handle various formats
    df['Check in'] = df['Check in'].apply(lambda x: parse(x, dayfirst=True, fuzzy=True) if pd.notnull(x) else x)

    # Sort by employee number and Date/Time
    df = df.sort_values(by=['Employee', 'Check in'])

    return df


def process_attendance(df):
    # List to store processed records
    processed_records = []

    # Group by employee number and date
    df['Date'] = df['Check in'].dt.date
    grouped = df.groupby(['Employee', 'Date'])

    for (employee, date), group in grouped:
        group = group.reset_index(drop=True)

        print(f"Processing employee {employee} on {date}")
        print(group)

        # Initialize shift counters
        shift_start = None
        shift_end = None

        for i in range(len(group)):
            current_time = group.loc[i, 'Check in']

            if shift_start is None:
                # Start a new shift
                shift_start = current_time
                shift_end = None
                print(f"Shift start: {shift_start}")
                continue

            if current_time - shift_start < timedelta(minutes=30):
                # Ignore attendance within 30 minutes of check-in
                print(f"Ignoring attendance within 30 minutes: {current_time}")
                continue

            if shift_end is None or current_time - shift_end >= timedelta(hours=1):
                if shift_end is not None:
                    # More than an hour from the last attendance, end the current shift and start a new one
                    processed_records.append((employee, shift_start, shift_end))
                    print(f"Recording shift: {shift_start} - {shift_end}")
                    shift_start = current_time
                    shift_end = None
                else:
                    # Update the shift end time
                    shift_end = current_time

        if shift_end is not None:
            processed_records.append((employee, shift_start, shift_end))
            print(f"Final recording shift: {shift_start} - {shift_end}")
        else:
            # Only one attendance record for the day, repeat it in the empty cell
            processed_records.append((employee, shift_start, shift_start))
            print(f"Single attendance repeated in empty cell: {shift_start} - {shift_start}")

    # Convert to DataFrame with the desired column names
    processed_df = pd.DataFrame(processed_records, columns=['Employee', 'Check in', 'Check out'])
    return processed_df


def main():
    file_path = 'sdasd.csv'  # Update this path to your local file
    df = preprocess_data(file_path)
    print("Preprocessed Data:")
    print(df.head())
    print(df.info())  # Display detailed info to verify data types and values

    processed_df = process_attendance(df)

    # Generate the output file name with date and time
    base_name = os.path.splitext(file_path)[0]
    current_date_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file_path = f"{base_name}_{current_date_time}.xlsx" if file_path.endswith('.xlsx') or file_path.endswith(
        '.xlsm') else f"{base_name}_{current_date_time}.csv"

    if file_path.endswith('.csv'):
        processed_df.to_csv(output_file_path, index=False)
    else:
        processed_df.to_excel(output_file_path, index=False)

    print(f"Processed attendance saved to '{output_file_path}'.")
    print(processed_df.head())  # Display the first few rows to verify


if __name__ == "__main__":
    main()

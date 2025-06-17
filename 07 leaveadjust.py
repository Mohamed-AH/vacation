import pandas as pd
from datetime import datetime, timedelta
import os # Import os module to get current working directory

# Define the reference period for "no leave" and for calculating overlap
REFERENCE_START_DATE = datetime(2025, 6, 14)
REFERENCE_END_DATE = datetime(2025, 6, 20)

def update_excel_with_leave_data(excel_file_path, csv_file_path, output_file_path):
    """
    Updates 'Start Date' and 'End Date' columns in an Excel sheet based on
    leave data from a CSV file.

    Args:
        excel_file_path (str): Path to the input Excel file (e.g., 'AL HARAM PROJECTS.xlsx').
        csv_file_path (str): Path to the input CSV file (e.g., 'leave_analysis_report.csv').
        output_file_path (str): Path where the updated Excel file will be saved.
    """
    # Print current working directory for debugging
    print(f"Current working directory: {os.getcwd()}")

    try:
        # Load the Excel file (.xlsx)
        # We need to specify header=1 because the first row is 'Al Shamiyah Project'
        df_excel = pd.read_excel(excel_file_path, header=1) # Read the second row (index 1) as header
        print(f"Successfully loaded Excel file: {excel_file_path}")
        # Clean column names by stripping whitespace
        df_excel.columns = df_excel.columns.str.strip()
        # --- DEBUGGING AID: Print Excel columns after stripping ---
        print(f"Excel columns after stripping: {df_excel.columns.tolist()}")
        # ---------------------------------------------------------
    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file_path}' not found. Please ensure it's in the correct directory.")
        return
    except Exception as e:
        print(f"Error loading Excel file '{excel_file_path}': {e}")
        return

    try:
        # Load the CSV file containing leave analysis report
        df_csv = pd.read_csv(csv_file_path)
        print(f"Successfully loaded CSV file: {csv_file_path}")
        # Clean column names by stripping whitespace
        df_csv.columns = df_csv.columns.str.strip()
        # --- DEBUGGING AID: Print CSV columns after stripping ---
        print(f"CSV columns after stripping: {df_csv.columns.tolist()}")
        # -------------------------------------------------------
    except FileNotFoundError:
        print(f"Error: CSV file '{csv_file_path}' not found. Please ensure it's in the correct directory.")
        return
    except Exception as e:
        print(f"Error loading CSV file '{csv_file_path}': {e}")
        return

    # Ensure 'Start Date' and 'End Date' columns exist in df_excel
    if 'Start Date' not in df_excel.columns:
        df_excel['Start Date'] = ''
    if 'End Date' not in df_excel.columns:
        df_excel['End Date'] = ''
    
    # Add a new column 'Coverage Status' for "partially covered" cases
    if 'Coverage Status' not in df_excel.columns:
        df_excel['Coverage Status'] = ''

    # Iterate through each row in the Excel sheet to find employee matches
    for index, excel_row in df_excel.iterrows():
        try:
            employee_id = excel_row['ID'] 
        except KeyError:
            print(f"Error: 'ID' column not found in Excel row at index {index}. Available columns: {df_excel.columns.tolist()}")
            continue # Skip this row and proceed to the next

        # Find the corresponding employee in the CSV data using 'Employee Code'
        if 'Employee Code' not in df_csv.columns:
            print(f"Error: 'Employee Code' column not found in CSV file. Available columns: {df_csv.columns.tolist()}")
            break # Stop processing if crucial CSV column is missing

        matching_employee = df_csv[df_csv['Employee Code'] == employee_id]

        if not matching_employee.empty:
            # Get the status and other relevant information from the first match
            status = matching_employee['Status'].iloc[0]
            pid = matching_employee['PID(s)'].iloc[0]
            csv_leave_start_str = matching_employee['Leave Start'].iloc[0]
            csv_leave_end_str = matching_employee['Leave End'].iloc[0]

            # Format PID as a whole number if it's a float (e.g., 640968.0 -> 640968)
            formatted_pid = int(pid) if pd.notna(pid) else ''

            # Parse CSV leave dates if available, for use in both fully and partially covered cases
            leave_start_csv_dt = None
            leave_end_csv_dt = None
            if pd.notna(csv_leave_start_str) and pd.notna(csv_leave_end_str):
                try:
                    leave_start_csv_dt = datetime.strptime(str(csv_leave_start_str), '%d/%m/%Y')
                    leave_end_csv_dt = datetime.strptime(str(csv_leave_end_str), '%d/%m/%Y')
                except ValueError:
                    print(f"Warning: Could not parse raw leave dates '{csv_leave_start_str}', '{csv_leave_end_str}' for Employee ID {employee_id}.")

            # Construct original leave period text for use in status messages
            original_leave_period_text = ""
            if leave_start_csv_dt and leave_end_csv_dt:
                original_leave_period_text = f" (Original Leave: {leave_start_csv_dt.strftime('%d %B %Y')} - {leave_end_csv_dt.strftime('%d %B %Y')})"

            if status == 'no leave':
                # If status is "no leave", fill with the reference period
                df_excel.at[index, 'Start Date'] = REFERENCE_START_DATE.strftime('%d %B %Y')
                df_excel.at[index, 'End Date'] = REFERENCE_END_DATE.strftime('%d %B %Y')
                df_excel.at[index, 'Coverage Status'] = '' # Clear if previously set
            elif status == 'fully covered':
                # If status is "fully covered", mention PID and original leave dates in 'Start Date'
                df_excel.at[index, 'Start Date'] = f"{formatted_pid}{original_leave_period_text}" if formatted_pid else ''
                df_excel.at[index, 'End Date'] = '' # Clear 'End Date'
                df_excel.at[index, 'Coverage Status'] = '' # Clear if previously set
            elif status == 'partially covered':
                if leave_start_csv_dt and leave_end_csv_dt:
                    try:
                        # Generate all days in the reference period
                        ref_days = set()
                        current_day = REFERENCE_START_DATE
                        while current_day <= REFERENCE_END_DATE:
                            ref_days.add(current_day)
                            current_day += timedelta(days=1)

                        # Generate all days covered by the employee's actual leave
                        leave_days_covered = set()
                        # Only add days to leave_days_covered if leave period is valid and overlaps with or touches reference period
                        if leave_start_csv_dt <= leave_end_csv_dt and \
                           not (leave_end_csv_dt < REFERENCE_START_DATE or leave_start_csv_dt > REFERENCE_END_DATE):
                            current_day = leave_start_csv_dt
                            while current_day <= leave_end_csv_dt:
                                leave_days_covered.add(current_day)
                                current_day += timedelta(days=1)
                        
                        # Find days in reference period that are NOT covered by the employee's leave
                        uncovered_days_in_reference = sorted(list(ref_days - leave_days_covered))

                        if uncovered_days_in_reference:
                            # Extract the first contiguous block of uncovered days
                            first_uncovered_start = uncovered_days_in_reference[0]
                            first_uncovered_end = first_uncovered_start
                            for i in range(1, len(uncovered_days_in_reference)):
                                if uncovered_days_in_reference[i] == first_uncovered_end + timedelta(days=1):
                                    first_uncovered_end = uncovered_days_in_reference[i]
                                else:
                                    # Found a gap, so the contiguous block ends here
                                    break
                            
                            df_excel.at[index, 'Start Date'] = first_uncovered_start.strftime('%d %B %Y')
                            df_excel.at[index, 'End Date'] = first_uncovered_end.strftime('%d %B %Y')
                            
                            # Provide more detail if there are multiple uncovered segments, and include PID
                            if len(uncovered_days_in_reference) > (first_uncovered_end - first_uncovered_start).days + 1:
                                df_excel.at[index, 'Coverage Status'] = f'Partially Covered (PID: {formatted_pid}{original_leave_period_text}, Multiple uncovered segments)'
                            else:
                                df_excel.at[index, 'Coverage Status'] = f'Partially Covered (PID: {formatted_pid}{original_leave_period_text})'

                        else:
                            # If no uncovered days were found, it means the existing partial leave
                            # fully covers the entire reference period [14-20 June].
                            df_excel.at[index, 'Start Date'] = '' # No additional days need covering
                            df_excel.at[index, 'End Date'] = ''
                            df_excel.at[index, 'Coverage Status'] = f'Fully Covered by Existing Partial Leave (PID: {formatted_pid}{original_leave_period_text})'

                    except ValueError:
                        print(f"Warning: Error during date calculation for Employee ID {employee_id}. Setting default dates.")
                        df_excel.at[index, 'Start Date'] = REFERENCE_START_DATE.strftime('%d %B %Y')
                        df_excel.at[index, 'End Date'] = REFERENCE_END_DATE.strftime('%d %B %Y')
                        df_excel.at[index, 'Coverage Status'] = f'Partially Covered (PID: {formatted_pid}{original_leave_period_text}, Calculation Error)'
                else:
                    # If leave start/end are missing in CSV for 'partially covered',
                    # assume the entire reference period is uncovered for them.
                    df_excel.at[index, 'Start Date'] = REFERENCE_START_DATE.strftime('%d %B %Y')
                    df_excel.at[index, 'End Date'] = REFERENCE_END_DATE.strftime('%d %B %Y')
                    df_excel.at[index, 'Coverage Status'] = f'Partially Covered (PID: {formatted_pid}{original_leave_period_text}, Dates Missing in CSV, assuming full reference uncovered)'
            else:
                # Handle unexpected status values
                print(f"Warning: Unknown status '{status}' for Employee ID {employee_id}. Skipping.")
                df_excel.at[index, 'Start Date'] = ''
                df_excel.at[index, 'End Date'] = ''
                df_excel.at[index, 'Coverage Status'] = ''
        else:
            # If employee ID is not found in the CSV, clear or leave blank
            df_excel.at[index, 'Start Date'] = ''
            df_excel.at[index, 'End Date'] = ''
            df_excel.at[index, 'Coverage Status'] = ''

    # Save the updated Excel file
    try:
        df_excel.to_excel(output_file_path, index=False, sheet_name='Sheet1')
        print(f"Successfully updated Excel file saved as: {output_file_path}")
    except Exception as e:
        print(f"Error saving updated Excel file '{output_file_path}': {e}")


# --- How to run the script ---
if __name__ == "__main__":
    # Corrected Excel input filename based on user clarification
    excel_input = 'AL HARAM PROJECTS.xlsx'
    csv_input = 'leave_analysis_report.csv'
    excel_output = 'AL HARAM PROJECTS_updated.xlsx' # Output will be a proper .xlsx file

    # Run the function
    update_excel_with_leave_data(excel_input, csv_input, excel_output)

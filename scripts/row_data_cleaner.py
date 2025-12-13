import re
import shutil
import os
import pandas as pd
import numpy as np
from datetime import datetime

#-------------------CONFIG--------------------
sourse_folder = r"D:\my_projects\project_p767\Taxi_management_db\data\Vendor_TripSheet_Report"
destination_folder = r"D:\my_projects\project_p767\Taxi_management_db\data\application_files"
os.makedirs(destination_folder, exist_ok=True)
PROCESSED_FOLDER = os.path.join(sourse_folder, "processed")


# --- HELPER FUNCTION: SAVE FORMATTED EXCEL ---
def save_formatted_excel(df, output_path):
    """Saves the dataframe with font size 13, row height 30, and auto-fitted columns."""
    try:
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # --- DEFINE FORMATS (Font Size 13 added here) ---
        base_format = {
            'font_size': 13,   # Set Font Size to 13
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        }

        # Header Format
        header_format = workbook.add_format({
            **base_format,     # Inherit font size 13
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#0070C0', 
            'font_color': '#FFFFFF'
        })
        
        # Data Formats
        date_format = workbook.add_format(base_format)
        time_format = workbook.add_format(base_format)
        datetime_format = workbook.add_format({
            **base_format,
            'num_format': 'dd-mm-yyyy hh:mm:ss' 
        })
        text_format = workbook.add_format(base_format)
        
        # Address Format (Left align + Text Wrap)
        address_format = workbook.add_format({
            **base_format, 
            'align': 'center', 
            'valign': 'vcenter',
            'text_wrap': True
        })

        # --- SET ROW HEIGHT ---
        # Set Header Row Height
        worksheet.set_row(0, 30) 
        
        # Set Data Row Heights (Loop through all rows with data)
        for row_idx in range(1, len(df) + 1):
            worksheet.set_row(row_idx, 30)

        # --- APPLY COLUMN WIDTHS & FORMATS ---
        for col_num, col_name in enumerate(df.columns):
            # Write header
            worksheet.write(0, col_num, col_name, header_format)
            
            # --- LOGIC TO DETERMINE WIDTH ---
            if 'ADDRESS' in str(col_name):
                column_width = 80
                current_format = address_format
            else:
                # Calculate max length of data to fit column
                max_data_len = df[col_name].astype(str).map(len).max()
                max_len = max(max_data_len, len(str(col_name))) + 2 
                column_width = min(max(max_len, 18), 50) # Increased min width slightly for larger font
                
                if 'DATE' in str(col_name):
                    current_format = date_format
                    column_width = 18 
                elif 'TIME' in str(col_name):
                    current_format = time_format
                    column_width = 12
                else:
                    current_format = text_format

            # Apply width and format
            worksheet.set_column(col_num, col_num, column_width, current_format)

        writer.close()
        print(f"SUCCESS: Saved {os.path.basename(output_path)}")
        
    except Exception as e:
        print(f"FAILED to save {os.path.basename(output_path)}: {e}")
# --- MAIN FUNCTION: CLEAN DATA ---
def clean_data(file_path, destination_folder):
    print(f"Processing: {os.path.basename(file_path)}")
    
    # 1. Load Data
    try:
        df = pd.read_excel(file_path, header=None) 
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    df.drop(index=1, inplace=True)
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # 2. Extract Trip ID
    df["Trip_ID"] = df[10].astype(str).apply(lambda x: x if x.startswith("T") else "")
    df["Trip_ID"] = df["Trip_ID"].replace("", pd.NA).ffill()

    # 3. Split Header vs Passenger Rows
    df[1] = df[1].astype(str)
    
    # Header rows (UNITED FACILITIES)
    df_headers = df[df[1].str.contains("UNITED FACILITIES", na=False)].copy()
    
    # Passenger rows (Start with 1-4)
    df_passengers = df[df[0].astype(str).str.match(r"^[12345]$")].copy()

    df_headers.reset_index(drop=True, inplace=True)
    df_passengers.reset_index(drop=True, inplace=True)

    # 4. Rename Columns
    header_mapping = {
        0: 'Trip_Date', 1: 'Agency_Name', 2: 'Driver_Login_Time', 3: 'Vehicle_No',
        4: 'Driver_Name', 5: 'Trip_Zone', 6: 'Driver_Mobile', 7: 'Marshall',
        8: 'Distance', 9: 'Emp_Count', 10: 'Trip_Count', 11: 'Trip_Sheet_ID_Raw'
    }
    df_headers_renamed = df_headers.rename(columns=header_mapping)

    passenger_mapping = {
        0: 'Pax_no', 1: 'Reporting_Time', 2: 'Employee_ID', 3: 'Employee_Name',
        4: 'Gender', 5: 'Emp_Category', 6: 'Flight_No.', 7: 'Address',
        8: 'Reporting_Location', 9: 'Landmark', 10: 'Passenger_Mobile', 11: 'Pax_Col_11_Empty'
    }
    df_passengers_renamed = df_passengers.rename(columns=passenger_mapping)

    # 5. Merge DataFrames
    final_df = pd.merge(df_passengers_renamed, df_headers_renamed, on='Trip_ID', how='left')

    # --- DATA CLEANING ---
    final_df['Trip_ID'] = final_df['Trip_ID'].astype(str).str.replace('T', '', regex=False)
    final_df['Vehicle_No'] = final_df['Vehicle_No'].astype(str).str.replace('-', '', regex=False)

    # Split Login Time
    split_data = final_df['Driver_Login_Time'].astype(str).str.strip().str.split(' ', n=1, expand=True)
    final_df['Direction'] = split_data[0]
    final_df['Shift_Time'] = pd.to_datetime(split_data[1], format='%H:%M', errors='coerce').dt.time
    final_df['Direction'] = final_df['Direction'].astype(str).str.replace('Login', 'Pickup', regex=False)
    final_df['Direction'] = final_df['Direction'].astype(str).str.replace('Logout', 'Drop', regex=False)
    final_df.loc[final_df['Pax_no'] == 2, 'Marshall'] = np.nan
    final_df['Marshall'] = final_df['Marshall'].astype(str).str.replace('MARSHALL', 'Guard', regex=False)


    # Handle Dates
    final_df['Trip_Date'] = pd.to_datetime(final_df['Trip_Date'], errors='coerce')
    final_df['Date'] = final_df['Trip_Date'].dt.date
    
    # Numeric Conversion
    cols_to_numeric = ['Trip_ID', 'Emp_Count', 'Employee_ID', 'Pax_no']
    final_df[cols_to_numeric] = final_df[cols_to_numeric].apply(pd.to_numeric, errors='coerce')

    final_df['Trip_Date'] = final_df['Trip_Date'].astype(str)

    # --- UPPERCASE & CLEANUP ---
    final_df.columns = final_df.columns.astype(str).str.strip().str.upper()
    str_cols = final_df.select_dtypes(include=['object']).columns
    final_df[str_cols] = final_df[str_cols].apply(lambda x: x.astype(str).str.strip().str.upper())

    # --- FINAL SELECTION ---
    desired_order = [
        'DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
        'GENDER', 'ADDRESS', 'LANDMARK', 'VEHICLE_NO', 'DIRECTION', 
        'SHIFT_TIME', 'TRIP_DATE', 'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION'
    ]
    final_df = final_df.reindex(columns=desired_order)

    # --- SAVE FILE (THIS WAS MISSING IN YOUR CODE) ---
  

    # --- SIMPLIFIED FILE NAMING ---
    
    # 1. Get the Date from the first row
    if 'DATE' in final_df.columns and not final_df['DATE'].empty:
        first_date = final_df['DATE'].iloc[0]
        
        # --- FIX: FORCE TO DATETIME ---
        # Whether it is a string or a date object, this ensures it works:
        date_str = pd.to_datetime(first_date).strftime('%d-%m-%Y')
    else:
        date_str = "Unknown_Date"

    # 2. Get the Direction from the first row
    if 'DIRECTION' in final_df.columns and not final_df['DIRECTION'].empty:
        direction_str = str(final_df['DIRECTION'].iloc[0]).capitalize()
    else:
        direction_str = "Report"

    # 3. Combine them
    # Result: "25-11-2025 Pickup.xlsx"
    output_filename = f"{date_str} {direction_str}.xlsx"
    output_path = os.path.join(destination_folder, output_filename)

    # 4. Save
    save_formatted_excel(final_df, output_path)
    shutil.move(file_path, os.path.join(PROCESSED_FOLDER, os.path.basename(file_path))) 
    print(f"Processed: {os.path.basename(file_path)}")
# --- EXECUTION LOOP ---
if __name__ == "__main__":
    print(f"Scanning folder: {sourse_folder}")
    files_found = 0
    for root, dirs, files in os.walk(sourse_folder):
        for file in files:
            if file.endswith(('.xls', '.xlsx')) and "final" not in root:
                files_found += 1
                file_path = os.path.join(root, file)
                clean_data(file_path, destination_folder)
    
    if files_found == 0:
        print("No Excel files found to process.")
    else:
        print("Processing complete.")
import re
import os
import pandas as pd
import numpy as np
import os
import io
import re
import shutil
import pandas as pd
import numpy as np
import datetime

# ------------------- CONFIGURATION --------------------
BASE_DIR = r"C:\Users\Ravi Pal\my_projects\project_p767\Taxi_management_db\data"
SOURCE_FOLDER = os.path.join(BASE_DIR, "manual_operation_data")
DESTINATION_FOLDER = os.path.join(BASE_DIR, "manul_files")
PROCESSED_FOLDER = os.path.join(SOURCE_FOLDER, "processed")

# Ensure directories exist
os.makedirs(DESTINATION_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

pd.set_option('future.no_silent_downcasting', True)

def clean_excel_file(file_path, filename):
    """
    Reads a raw Excel file, applies cleaning logic, and returns a clean DataFrame.
    """
    try:
        # --- STEP 1: LOAD DIRTY DATA ---
        raw_df = pd.read_excel(file_path, header=None)
        
        if raw_df.empty:
            print(f"Skipping empty file: {filename}")
            return None

        # The "CSV Trick"
        csv_buffer = io.StringIO()
        raw_df.to_csv(csv_buffer, index=False, header=False)
        csv_buffer.seek(0)
        df = pd.read_csv(csv_buffer, header=None)

        # Basic Cleanup
        df = df.replace(r'^\s*$', np.nan, regex=True)
        df = df.dropna(how='all').reset_index(drop=True)

        # --- STEP 2: EXTRACT METADATA ---
        # A. Reporting Location
        target_str = "EMPLOYEE ADDRESS"
        if 4 in df.columns:
            df["REPORTING_LOCATION"] = df[4].astype(str).apply(lambda x: x if target_str in x else np.nan)
            df["REPORTING_LOCATION"] = df["REPORTING_LOCATION"].str.extract(r'"([^"]*)"')
            df["REPORTING_LOCATION"] = df["REPORTING_LOCATION"].ffill()
        else:
            df["REPORTING_LOCATION"] = np.nan

        # B. Date from Filename
        date_match = re.search(r'(\d{2}-\d{2}-\d{4})', filename)
        file_date_str = date_match.group(1) if date_match else np.nan
        
        # --- STEP 3: RENAME & FILTER COLUMNS ---
        header_mapping = {
            0: 'TRIP_ID', 1: 'TRG_TYPE', 2: 'EMPLOYEE_ID', 3: 'EMPLOYEE_NAME',
            4: 'ADDRESS', 5: 'DRIVER_MOBILE', 6: 'CAB_4_DIGIT', 7: 'PICKUP_TIME',
            8: 'SHIFT_TIME', 9: 'MIS_REMARKS', 10: 'GENDER', 11: 'TRIP_SHEET_ID_RAW'
        }
        df = df.rename(columns=header_mapping)

        # Drop "Header" rows (EMP ID)
        if "EMPLOYEE_ID" in df.columns:
            df = df[~df["EMPLOYEE_ID"].astype(str).str.contains("EMP ID", na=False, regex=False)]

        # --- STEP 4: ENRICH DATA ---
        # A. Add Direction
        df["DIRECTION"] = "PICKUP"

        # B. Trip ID & No
        df["TRIP_ID"] = df["TRIP_ID"].ffill()
        df["TRIP_NO"] = "ROUTE NO : " + df["TRIP_ID"].astype(str)

        # C. Dates & Times
        date_obj = pd.to_datetime(file_date_str, dayfirst=True, errors='coerce')
        df['DATE'] = date_obj.date() if pd.notnull(date_obj) else np.nan

        df['SHIFT_TIME'] = df['SHIFT_TIME'].astype(str).str.strip()
        temp_date_str = date_obj.strftime('%Y-%m-%d') if pd.notnull(date_obj) else ""
        df['temp_combined'] = temp_date_str + ' ' + df['SHIFT_TIME']
        
        df['REPORTING_TIME'] = pd.to_datetime(df['temp_combined'], errors='coerce')

        # --- STEP 5: TYPE CONVERSION ---
        cols_to_numeric = ['EMPLOYEE_ID', 'CAB_4_DIGIT']
        for col in cols_to_numeric:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        str_cols = df.select_dtypes(include=['object']).columns
        df[str_cols] = df[str_cols].apply(lambda x: x.astype(str).str.strip().str.upper())

        # --- STEP 6: REORDER ---
        desired_order = [
            'DATE', 'TRIP_ID', 'TRIP_NO', 'DIRECTION', 'TRG_TYPE', 'EMPLOYEE_ID',
            'GENDER', 'EMPLOYEE_NAME', 'ADDRESS', 'CAB_4_DIGIT',
            'SHIFT_TIME', 'REPORTING_TIME', 'REPORTING_LOCATION', 'MIS_REMARKS'
        ]
        df = df.reindex(columns=desired_order)
        
        return df

    except Exception as e:
        print(f"Error processing {filename}: {e}")
        return None

# ------------------- MAIN EXECUTION --------------------
if __name__ == "__main__":
    print(f"Scanning folder: {SOURCE_FOLDER}...\n")
    
    files_processed_count = 0

    for root, dirs, files in os.walk(SOURCE_FOLDER):
        if "processed" in root:
            continue

        for file in files:
            if file.endswith(('.xls', '.xlsx')):
                file_path = os.path.join(root, file)
                print(f"Processing: {file}...")

                # 1. CLEAN THE DATA
                cleaned_df = clean_excel_file(file_path, file)

                # 2. SAVE INDIVIDUALLY
                if cleaned_df is not None and not cleaned_df.empty:
                    try:
                        # Grab the date from the first row
                        first_date = cleaned_df['DATE'].iloc[0]
                        
                        # --- FIX IS HERE ---
                        # We force convert it back to a datetime object before formatting
                        # This works whether 'first_date' is a String OR a Date object
                        if pd.notnull(first_date):
                            date_str = pd.to_datetime(first_date).strftime("%Y%m%d")
                            output_filename = f"manual_PICKUP_{date_str}.xlsx"
                        else:
                            output_filename = f"manual_PICKUP_UNKNOWN_{file}"
                        # -------------------

                        output_path = os.path.join(DESTINATION_FOLDER, output_filename)
                        
                        # Save the individual file
                        cleaned_df.to_excel(output_path, index=False)
                        print(f"   -> Saved to: {output_filename}")

                        # 3. MOVE RAW FILE TO PROCESSED FOLDER
                        shutil.move(file_path, os.path.join(PROCESSED_FOLDER, file))
                        print(f"   -> Moved raw file to 'processed'")
                        
                        files_processed_count += 1
                        
                    except Exception as save_err:
                        print(f"   -> Error saving/moving: {save_err}")
                else:
                    print(f"   -> Skipped (Empty or Error)")

    print("-" * 40)
    if files_processed_count == 0:
        print("No valid files found to process.")
    else:
        print(f"Processing complete. {files_processed_count} files saved individually.")
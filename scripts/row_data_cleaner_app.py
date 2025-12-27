import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta

# --- HELPER: SAVE BILLING EXCEL (SIMPLE) ---
def to_excel_billing(df):
    """
    Saves the dataframe for Billing:
    - Font size 13
    - Row height 30
    - No empty rows, no repeated headers
    """
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Define Formats
    base_format = workbook.add_format({
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    header_format = workbook.add_format({
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bold': True, 'text_wrap': True, 'fg_color': '#0070C0', 'font_color': '#FFFFFF'
    })
    
    # Set Row Heights
    worksheet.set_row(0, 30) # Header
    for row_idx in range(len(df)):
        worksheet.set_row(row_idx + 1, 30)

    # Set Column Widths
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        width = 80 if 'ADDRESS' in str(col_name) else 25
        worksheet.set_column(col_num, col_num, width, base_format)

    writer.close()
    output.seek(0)
    return output

# --- HELPER: SAVE OPERATIONS EXCEL (COMPLEX) ---
def to_excel_operations(df):
    """
    Saves the dataframe for Operations:
    - HOME TIME column added
    - Empty rows (Height 40)
    - Repeated Headers after empty rows
    """
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Write data (Index=False, Header=True initially)
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # --- FORMATS ---
    base_format = {
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    }
    
    # Normal Data
    data_format = workbook.add_format(base_format)
    
    # Header Format (Blue)
    header_format = workbook.add_format({
        **base_format,
        'bold': True, 'text_wrap': True, 'fg_color': '#0070C0', 'font_color': '#FFFFFF'
    })
    
    # Empty Row Format (White, no border)
    empty_row_format = workbook.add_format({'font_size': 13, 'border': 0})

    # --- APPLY ROW HEIGHTS & FORMATTING ---
    # 1. Format the very first header row
    worksheet.set_row(0, 30)
    
    # 2. Loop through all data rows to check if they are Data, Empty, or Repeated Header
    # We use the 'TRIP_ID' column to detect what kind of row it is
    trip_id_col_idx = df.columns.get_loc("TRIP_ID") if "TRIP_ID" in df.columns else 0

    for row_idx in range(len(df)):
        excel_row = row_idx + 1
        
        # Get the value in the TRIP_ID column for this row
        cell_value = df.iloc[row_idx, trip_id_col_idx]
        
        # LOGIC:
        # If NaN -> It is an Spacer Row
        # If value == "TRIP_ID" -> It is a Repeated Header
        # Otherwise -> It is Data
        
        if pd.isna(cell_value):
            # Empty Row
            worksheet.set_row(excel_row, 40, empty_row_format)
        elif str(cell_value) == "TRIP_ID":
            # Repeated Header Row
            worksheet.set_row(excel_row, 30)
            # Apply header format to this entire row
            for col in range(len(df.columns)):
                worksheet.write(excel_row, col, df.iloc[row_idx, col], header_format)
        else:
            # Normal Data Row
            worksheet.set_row(excel_row, 30)

    # --- COLUMN WIDTHS ---
    for col_num, col_name in enumerate(df.columns):
        # Apply standard header format to top row
        worksheet.write(0, col_num, col_name, header_format)
        
        width = 80 if 'ADDRESS' in str(col_name) else 25
        worksheet.set_column(col_num, col_num, width, data_format)

    writer.close()
    output.seek(0)
    return output

# --- MAIN LOGIC ---
def process_data(uploaded_file):
    # 1. Load Data
    try:
        df = pd.read_excel(uploaded_file, header=None)
    except Exception as e:
        # Fallback for .xls files if xlrd is installed
        try:
            df = pd.read_excel(uploaded_file, header=None, engine='xlrd')
        except:
            return None, None, f"Error: {e}"

    df.drop(index=1, inplace=True)
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # 2. Extract Trip ID
    df["Trip_ID"] = df[10].astype(str).apply(lambda x: x if x.startswith("T") else "")
    df["Trip_ID"] = df["Trip_ID"].replace("", pd.NA).ffill()

    # 3. Split Header vs Passenger Rows
    df[1] = df[1].astype(str)
    df_headers = df[df[1].str.contains("UNITED FACILITIES", na=False)].copy()
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

    # 5. Merge
    final_df = pd.merge(df_passengers_renamed, df_headers_renamed, on='Trip_ID', how='left')

    # --- CLEANING ---
    final_df['Trip_ID'] = final_df['Trip_ID'].astype(str).str.replace('T', '', regex=False)
    final_df['Vehicle_No'] = final_df['Vehicle_No'].astype(str).str.replace('-', '', regex=False)

    split_data = final_df['Driver_Login_Time'].astype(str).str.strip().str.split(' ', n=1, expand=True)
    final_df['Direction'] = split_data[0]
    
    # Store Shift Time as datetime object for calculation
    final_df['Shift_Time_Obj'] = pd.to_datetime(split_data[1], format='%H:%M', errors='coerce')
    final_df['Shift_Time'] = final_df['Shift_Time_Obj'].dt.time
    
    final_df['Direction'] = final_df['Direction'].astype(str).str.replace('Login', 'Pickup', regex=False)
    final_df['Direction'] = final_df['Direction'].astype(str).str.replace('Logout', 'Drop', regex=False)
    final_df.loc[final_df['Pax_no'] == 2, 'Marshall'] = np.nan
    final_df['Marshall'] = final_df['Marshall'].astype(str).str.replace('MARSHALL', 'Guard', regex=False)

    # --- DATE CLEANING (FIX FOR "JUST DATE") ---
    # 1. Convert to proper datetime first
    final_df['Trip_Date'] = pd.to_datetime(final_df['Trip_Date'], errors='coerce')
    
    # 2. Create the file naming string BEFORE we force it to string format (safest)
    try:
        date_val = final_df['Trip_Date'].iloc[0].strftime('%d-%m-%Y')
    except:
        date_val = "Unknown_Date"
    
    # 3. FORCE Trip_Date to be a String in DD-MM-YYYY format
    # This removes any time component permanently
    final_df['Trip_Date'] = final_df['Trip_Date'].dt.strftime('%d-%m-%Y')
    
    # Clean up strings
    final_df.columns = final_df.columns.astype(str).str.strip().str.upper()
    str_cols = final_df.select_dtypes(include=['object']).columns
    final_df[str_cols] = final_df[str_cols].apply(lambda x: x.astype(str).str.strip().str.upper())

    # --- FILE NAMING ---
    if 'DATE' in final_df.columns and not final_df['DATE'].empty:
         # Attempt to find date in Trip_Date column since DATE column might be empty initially
        try:
            date_val = pd.to_datetime(final_df['TRIP_DATE'].iloc[0]).strftime('%d-%m-%Y')
        except:
            date_val = "Unknown_Date"
    else:
        date_val = "Unknown_Date"

    if 'DIRECTION' in final_df.columns and not final_df['DIRECTION'].empty:
        dir_val = str(final_df['DIRECTION'].iloc[0]).capitalize()
    else:
        dir_val = "Report"
        
    base_filename = f"{date_val} {dir_val}"

    # --- PREPARE BILLING DATAFRAME (Clean, no extra cols) ---
    billing_cols = [
        'TRIP_DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
        'GENDER', 'ADDRESS','PASSENGER_MOBILE', 'LANDMARK', 'VEHICLE_NO', 'DIRECTION', 
        'SHIFT_TIME', 'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION'
    ]
    # Filter columns that actually exist
    billing_cols = [c for c in billing_cols if c in final_df.columns]
    billing_df = final_df[billing_cols].copy()

    # --- PREPARE OPERATIONS DATAFRAME (Home Time + Gaps + Headers) ---
    ops_df = final_df.copy()
    
    # 1. Add HOME TIME (Shift Time - 2 Hours)

    # We use the temp object column we created earlier
    ops_df['HOME_TIME'] = (ops_df['SHIFT_TIME_OBJ'] - timedelta(hours=2)).dt.time
    
    # Select Cols
    ops_cols = [
        'TRIP_DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
        'GENDER', 'ADDRESS','PASSENGER_MOBILE', 'LANDMARK', 'VEHICLE_NO', 'DIRECTION', 
        'HOME_TIME', 'SHIFT_TIME', 'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION'
    ]
    ops_cols = [c for c in ops_cols if c in ops_df.columns]
    ops_df = ops_df[ops_cols]

    # 2. Insert Empty Rows AND Headers
    df_list = []
    empty_row = pd.DataFrame([[np.nan] * len(ops_df.columns)], columns=ops_df.columns)
    
    # Create a dataframe that looks like the header
    header_row = pd.DataFrame([ops_df.columns.values], columns=ops_df.columns)
    
    groups = list(ops_df.groupby('TRIP_ID', sort=False))
    
    for i, (trip_id, group) in enumerate(groups):
        df_list.append(group)
        if i < len(groups) - 1:
            df_list.append(empty_row)  # 40px gap
            df_list.append(header_row) # Repeated Header
            
    ops_final = pd.concat(df_list, ignore_index=True)

    return billing_df, ops_final, base_filename

# --- STREAMLIT UI ---
st.title("Air India TripSheet Cleaner")
st.write("Upload Raw File -> Get separate files for Billing and Operations")

uploaded_file = st.file_uploader("Choose an Excel file", type=['xls', 'xlsx'])

if uploaded_file is not None:
    with st.spinner('Processing...'):
        billing_df, ops_df, filename = process_data(uploaded_file)
        
        if billing_df is not None:
            st.success("File processed successfully!")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("1. Billing Team")
                st.write("Clean data, no empty rows.")
                billing_buffer = to_excel_billing(billing_df)
                st.download_button(
                    label="Download Billing File",
                    data=billing_buffer,
                    file_name=f"BILLING_{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                st.subheader("2. Operations Team")
                st.write("Includes HOME TIME (-2hr) & Headers after gaps.")
                ops_buffer = to_excel_operations(ops_df)
                st.download_button(
                    label="Download Ops File",
                    data=ops_buffer,
                    file_name=f"OPS_{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            st.write("---")
            st.write("Preview (Billing Data):")
            st.dataframe(billing_df.head())
        else:
            st.error(filename) # Displays error message
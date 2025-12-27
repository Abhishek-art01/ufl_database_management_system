import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta

# --- HELPER: SAVE BILLING EXCEL (STRICT FORMATTING + TIME FIX) ---
def to_excel_billing(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook  = writer.book
    
    # Add Sheet
    worksheet = workbook.add_worksheet('Sheet1')
    
    # --- FORMATS ---
    base_format_props = {
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    }
    base_format = workbook.add_format(base_format_props)
    
    # Time Format (Fixes 0.222 issue)
    time_format = workbook.add_format({
        **base_format_props,
        'num_format': 'hh:mm'
    })
    
    header_format = workbook.add_format({
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bold': True, 'text_wrap': True, 'fg_color': '#0070C0', 'font_color': '#FFFFFF'
    })

    # --- WRITE HEADER (Row 0) ---
    worksheet.set_row(0, 30) # Header Height
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        # Set Width (Format=None to avoid infinite borders)
        width = 80 if 'ADDRESS' in str(col_name) else 25
        worksheet.set_column(col_num, col_num, width, None)

    # --- WRITE DATA (Rows 1 to N) ---
    for row_idx, row_data in df.iterrows():
        excel_row = row_idx + 1
        worksheet.set_row(excel_row, 30) # Data Height
        
        for col_num, value in enumerate(row_data):
            col_name = df.columns[col_num]
            
            # Apply Time Format to specific columns
            if col_name in ['SHIFT_TIME', 'HOME_TIME']:
                cell_format = time_format
            else:
                cell_format = base_format
            
            val_to_write = value if pd.notna(value) else ""
            worksheet.write(excel_row, col_num, val_to_write, cell_format)

    writer.close()
    output.seek(0)
    return output

# --- HELPER: SAVE OPERATIONS EXCEL (STRICT FORMATTING + TIME FIX) ---
def to_excel_operations(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook  = writer.book
    
    # Add Sheet
    worksheet = workbook.add_worksheet('Sheet1')
    
    # --- FORMATS ---
    base_format_props = {
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    }
    base_format = workbook.add_format(base_format_props)
    
    # Time Format (Fixes 0.222 issue)
    time_format = workbook.add_format({
        **base_format_props,
        'num_format': 'hh:mm'
    })
    
    header_format = workbook.add_format({
        'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bold': True, 'text_wrap': True, 'fg_color': '#0070C0', 'font_color': '#FFFFFF'
    })
    
    # --- SETUP COLUMNS (Width Only, No Styles) ---
    for col_num, col_name in enumerate(df.columns):
        width = 80 if 'ADDRESS' in str(col_name) else 25
        worksheet.set_column(col_num, col_num, width, None)

    # --- WRITE HEADER (Row 0) ---
    worksheet.set_row(0, 30)
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)

    # --- WRITE DATA (Rows 1 to N) ---
    trip_id_col_idx = df.columns.get_loc("TRIP_ID") if "TRIP_ID" in df.columns else 0

    for row_idx, row_data in df.iterrows():
        excel_row = row_idx + 1
        
        # Check Trip ID to see if this is Data, Spacer, or Header
        cell_value = row_data.iloc[trip_id_col_idx]
        
        if pd.isna(cell_value):
            # --- SPACER ROW ---
            worksheet.set_row(excel_row, 40)
            
        elif str(cell_value) == "TRIP_ID":
            # --- REPEATED HEADER ---
            worksheet.set_row(excel_row, 30)
            for col_num, value in enumerate(row_data):
                worksheet.write(excel_row, col_num, value, header_format)
                
        else:
            # --- NORMAL DATA ---
            worksheet.set_row(excel_row, 30)
            for col_num, value in enumerate(row_data):
                col_name = df.columns[col_num]
                
                # Apply Time Format to specific columns
                if col_name in ['SHIFT_TIME', 'HOME_TIME']:
                    cell_format = time_format
                else:
                    cell_format = base_format
                    
                val_to_write = value if pd.notna(value) else ""
                worksheet.write(excel_row, col_num, val_to_write, cell_format)

    writer.close()
    output.seek(0)
    return output

# --- MAIN LOGIC ---
def process_data(uploaded_file):
    # 1. Load Data
    try:
        df = pd.read_excel(uploaded_file, header=None)
    except Exception as e:
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
    
    final_df['Shift_Time_Obj'] = pd.to_datetime(split_data[1], format='%H:%M', errors='coerce')
    final_df['Shift_Time'] = final_df['Shift_Time_Obj'].dt.time
    
    final_df['Direction'] = final_df['Direction'].astype(str).str.replace('Login', 'Pickup', regex=False)
    final_df['Direction'] = final_df['Direction'].astype(str).str.replace('Logout', 'Drop', regex=False)
    final_df.loc[final_df['Pax_no'] == 2, 'Marshall'] = np.nan
    final_df['Marshall'] = final_df['Marshall'].astype(str).str.replace('MARSHALL', 'Guard', regex=False)

    # --- DATE CLEANING ---
    final_df['Trip_Date'] = pd.to_datetime(final_df['Trip_Date'], errors='coerce')
    try:
        date_val = final_df['Trip_Date'].iloc[0].strftime('%d-%m-%Y')
    except:
        date_val = "Unknown_Date"
    final_df['Trip_Date'] = final_df['Trip_Date'].dt.strftime('%d-%m-%Y')
    
    # Clean Strings
    final_df.columns = final_df.columns.astype(str).str.strip().str.upper()
    str_cols = final_df.select_dtypes(include=['object']).columns
    final_df[str_cols] = final_df[str_cols].apply(lambda x: x.astype(str).str.strip().str.upper())

    # --- SORTING ---
    if 'SHIFT_TIME' in final_df.columns:
        final_df = final_df.sort_values(by='SHIFT_TIME')

    # --- FILE NAMING ---
    if 'DIRECTION' in final_df.columns and not final_df['DIRECTION'].empty:
        dir_val = str(final_df['DIRECTION'].iloc[0]).capitalize()
    else:
        dir_val = "Report"
    base_filename = f"{date_val} {dir_val}"

    # --- PREPARE DATASETS ---
    billing_cols = [
        'TRIP_DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
        'GENDER', 'ADDRESS','PASSENGER_MOBILE', 'LANDMARK', 'VEHICLE_NO', 'DIRECTION', 
        'SHIFT_TIME', 'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION'
    ]
    billing_cols = [c for c in billing_cols if c in final_df.columns]
    billing_df = final_df[billing_cols].copy()

    # Ops Data
    ops_df = final_df.copy()
    ops_df['HOME_TIME'] = (ops_df['SHIFT_TIME_OBJ'] - timedelta(hours=2)).dt.time
    
    ops_cols = [
        'TRIP_DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
        'GENDER', 'ADDRESS','PASSENGER_MOBILE', 'LANDMARK', 'VEHICLE_NO', 'DIRECTION', 
        'HOME_TIME', 'SHIFT_TIME', 'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION'
    ]
    ops_cols = [c for c in ops_cols if c in ops_df.columns]
    ops_df = ops_df[ops_cols]

    # Insert Gaps & Headers
    df_list = []
    empty_row = pd.DataFrame([[np.nan] * len(ops_df.columns)], columns=ops_df.columns)
    header_row = pd.DataFrame([ops_df.columns.values], columns=ops_df.columns)
    
    groups = list(ops_df.groupby('TRIP_ID', sort=False))
    
    for i, (trip_id, group) in enumerate(groups):
        df_list.append(group)
        if i < len(groups) - 1:
            df_list.append(empty_row)  # Spacer
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
                billing_buffer = to_excel_billing(billing_df)
                st.download_button(
                    label="Download Billing File",
                    data=billing_buffer,
                    file_name=f"BILLING_{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                st.subheader("2. Operations Team")
                ops_buffer = to_excel_operations(ops_df)
                st.download_button(
                    label="Download Ops File",
                    data=ops_buffer,
                    file_name=f"OPS_{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            st.write("---")
            st.dataframe(billing_df.head())
        else:
            st.error(filename)
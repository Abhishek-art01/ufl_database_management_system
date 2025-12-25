import streamlit as st
import pandas as pd
import numpy as np
import io

# --- HELPER: SAVE FORMATTED EXCEL ---
def to_excel_with_formatting(df):
    """
    Saves the dataframe to a memory buffer with:
    - Font size 13
    - Data row height 30
    - Empty separator row height 40
    - Auto-fitted columns
    """
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Write data to sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # --- DEFINE FORMATS ---
    base_format_props = {
        'font_size': 13,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    }

    # Standard formats
    header_format = workbook.add_format({
        **base_format_props,
        'bold': True,
        'text_wrap': True,
        'fg_color': '#0070C0', 
        'font_color': '#FFFFFF'
    })
    
    text_format = workbook.add_format(base_format_props)
    date_format = workbook.add_format(base_format_props) # You can add num_format if needed
    time_format = workbook.add_format(base_format_props)
    
    # Address Format (Left align + Text Wrap)
    address_format = workbook.add_format({
        **base_format_props, 
        'text_wrap': True
    })

    # Empty Row Format (No border, just white)
    empty_row_format = workbook.add_format({
        'font_size': 13,
        'border': 0
    })

    # --- APPLY ROW HEIGHTS ---
    # Header Height
    worksheet.set_row(0, 30)
    
    # Loop through data to set heights based on content
    # We check if 'TRIP_ID' is NaN to identify the empty rows we added
    trip_id_col_idx = df.columns.get_loc("TRIP_ID") if "TRIP_ID" in df.columns else 0

    for row_idx in range(len(df)):
        excel_row = row_idx + 1 # +1 because header is 0
        
        # Check if this is one of our inserted empty rows (Trip ID is null)
        is_empty_row = pd.isna(df.iloc[row_idx, trip_id_col_idx])
        
        if is_empty_row:
            worksheet.set_row(excel_row, 40) # Requirement: Height 40 for empty rows
        else:
            worksheet.set_row(excel_row, 30) # Standard height 30

    # --- APPLY COLUMN WIDTHS & FORMATS ---
    for col_num, col_name in enumerate(df.columns):
        # Write header
        worksheet.write(0, col_num, col_name, header_format)
        
        # Determine width and format
        if 'ADDRESS' in str(col_name):
            column_width = 80
            current_format = address_format
        else:
            # Calculate max length
            max_data_len = df[col_name].astype(str).map(len).max()
            max_len = max(max_data_len, len(str(col_name))) + 2 
            column_width = min(max(max_len, 18), 50)
            
            if 'DATE' in str(col_name):
                current_format = date_format
                column_width = 18 
            elif 'TIME' in str(col_name):
                current_format = time_format
                column_width = 12
            else:
                current_format = text_format

        # Apply column format
        worksheet.set_column(col_num, col_num, column_width, current_format)

    writer.close()
    output.seek(0)
    return output

# --- MAIN LOGIC: PROCESS DATA ---
def process_data(uploaded_file):
    # 1. Load Data
    try:
        df = pd.read_excel(uploaded_file, header=None)
    except Exception as e:
        return None, f"Error reading file: {e}"

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
        'GENDER', 'ADDRESS', 'VEHICLE_NO','SHIFT_TIME', 'TRIP_DATE','DRIVER_NAME', 'LANDMARK', 'DIRECTION', 
        'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION'
    ]
    # Ensure all columns exist before reindexing
    existing_cols = [c for c in desired_order if c in final_df.columns]
    final_df = final_df[existing_cols]

    # --- NEW FEATURE: INSERT EMPTY ROW AFTER EACH TRIP_ID ---
    # We will create a list of dataframes and concatenate them with an empty row in between
    df_list = []
    
    # Create an empty row with the same columns, filled with NaN
    empty_row = pd.DataFrame([[np.nan] * len(final_df.columns)], columns=final_df.columns)
    
    # Group by Trip ID (preserve order with sort=False)
    groups = list(final_df.groupby('TRIP_ID', sort=False))
    
    for i, (trip_id, group) in enumerate(groups):
        df_list.append(group)
        # Add empty row if it's not the very last group
        if i < len(groups) - 1:
            df_list.append(empty_row)
            
    # Combine back into one dataframe
    final_df_with_gaps = pd.concat(df_list, ignore_index=True)

    # --- DETERMINE FILE NAME PARTS ---
    if 'DATE' in final_df.columns and not final_df['DATE'].dropna().empty:
        first_date = final_df['DATE'].dropna().iloc[0]
        date_str = pd.to_datetime(first_date).strftime('%d-%m-%Y')
    else:
        date_str = "Unknown_Date"

    if 'DIRECTION' in final_df.columns and not final_df['DIRECTION'].dropna().empty:
        direction_str = str(final_df['DIRECTION'].dropna().iloc[0]).capitalize()
    else:
        direction_str = "Report"

    output_filename = f"{date_str} {direction_str}.xlsx"

    return final_df_with_gaps, output_filename

# --- STREAMLIT UI ---
st.title("Air India TripSheet Cleaner")
st.write("Upload the raw TripSheet Excel file to process it.")

uploaded_file = st.file_uploader("Choose an Excel file", type=['xls', 'xlsx'])

if uploaded_file is not None:
    with st.spinner('Processing data...'):
        # Run processing
        processed_df, file_name = process_data(uploaded_file)
        
        if processed_df is not None:
            # Show preview
            st.success("File processed successfully!")
            st.write("Preview of cleaned data (first 10 rows):")
            st.dataframe(processed_df.head(10))
            
            # Convert to formatted Excel in memory
            excel_buffer = to_excel_with_formatting(processed_df)
            
            # Download Button
            st.download_button(
                label=f"Download {file_name}",
                data=excel_buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error(f"Failed to process file. {file_name}") # file_name holds error msg here
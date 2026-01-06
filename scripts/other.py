import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import timedelta

# --- 1. EXCEL FORMATTER CLASS ---
class ExcelFormatter:
    def __init__(self, df):
        self.df = df
        self.output = io.BytesIO()
        self.writer = pd.ExcelWriter(self.output, engine='xlsxwriter')
        self.workbook = self.writer.book
        self.worksheet = self.workbook.add_worksheet('Sheet1')
        
        # Styles
        self.base_fmt = self.workbook.add_format({
            'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        self.header_fmt = self.workbook.add_format({
            'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'bold': True, 'text_wrap': True, 'fg_color': '#0070C0', 'font_color': '#FFFFFF'
        })
        self.time_fmt = self.workbook.add_format({
            'font_size': 13, 'border': 1, 'align': 'center', 'valign': 'vcenter', 
            'text_wrap': True, 'num_format': 'hh:mm'
        })

    def set_column_widths(self, mode='BILLING'):
        if mode == 'BILLING':
            for i, col in enumerate(self.df.columns):
                c_upper = str(col).upper()
                width = 80 if 'ADDRESS' in c_upper else (40 if 'EMPLOYEE_NAME' in c_upper else 25)
                self.worksheet.set_column(i, i, width)
        
        elif mode == 'OPS':
            width_map = {
                'TRIP_DATE': 13, 'TRIP_ID': 11, 'FLIGHT_NO.': 13, 'EMPLOYEE_ID': 12,
                'EMPLOYEE_NAME': 23, 'ADDRESS': 110, 'PASSENGER_MOBILE': 14.5,
                'LANDMARK': 22, 'REPORTING_LOCATION': 14, 'VEHICLE_NO': 15,
                'DIRECTION': 12, 'PICKUP POINT': 11, 'SHIFT_TIME': 13, 'MARSHALL': 15
            }
            for i, col in enumerate(self.df.columns):
                self.worksheet.set_column(i, i, width_map.get(str(col).upper(), 20))

    def write_data(self, mode='BILLING'):
        self.worksheet.set_row(0, 30 if mode == 'BILLING' else 50)
        for col_num, val in enumerate(self.df.columns):
            self.worksheet.write(0, col_num, val, self.header_fmt)

        trip_id_idx = self.df.columns.get_loc("TRIP_ID") if "TRIP_ID" in self.df.columns else 0
        
        for row_idx, row in self.df.iterrows():
            excel_row = row_idx + 1
            cell_val = row.iloc[trip_id_idx]
            
            # Ops Spacer/Header logic
            if mode == 'OPS' and pd.isna(cell_val): 
                self.worksheet.set_row(excel_row, 40)
                continue
            elif mode == 'OPS' and str(cell_val) == "TRIP_ID":
                self.worksheet.set_row(excel_row, 50)
                for c, val in enumerate(row):
                    self.worksheet.write(excel_row, c, val, self.header_fmt)
                continue

            self.worksheet.set_row(excel_row, 30 if mode == 'BILLING' else 45)
            for col_num, val in enumerate(row):
                col_name = self.df.columns[col_num]
                fmt = self.time_fmt if col_name in ['SHIFT_TIME', 'PICKUP POINT', 'HOME_TIME'] else self.base_fmt
                val_clean = val if pd.notna(val) else ""
                self.worksheet.write(excel_row, col_num, val_clean, fmt)

    def get_file(self):
        self.writer.close()
        self.output.seek(0)
        return self.output

# --- 2. DATA PROCESSING LOGIC ---
def clean_dataframe(df):
    df.columns = df.columns.astype(str).str.strip().str.upper()
    str_cols = df.select_dtypes(include=['object']).columns
    df[str_cols] = df[str_cols].apply(lambda x: x.astype(str).str.strip().str.upper())
    return df

def process_data(uploaded_file):
    try:
        try:
            df = pd.read_excel(uploaded_file, header=None)
        except:
            df = pd.read_excel(uploaded_file, header=None, engine='xlrd')

        # Cleanup basic
        df.drop(index=1, inplace=True)
        df.dropna(how="all", inplace=True)
        df.reset_index(drop=True, inplace=True)

        # ---------------------------------------------------------
        # UPDATED LOGIC: Find Trip ID Column Dynamically
        # ---------------------------------------------------------
        # We look for a column that has "T" followed by numbers (e.g., T3599973)
        trip_col_idx = 10 # Default fallback
        for col in df.columns:
            # Check if any value in this column starts with 'T' and has digits
            sample = df[col].astype(str).head(10)
            if sample.str.contains(r'^T\d+', na=False).any():
                trip_col_idx = col
                break
        
        # Create a clean Trip ID column
        df["Trip_ID_Clean"] = df[trip_col_idx].astype(str).apply(
            lambda x: x if str(x).startswith("T") else pd.NA
        ).ffill()

        # ---------------------------------------------------------
        # UPDATED LOGIC: Identify Rows Agnostically
        # ---------------------------------------------------------
        # Instead of looking for "UNITED FACILITIES", we look for "Log" in column 2 
        # (matches "Login", "Logout", "Login 00:00", "Logout 00:00")
        mask_headers = df[2].astype(str).str.contains("LOG", na=False, case=False)
        
        # Passenger rows are those where Column 0 is a number (1, 2, 3...)
        mask_pax = df[0].astype(str).str.match(r"^\d+$")

        df_headers = df[mask_headers].rename(columns={
            0: 'Trip_Date', 1: 'Agency_Name', 2: 'Driver_Login_Time', 3: 'Vehicle_No',
            4: 'Driver_Name', 5: 'Trip_Zone', 6: 'Driver_Mobile', 7: 'Marshall',
            8: 'Distance', 9: 'Emp_Count', 10: 'Trip_Count', 11: 'Trip_Sheet_ID_Raw'
        })
        
        df_pax = df[mask_pax].rename(columns={
            0: 'Pax_no', 1: 'Reporting_Time', 2: 'Employee_ID', 3: 'Employee_Name',
            4: 'Gender', 5: 'Emp_Category', 6: 'Flight_No.', 7: 'Address',
            8: 'Reporting_Location', 9: 'Landmark', 10: 'Passenger_Mobile'
        })

        # Merge
        final = pd.merge(df_pax, df_headers[['Trip_ID_Clean', 'Trip_Date', 'Driver_Login_Time', 'Vehicle_No', 'Marshall']], 
                         left_on='Trip_ID_Clean', right_on='Trip_ID_Clean', how='left')
        
        final.rename(columns={'Trip_ID_Clean': 'Trip_ID'}, inplace=True)
        final['Trip_ID'] = final['Trip_ID'].str.replace('T', '', regex=False)
        final['Vehicle_No'] = final['Vehicle_No'].astype(str).str.replace('-', '', regex=False)
        
        # ---------------------------------------------------------
        # UPDATED LOGIC: Robust Time Split
        # ---------------------------------------------------------
        split_data = final['Driver_Login_Time'].astype(str).str.strip().str.split(expand=True)
        
        if split_data.shape[1] > 1:
            final['Direction'] = split_data[0]
            final['Shift_Time_Obj'] = pd.to_datetime(split_data[1], format='%H:%M', errors='coerce')
        else:
            final['Direction'] = split_data[0] if not split_data.empty else ""
            final['Shift_Time_Obj'] = pd.NaT

        final['Shift_Time'] = final['Shift_Time_Obj'].dt.time
        
        # Normalize Directions
        final['Direction'] = final['Direction'].astype(str).str.upper()
        final['Direction'] = final['Direction'].replace({
            'LOGIN': 'PICKUP', 
            'LOGOUT': 'DROP'
        }, regex=True)
        
        # Marshall cleanup
        final.loc[final['Pax_no'] == 2, 'Marshall'] = ''

        # Date Format
        final['Trip_Date'] = pd.to_datetime(final['Trip_Date'], errors='coerce').dt.strftime('%d-%m-%Y')
        
        # Final Polish
        final = clean_dataframe(final)
        if 'SHIFT_TIME' in final.columns:
            final.sort_values('SHIFT_TIME', inplace=True)

        # --- PREPARE OUTPUTS ---
        # Billing
        billing_cols = ['TRIP_DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
                        'GENDER', 'ADDRESS', 'PASSENGER_MOBILE', 'LANDMARK', 'VEHICLE_NO', 
                        'DIRECTION', 'SHIFT_TIME', 'EMP_COUNT', 'PAX_NO', 'MARSHALL', 'REPORTING_LOCATION']
        billing_out = final[[c for c in billing_cols if c in final.columns]].copy()
        
        # Ops
        ops_df = final.copy()
        ops_df['PICKUP POINT'] = (ops_df['SHIFT_TIME_OBJ'] - timedelta(hours=2)).dt.time
        if 'MARSHALL' in ops_df.columns:
            ops_df['MARSHALL'] = ops_df['MARSHALL'].fillna('')
            
        ops_cols = ['TRIP_DATE', 'TRIP_ID', 'FLIGHT_NO.', 'EMPLOYEE_ID', 'EMPLOYEE_NAME', 
                    'ADDRESS', 'LANDMARK', 'REPORTING_LOCATION', 'PASSENGER_MOBILE', 
                    'VEHICLE_NO', 'DIRECTION', 'PICKUP POINT', 'SHIFT_TIME', 'MARSHALL']
        ops_df = ops_df[[c for c in ops_cols if c in ops_df.columns]]

        # Ops Spacers
        df_list = []
        header_row = pd.DataFrame([ops_df.columns], columns=ops_df.columns)
        empty_row = pd.DataFrame([[np.nan]*len(ops_df.columns)], columns=ops_df.columns)
        
        for i, (tid, group) in enumerate(ops_df.groupby('TRIP_ID', sort=False)):
            df_list.append(group)
            if i < len(ops_df['TRIP_ID'].unique()) - 1:
                df_list.append(empty_row)
                df_list.append(header_row)
                
        ops_out = pd.concat(df_list, ignore_index=True)
        
        filename = f"{final['TRIP_DATE'].iloc[0]} {final['DIRECTION'].iloc[0].title()}" if not final.empty else "Processed_File"

        return billing_out, ops_out, filename

    except Exception as e:
        return None, None, f"Error: {str(e)}"

# --- 3. STREAMLIT APP ---
st.set_page_config(page_title="TripSheet Manager", layout="wide")
st.title("✈️ Universal TripSheet Cleaner")
st.markdown("Works with **J Travels, United Facilities, Bajaj**, and others.")

uploaded_file = st.file_uploader("Drop Excel File Here", type=['xls', 'xlsx'])

if uploaded_file:
    with st.spinner('🚀 Processing...'):
        billing_df, ops_df, fname = process_data(uploaded_file)
        
        if billing_df is not None:
            st.success(f"✅ Success! File: **{fname}**")
            c1, c2 = st.columns(2)
            
            with c1:
                formatter = ExcelFormatter(billing_df)
                formatter.set_column_widths('BILLING')
                formatter.write_data('BILLING')
                st.download_button("📥 Billing Excel", data=formatter.get_file(), 
                                   file_name=f"BILLING_{fname}.xlsx", 
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                   use_container_width=True, type="primary")

            with c2:
                formatter = ExcelFormatter(ops_df)
                formatter.set_column_widths('OPS')
                formatter.write_data('OPS')
                st.download_button("📥 Ops Excel", data=formatter.get_file(), 
                                   file_name=f"OPS_{fname}.xlsx", 
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                   use_container_width=True, type="primary")
        else:
            st.error(f"❌ {fname}")
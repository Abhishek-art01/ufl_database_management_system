import os
import shutil
import pandas as pd
import psycopg2
import toml  # <--- Library to read your secrets.toml file

# --- CONFIGURATION ---
APP_FOLDER = "application_files"
MANUAL_FOLDER = "manual_files"
SECRETS_PATH = ".streamlit/secrets.toml"

def get_db_connection():
    try:
        # 1. Check if secrets file exists
        if not os.path.exists(SECRETS_PATH):
            print(f"❌ Error: Could not find secrets file at '{SECRETS_PATH}'")
            return None
            
        # 2. Load the secrets
        secrets = toml.load(SECRETS_PATH)
        db = secrets["postgres"] # Access the [postgres] section

        # 3. Connect using the details from the file
        return psycopg2.connect(
            host=db["host"],
            database=db["dbname"],
            user=db["user"],
            password=db["password"],
            port=db["port"],
            sslmode="require"
        )
    except Exception as e:
        print(f"❌ DB Connection Failed: {e}")
        return None

def process_folder(folder_path, table_name):
    # Ensure 'processed' folder exists
    processed_path = os.path.join(folder_path, "processed")
    if not os.path.exists(processed_path):
        os.makedirs(processed_path)

    # Get list of Excel files
    files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
    
    if not files:
        print(f"ℹ️ No new files found in {folder_path}")
        return

    conn = get_db_connection()
    if not conn: return
    cur = conn.cursor()

    print(f"📂 Processing {len(files)} files for table '{table_name}'...")

    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        print(f"   Reading: {file_name}...")
        
        try:
            df = pd.read_excel(file_path)
            # Clean column names
            df.columns = df.columns.str.strip()

            rows_inserted = 0
            for _, row in df.iterrows():
                sql = f"""
                    INSERT INTO {table_name} (
                        raw_date, trip_id, flight_no, employee_id, employee_name, gender, 
                        address, landmark, vehicle_no, direction, shift_time, trip_date, 
                        emp_count, pax_no, marshall, reporting_location, trip_zone
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                
                # Use .get() to handle missing columns safely
                values = (
                    row.get('DATE'), row.get('TRIP_ID'), row.get('FLIGHT_NO.'), 
                    row.get('EMPLOYEE_ID'), row.get('EMPLOYEE_NAME'), row.get('GENDER'),
                    row.get('ADDRESS'), row.get('LANDMARK'), row.get('VEHICLE_NO'),
                    row.get('DIRECTION'), str(row.get('SHIFT_TIME')), row.get('TRIP_DATE'),
                    row.get('EMP_COUNT'), row.get('PAX_NO'), row.get('MARSHALL'),
                    row.get('REPORTING_LOCATION'), row.get('TRIP_ZONE')
                )
                
                cur.execute(sql, values)
                rows_inserted += 1

            conn.commit()
            print(f"   ✅ Success! Inserted {rows_inserted} rows.")

            # MOVE file to processed folder
            shutil.move(file_path, os.path.join(processed_path, file_name))
            print(f"   📦 Moved {file_name} to 'processed' folder.")

        except Exception as e:
            print(f"   ❌ Error processing file {file_name}: {e}")
            if conn: conn.rollback()

    if cur: cur.close()
    if conn: conn.close()

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    print("🚀 Starting Data Import...")
    
    # 1. Process Application Files
    if os.path.exists(APP_FOLDER):
        process_folder(APP_FOLDER, "application_data_dump")
    else:
        print(f"⚠️ Folder '{APP_FOLDER}' not found. Please create it.")

    # 2. Process Manual Files
    if os.path.exists(MANUAL_FOLDER):
        process_folder(MANUAL_FOLDER, "manual_data_dump")
    else:
        print(f"⚠️ Folder '{MANUAL_FOLDER}' not found. Please create it.")
        
    print("\n✨ All tasks finished.")
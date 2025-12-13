import streamlit as st
import pandas as pd
import psycopg2
import toml
import time
import string
from datetime import datetime

# 1. PAGE CONFIG
st.set_page_config(page_title="Taxi Travel Management System", layout="wide", initial_sidebar_state="collapsed")

# CSS style
st.markdown("""
    <style>
        .block-container { padding-top: 3rem; padding-bottom: 1rem; }
        h1 { margin-top: -50px; }
        .stButton button { width: 100%; }
    </style>
""", unsafe_allow_html=True)

def get_db_connection():
    try:
        # st.secrets works on Cloud (reads dashboard settings) AND Local (reads the file automatically)
        db = st.secrets["postgres"] 
        
        return psycopg2.connect(
            host=db["host"],
            database=db["dbname"],
            user=db["user"],
            password=db["password"],
            port=db["port"]
        )
    except Exception as e:
        st.error(f"❌ Connection Failed: {e}")
        return None

def run_query(query, params=None, fetch=False):
    conn = get_db_connection()
    if conn is None: return None
    try:
        cur = conn.cursor()
        cur.execute(query, params)
        if fetch:
            columns = [desc[0] for desc in cur.description]
            data = cur.fetchall()
            conn.close()
            return pd.DataFrame(data, columns=columns)
        else:
            conn.commit()
            conn.close()
            return True
    except: 
        if conn: conn.close()
        return None

def get_next_voucher_number():
    """
    Finds the highest existing sequence number for TODAY
    and returns the next one.
    Format: YYYYMMDD-SEQ (e.g., 20251213-01, 20251213-02)
    """
    today = datetime.now().strftime("%Y%m%d") # Format: YYYYMMDD
    
    # Prefix is just the date plus a hyphen
    search_prefix = f"{today}-"
    search_pattern = f"{search_prefix}%"
    
    # Logic: Look for vouchers starting with today's date.
    # We take the substring AFTER the hyphen and convert to INT to find the max.
    sql = """
    SELECT SUBSTRING(voucher_no FROM LENGTH(%s) + 1)::INT
    FROM taxi_travels
    WHERE voucher_no LIKE %s
    ORDER BY 1 DESC
    LIMIT 1
    """
    params = (search_prefix, search_pattern)
    df = run_query(sql, params, fetch=True)
    
    if df is not None and not df.empty and pd.notnull(df.iloc[0, 0]):
        # Found existing, increment the highest sequence number
        last_seq = df.iloc[0, 0]
        next_seq = last_seq + 1
    else:
        # First voucher of the day
        next_seq = 1
        
    # Format: YYYYMMDD-01 (Using 2 digits for sequence)
    return f"{search_prefix}{next_seq:02d}"

# --- STATE ---
if "found_employees" not in st.session_state: st.session_state["found_employees"] = []
if "search_done" not in st.session_state: st.session_state["search_done"] = False
if "view_data" not in st.session_state: st.session_state["view_data"] = None

# --- UI HEADER ---
st.markdown("#### 🚖 Taxi Travel Management System")

tab_entry, tab_view = st.tabs(["📝 Entry", "📊 Records"])

# ================= TAB 1: ENTRY =================
# ================= TAB 1: ENTRY =================
with tab_entry:
    
    # --- ROW 1: SEARCH BAR ---
    search_container = st.container()
    with search_container:
        col_type, col_rest = st.columns([1, 4])
        with col_type:
            s_type = st.selectbox("Type", ["Application", "Manual"], label_visibility="collapsed", key="search_type_selector")

        if s_type == "Manual":
            c2, c3, c4 = col_rest.columns([1.5, 1.5, 1])
            search_date = c2.date_input("Search Date", label_visibility="collapsed", key="s_date")
            search_trip_id = c3.text_input("Trip ID", placeholder="Trip ID", label_visibility="collapsed", key="s_id_man")
            do_search = c4.button("🔍 Find", use_container_width=True, key="btn_man")
        else:
            c2, c3 = col_rest.columns([3, 1])
            search_date = None
            search_trip_id = c2.text_input("Trip ID", placeholder="Trip ID (7 Digits)", label_visibility="collapsed", key="s_id_app")
            do_search = c3.button("🔍 Find", use_container_width=True, key="btn_app")

        if do_search:
            st.session_state["found_employees"] = []
            st.session_state["search_done"] = False
            valid = True
            
            if s_type == "Application":
                if not search_trip_id.isdigit() or len(search_trip_id) != 7:
                    st.error("Trip ID must be 7 digits."); valid = False
            if s_type == "Manual" and not search_trip_id:
                st.error("Trip ID is required."); valid = False
            
            if valid and search_trip_id:
                if s_type == "Application":
                    tbl = "application_data_dump"
                    sql = f"SELECT employee_id, employee_name, gender, address, direction, trip_date, shift_time FROM {tbl} WHERE CAST(trip_id AS TEXT) = %s"
                    params = (search_trip_id,)
                else:
                    tbl = "manual_data_dump"
                    sql = f"SELECT employee_id, employee_name, gender, address, direction, trip_date, shift_time FROM {tbl} WHERE CAST(trip_id AS TEXT) = %s AND trip_date = %s"
                    params = (search_trip_id, search_date)
                
                df = run_query(sql, params, fetch=True)
                if df is not None and not df.empty:
                    df.insert(0, "Select", False) 
                    st.session_state["found_employees"] = df.to_dict('records')
                    st.session_state["search_done"] = True
                else:
                    st.warning("No data found.")

    # --- PRE-FILL & EDITOR ---
    emp_options = st.session_state["found_employees"]
    pre_dir, pre_shift, pre_date = "Pick Up", "", datetime.today()
    edited_df = pd.DataFrame() 

    if st.session_state["search_done"] and emp_options:
        e = emp_options[0]
        pre_dir = e.get('direction', "Pick Up")
        pre_shift = e.get('shift_time', "")
        if e.get('trip_date'): pre_date = e.get('trip_date')
        
        st.caption("👇 Select employees for this trip:")
        edited_df = st.data_editor(
            pd.DataFrame(emp_options),
            column_config={
                "Select": st.column_config.CheckboxColumn("Add?", default=False, width="small"),
                "employee_name": st.column_config.TextColumn("Employee Name", width="medium", disabled=True),
                "address": st.column_config.TextColumn("Address", width="large", disabled=True),
                "employee_id": st.column_config.NumberColumn("ID", disabled=True, format="%d"),
                "gender": st.column_config.TextColumn("Gender", disabled=True),
            },
            hide_index=True, use_container_width=True
        )

    # --- SAVE FORM ---
    with st.form("entry_form"):
        st.markdown("---")
        is_locked = st.session_state["search_done"]
        
        # 1. Calculate the preview voucher (Visual only)
        # We assume the user wants the next available number for TODAY
        preview_voucher = get_next_voucher_number()

        c1, c2, c3, c4 = st.columns(4)
        disp_trip_id = search_trip_id if st.session_state.get("search_done") else ""
        
        f_trip = c1.text_input("Trip ID", value=disp_trip_id, disabled=is_locked)
        f_date = c2.date_input("Date", value=pre_date, disabled=is_locked)
        f_shift = c3.text_input("Shift", value=pre_shift, disabled=is_locked)
        d_opts = ["Pick Up", "Drop"]
        d_idx = d_opts.index(pre_dir) if pre_dir in d_opts else 0
        f_dir = c4.selectbox("Direction", d_opts, index=d_idx, disabled=is_locked)

        # 2. Display the Voucher in a Disabled Field
        c5, c6, c7 = st.columns([1, 1, 2])
        # This field is DISABLED so they can see it but not change it
        st.write("") # Spacer
        f_vouch_display = c5.text_input("Voucher No (Auto)", value=preview_voucher, disabled=True)
        f_amt = c6.number_input("Amount (₹)", min_value=0.0, step=10.0)
        f_reason = c7.text_input("Reason for Taxi")

        if st.form_submit_button("💾 Save & Generate Voucher", use_container_width=True, type="primary"):
            errs = []
            selected_rows = edited_df[edited_df["Select"] == True] if not edited_df.empty else pd.DataFrame()

            if selected_rows.empty: errs.append("⚠️ You must select at least one employee.")
            
            if errs: 
                for e in errs: st.error(e)
            else:
                conn = get_db_connection()
                cur = conn.cursor()
                try:
                    # RE-CALCULATE strictly at save time to prevent race conditions
                    full_base_vouch_no = get_next_voucher_number()
                    
                    suffixes = list(string.ascii_uppercase)
                    
                    for i, row in enumerate(selected_rows.itertuples()):
                        amt = f_amt if i == 0 else 0.0
                        
                        if i == 0:
                            final_vouch_no = full_base_vouch_no
                        else:
                            suffix_letter = suffixes[i - 1] 
                            final_vouch_no = f"{full_base_vouch_no}{suffix_letter}"

                        sql = """INSERT INTO taxi_travels (travel_date, travel_type, direction, shift_time, trip_id, sap_id, emp_name, address, reason, amount, voucher_no)
                                 VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
                        cur.execute(sql, (f_date, s_type, f_dir, f_shift, int(f_trip) if f_trip.isdigit() else 0, 
                                          int(row.employee_id), row.employee_name, row.address, f_reason.upper(), amt, final_vouch_no))
                    
                    conn.commit()
                    conn.close()
                    
                    st.success(f"✅ Saved! Voucher: {full_base_vouch_no}")
                    st.session_state["found_employees"] = []
                    st.session_state["search_done"] = False
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    if conn: conn.close()
                    st.error(f"Error: {e}")

# ================= TAB 2: VIEW RECORDS =================
with tab_view:
    
    # --- 1. SEARCH SECTION ---
    v_container = st.container()
    with v_container:
        vc1, vc2 = st.columns([1, 4])
        with vc1:
            v_type = st.selectbox("Search By", ["All Recent", "Manual Search"], label_visibility="collapsed", key="v_type_sel")

        if v_type == "Manual Search":
            vc2a, vc2b, vc2c = vc2.columns([1.5, 1.5, 1])
            v_date = vc2a.date_input("Date", label_visibility="collapsed", key="v_date")
            v_trip = vc2b.text_input("Trip ID", placeholder="Trip ID", label_visibility="collapsed", key="v_trip")
            v_btn = vc2c.button("🔍 Search", use_container_width=True, key="v_btn")
        else:
            vc2.info("Showing last 50 records. Switch to 'Manual Search' to find specific trips.")
            v_btn = False

    # --- 2. DATA FETCHING LOGIC ---
    if v_btn or (v_type == "All Recent" and st.session_state["view_data"] is None):
        if v_type == "Manual Search":
            if v_trip:
                sql = "SELECT * FROM taxi_travels WHERE CAST(trip_id AS TEXT) = %s"
                params = [v_trip]
                if v_date:
                    sql += " AND travel_date = %s"
                    params.append(v_date)
                sql += " ORDER BY s_no ASC"
                st.session_state["view_data"] = run_query(sql, tuple(params), fetch=True)
            else:
                st.warning("Please enter a Trip ID to search.")
        else:
            st.session_state["view_data"] = run_query("SELECT * FROM taxi_travels ORDER BY s_no DESC LIMIT 50", fetch=True)
            
    # --- 3. DATA DISPLAY ---
    if st.button("🔄 Refresh Table", key="refresh_view"): 
        st.session_state["view_data"] = run_query("SELECT * FROM taxi_travels ORDER BY s_no DESC LIMIT 50", fetch=True)
        st.rerun()

    df_view = st.session_state["view_data"]
    
    if df_view is not None and not df_view.empty:
        st.caption(f"Found {len(df_view)} records:")
        st.data_editor(
            df_view,
            column_config={
                "emp_name": st.column_config.TextColumn("Employee Name", width="medium"),
                "address": st.column_config.TextColumn("Address", width="large"),
                "reason": st.column_config.TextColumn("Reason", width="medium"),
                "travel_date": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
                "voucher_no": st.column_config.TextColumn("Voucher"),
                "amount": st.column_config.NumberColumn("Amount (₹)", format="%.2f"),
                "s_no": st.column_config.NumberColumn("Ref No", format="%d"),
            },
            disabled=True,
            hide_index=True,
            use_container_width=True,
            height=500
        )
    elif df_view is not None and df_view.empty:
        st.info("No records found matching your criteria.")
import streamlit as st
import pandas as pd
import json
import uuid
import datetime
import io
import os
import re
import time 

st.set_page_config(page_title="Item Invoicing System", layout="wide")

# --- LOCAL STORAGE SETUP ---
DATA_DIR = "app_data"
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

ORDERS_FILE = os.path.join(DATA_DIR, "active_orders.json")
CUST_FILE = os.path.join(DATA_DIR, "customers_master.xlsx")
ITEMS_FILE = os.path.join(DATA_DIR, "items_master.xlsx")
SLABS_FILE = os.path.join(DATA_DIR, "slabs_master.xlsx")
HISTORY_FILE = os.path.join(DATA_DIR, "history_master.xlsx")

# --- HELPER FUNCTIONS FOR PERSISTENCE ---
def load_orders():
    if os.path.exists(ORDERS_FILE):
        with open(ORDERS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_orders(orders):
    with open(ORDERS_FILE, 'w') as f:
        json.dump(orders, f, indent=4)

def load_excel_or_csv(filepath):
    if os.path.exists(filepath):
        try:
            return pd.read_excel(filepath, engine='openpyxl')
        except:
            return pd.read_csv(filepath)
    return pd.DataFrame()

# Initialize State from Local Files (Persistent)
if 'customers_df' not in st.session_state: st.session_state.customers_df = load_excel_or_csv(CUST_FILE)
if 'items_df' not in st.session_state: st.session_state.items_df = load_excel_or_csv(ITEMS_FILE)
if 'slabs_df' not in st.session_state: st.session_state.slabs_df = load_excel_or_csv(SLABS_FILE)
if 'history_df' not in st.session_state: st.session_state.history_df = load_excel_or_csv(HISTORY_FILE)
if 'placed_orders' not in st.session_state: st.session_state.placed_orders = load_orders()
if 'drafts' not in st.session_state: st.session_state.drafts = {}

# Initialize Inventory State (NON-PERSISTENT - Session Only)
if 'inventory_df' not in st.session_state: st.session_state.inventory_df = pd.DataFrame()

# State trackers for UI elements
if 'last_cust_id' not in st.session_state: st.session_state.last_cust_id = None
if 'last_items_id' not in st.session_state: st.session_state.last_items_id = None
if 'last_slabs_id' not in st.session_state: st.session_state.last_slabs_id = None
if 'last_history_id' not in st.session_state: st.session_state.last_history_id = None
if 'last_inv_id' not in st.session_state: st.session_state.last_inv_id = None

# --- ROBUST DATA PARSING HELPERS ---
def normalize_code(code):
    if isinstance(code, pd.Series): code = code.iloc[0] 
    return str(code).split('.')[0].strip()

def get_col_name(df, target_names):
    if isinstance(target_names, str): target_names = [target_names]
    for target in target_names:
        target_lower = target.strip().lower()
        for col in df.columns:
            if str(col).strip().lower() == target_lower: return col
        for col in df.columns:
            if target_lower in str(col).strip().lower(): return col
    return None

def get_single_col(df, col_name):
    col_data = df[col_name]
    if isinstance(col_data, pd.DataFrame):
        return col_data.iloc[:, 0] 
    return col_data

def get_row_val(row, target_names, default=None):
    if isinstance(target_names, str): target_names = [target_names]
    for target in target_names:
        target_lower = target.strip().lower()
        for col in row.index:
            if str(col).strip().lower() == target_lower: 
                val = row[col]
                return val.iloc[0] if isinstance(val, pd.Series) else val
        for col in row.index:
            if target_lower in str(col).strip().lower(): 
                val = row[col]
                return val.iloc[0] if isinstance(val, pd.Series) else val
    return default

# --- SLAB PARSER FUNCTION ---
def is_in_slab(slab_str, qty):
    slab_str = str(slab_str).lower().replace(" ", "")
    try:
        if 'to' in slab_str:
            parts = slab_str.split('to')
            return int(float(parts[0])) <= qty <= int(float(parts[1]))
        elif '-' in slab_str:
            parts = slab_str.split('-')
            return int(float(parts[0])) <= qty <= int(float(parts[1]))
        elif '>=' in slab_str: return qty >= int(float(re.sub(r"[^0-9\.]", "", slab_str)))
        elif '>' in slab_str: return qty >= int(float(re.sub(r"[^0-9\.]", "", slab_str)))
        elif '<=' in slab_str: return qty <= int(float(re.sub(r"[^0-9\.]", "", slab_str)))
        elif '<' in slab_str: return qty <= int(float(re.sub(r"[^0-9\.]", "", slab_str)))
    except Exception: pass
    return False

# --- PRICING FUNCTION ---
def calculate_price(item_row, quantity, slabs_df, apply_cd=True):
    product_code = normalize_code(get_row_val(item_row, 'Product Code', ''))
    
    dbp_per_ltr = float(get_row_val(item_row, 'DBP', 0)) if pd.notna(get_row_val(item_row, 'DBP', 0)) else 0.0
    pack_size = float(get_row_val(item_row, 'Pack Size', 1)) if pd.notna(get_row_val(item_row, 'Pack Size', 1)) else 1.0
    cd_percent = float(get_row_val(item_row, 'CD', 0)) if pd.notna(get_row_val(item_row, 'CD', 0)) else 0.0
    landing_price_unit = float(get_row_val(item_row, 'DLP per pack', 0)) if pd.notna(get_row_val(item_row, 'DLP per pack', 0)) else 0.0

    dbp_per_pack = dbp_per_ltr * pack_size
    is_slab_item = False
    applied_slab_str = "None"
    
    if not slabs_df.empty:
        pc_col = get_col_name(slabs_df, 'Product Code')
        qty_col = get_col_name(slabs_df, 'Qty discount Slabs')
        dlp_col = get_col_name(slabs_df, 'DLP for each slab')
        
        if pc_col and qty_col and dlp_col:
            slabs_df['Normalized_Code'] = get_single_col(slabs_df, pc_col).apply(normalize_code)
            prod_slabs = slabs_df[slabs_df['Normalized_Code'] == product_code]
            
            if not prod_slabs.empty:
                is_slab_item = True
                landing_price_unit = 0.0 
                applied_slab_str = "Qty out of Slab Bounds (Price = 0)" 
                
                for _, slab_row in prod_slabs.iterrows():
                    slab_str = str(get_row_val(slab_row, qty_col))
                    if is_in_slab(slab_str, quantity):
                        landing_price_unit = float(get_row_val(slab_row, dlp_col, 0))
                        applied_slab_str = slab_str
                        slab_pack_size_val = get_row_val(slab_row, 'Pack Size', None)
                        if slab_pack_size_val is not None and pd.notna(slab_pack_size_val):
                            pack_size = float(slab_pack_size_val)
                            dbp_per_pack = dbp_per_ltr * pack_size 
                        break

    if not apply_cd: cd_percent = 0.0  
        
    cash_discount_amount = dbp_per_pack * (cd_percent / 100)
    final_price_per_unit = landing_price_unit - cash_discount_amount
    total_price = final_price_per_unit * quantity
    
    return {
        "dbp_per_ltr": round(dbp_per_ltr, 2), "pack_size": pack_size, "dbp_per_pack": round(dbp_per_pack, 2),
        "landing_price_unit": round(landing_price_unit, 2), "cd_active": "Yes" if apply_cd else "No",
        "cash_discount_applied": round(cash_discount_amount, 2), "is_slab_item": "Yes" if is_slab_item else "No",
        "applied_slab": applied_slab_str, "final_price_per_unit": round(final_price_per_unit, 2), "total_price": round(total_price, 2)
    }

st.title("📦 Item Invoicing & Order Management System")
tab1, tab2, tab3 = st.tabs(["📁 Master Data Upload", "🛒 Create Order (Drafts)", "📋 Manage Active Orders"])

# --- TAB 1: MASTER DATA UPLOAD ---
with tab1:
    st.header("Upload Master Data")
    st.info("Files 1-4 are saved locally and persistent. File 5 (Inventory) is temporary for the current session.")
    
    row1_col1, row1_col2 = st.columns(2)
    with row1_col1:
        st.subheader("1. Customers")
        cust_file = st.file_uploader("Update Customers", type=['xlsx', 'csv'], key="cust")
        if cust_file and st.session_state.last_cust_id != cust_file.file_id:
            prog = st.progress(0, text="Uploading Customers...")
            for i in range(100): time.sleep(0.002); prog.progress(i + 1)
            with open(CUST_FILE, "wb") as f: f.write(cust_file.getbuffer())
            st.session_state.customers_df = load_excel_or_csv(CUST_FILE)
            st.session_state.last_cust_id = cust_file.file_id
            prog.empty()
            st.toast("Customers Master updated successfully!", icon="✅")
        if not st.session_state.customers_df.empty:
            st.success(f"✅ Loaded: {len(st.session_state.customers_df)} customers")
            st.dataframe(st.session_state.customers_df.head(3), use_container_width=True)

    with row1_col2:
        st.subheader("2. Main Items")
        items_file = st.file_uploader("Update Main Items", type=['xlsx', 'csv'], key="items")
        if items_file and st.session_state.last_items_id != items_file.file_id:
            prog = st.progress(0, text="Uploading Items...")
            for i in range(100): time.sleep(0.002); prog.progress(i + 1)
            with open(ITEMS_FILE, "wb") as f: f.write(items_file.getbuffer())
            st.session_state.items_df = load_excel_or_csv(ITEMS_FILE)
            st.session_state.last_items_id = items_file.file_id
            prog.empty()
            st.toast("Items Master updated successfully!", icon="✅")
        if not st.session_state.items_df.empty:
            st.success(f"✅ Loaded: {len(st.session_state.items_df)} items")
            st.dataframe(st.session_state.items_df.head(3), use_container_width=True)

    st.markdown("---")
    row2_col1, row2_col2 = st.columns(2)
    
    with row2_col1:
        st.subheader("3. Slabs Master")
        slabs_file = st.file_uploader("Update Slabs Master", type=['xlsx', 'csv'], key="slabs")
        if slabs_file and st.session_state.last_slabs_id != slabs_file.file_id:
            prog = st.progress(0, text="Uploading Slabs...")
            for i in range(100): time.sleep(0.002); prog.progress(i + 1)
            with open(SLABS_FILE, "wb") as f: f.write(slabs_file.getbuffer())
            st.session_state.slabs_df = load_excel_or_csv(SLABS_FILE)
            st.session_state.last_slabs_id = slabs_file.file_id
            prog.empty()
            st.toast("Slabs Master updated successfully!", icon="✅")
        if not st.session_state.slabs_df.empty:
            st.success(f"✅ Loaded: {len(st.session_state.slabs_df)} slab rules")
            st.dataframe(st.session_state.slabs_df.head(3), use_container_width=True)

    with row2_col2:
        st.subheader("4. Historical Data")
        history_file = st.file_uploader("Update Historical Data", type=['xlsx', 'csv'], key="hist")
        if history_file and st.session_state.last_history_id != history_file.file_id:
            prog = st.progress(0, text="Uploading Historical Data...")
            for i in range(100): time.sleep(0.005); prog.progress(i + 1)
            with open(HISTORY_FILE, "wb") as f: f.write(history_file.getbuffer())
            st.session_state.history_df = load_excel_or_csv(HISTORY_FILE)
            st.session_state.last_history_id = history_file.file_id
            prog.empty()
            st.toast("Historical Data updated successfully!", icon="✅")
        if not st.session_state.history_df.empty:
            st.success(f"✅ Loaded: {len(st.session_state.history_df)} historical records")
            st.dataframe(st.session_state.history_df.head(3), use_container_width=True)

    st.markdown("---")
    row3_col1, row3_col2 = st.columns(2)
    
    with row3_col1:
        st.subheader("5. Current Inventory (Optional)")
        st.caption("⚠️ **Session Only**: Needs to be re-uploaded if app is restarted. Required cols: Product Code, Stock")
        inv_file = st.file_uploader("Upload Inventory Data", type=['xlsx', 'csv'], key="inv")
        
        if inv_file and st.session_state.last_inv_id != inv_file.file_id:
            prog = st.progress(0, text="Loading Inventory to Memory...")
            for i in range(100): time.sleep(0.002); prog.progress(i + 1)
            # Read directly to session state without saving to drive
            if inv_file.name.endswith('.csv'):
                st.session_state.inventory_df = pd.read_csv(inv_file)
            else:
                st.session_state.inventory_df = pd.read_excel(inv_file, engine='openpyxl')
            st.session_state.last_inv_id = inv_file.file_id
            prog.empty()
            st.toast("Temporary Inventory Data Loaded!", icon="✅")
            
        if not st.session_state.inventory_df.empty:
            st.success(f"✅ Active in Session: {len(st.session_state.inventory_df)} stock records")
            st.dataframe(st.session_state.inventory_df.head(3), use_container_width=True)
            if st.button("Clear Inventory Data"):
                st.session_state.inventory_df = pd.DataFrame()
                st.session_state.last_inv_id = None
                st.rerun()

# --- TAB 2: CREATE ORDER (DRAFTS) ---
with tab2:
    st.header("Order Entry Form")
    
    if st.session_state.customers_df.empty or st.session_state.items_df.empty:
        st.warning("Please upload both Customers and Items master files to begin.")
    else:
        items_df = st.session_state.items_df
        cust_df = st.session_state.customers_df
        
        items_pc_col = get_col_name(items_df, ['Product Code'])
        items_pn_col = get_col_name(items_df, ['Product Name'])
        items_pg_col = get_col_name(items_df, ['Product Group', 'Group'])
        cust_pc_col = get_col_name(cust_df, ['Customer Code'])
        cust_pn_col = get_col_name(cust_df, ['Customer Name'])
        
        if items_pc_col and items_pn_col:
            items_df['Display Name'] = get_single_col(items_df, items_pc_col).apply(normalize_code) + " - " + get_single_col(items_df, items_pn_col).astype(str)
            product_list = items_df['Display Name'].tolist()
        else: product_list = []
            
        if cust_pc_col and cust_pn_col:
            cust_df['Display Name'] = get_single_col(cust_df, cust_pc_col).apply(normalize_code) + " - " + get_single_col(cust_df, cust_pn_col).astype(str)
            customer_list = cust_df['Display Name'].tolist()
        else: customer_list = []
            
        selected_customer = st.selectbox("Select Customer to start/resume draft", customer_list)
        
        # --- HISTORICAL DATA & ANALYTICS ---
        if selected_customer:
            cust_code = selected_customer.split(" - ")[0]
            
            with st.expander("📊 Customer Insights & Historical Data (Click to Expand)", expanded=True):
                hist_df = st.session_state.history_df.copy()
                hist_cust_col = get_col_name(hist_df, ['Customer Code'])
                hist_date_col = get_col_name(hist_df, ['Month Year', 'Month and Year', 'Date', 'Month_Year'])
                hist_mt_col = get_col_name(hist_df, ['Quantity Billed in MT', 'Qty in MT'])
                
                target_q_l = 0.0
                target_y_l = 0.0
                curr_y = 0
                curr_q = 0
                hist_pg_vols = {} 
                
                # --- CALCULATE CURRENT FY/QTR REGARDLESS OF HISTORY ---
                today = datetime.date.today()
                if today.month >= 4:
                    curr_y = today.year
                    curr_q = (today.month - 4) // 3 + 1
                else:
                    curr_y = today.year - 1
                    curr_q = 4
                
                if hist_cust_col and hist_date_col and hist_mt_col:
                    hist_df['Norm_Cust'] = get_single_col(hist_df, hist_cust_col).apply(normalize_code)
                    c_hist = hist_df[hist_df['Norm_Cust'] == cust_code].copy()
                    
                    if not c_hist.empty:
                        c_hist['Str_Date'] = get_single_col(c_hist, hist_date_col).astype(str).str.strip()
                        c_hist['ParsedDate'] = pd.to_datetime(c_hist['Str_Date'], format='%m.%Y', errors='coerce')
                        c_hist['ParsedDate'] = c_hist['ParsedDate'].fillna(pd.to_datetime(c_hist['Str_Date'], errors='coerce'))
                        c_hist = c_hist.dropna(subset=['ParsedDate'])
                        
                        if not c_hist.empty:
                            c_hist['Month'] = c_hist['ParsedDate'].dt.month
                            c_hist['CalYear'] = c_hist['ParsedDate'].dt.year
                            c_hist['Year'] = c_hist.apply(lambda x: x['CalYear'] if x['Month'] >= 4 else x['CalYear'] - 1, axis=1)
                            c_hist['Quarter'] = c_hist['Month'].apply(lambda m: (m - 4) // 3 + 1 if m >= 4 else 4)
                            c_hist['Quantity Billed in L'] = get_single_col(c_hist, hist_mt_col) * 1250 
                            
                            # 1. Raw Historical Table
                            q_hist = c_hist[(c_hist['Quarter'] == curr_q) & (c_hist['Year'].isin([curr_y-1, curr_y-2]))].copy()
                            if not q_hist.empty:
                                cols_to_show = [c for c in [hist_cust_col, 'Customer Name', hist_date_col, 'Product Group', 'Product Code', 'Product Name', hist_mt_col, 'Quantity Billed in L'] if c in q_hist.columns]
                                st.markdown(f"**Historical Sales (FY Q{curr_q} Data for FY {curr_y-1} & FY {curr_y-2})**")
                                st.dataframe(q_hist.sort_values(by='ParsedDate', ascending=False)[cols_to_show], hide_index=True)
                                
                                # 2. Aggregated Summary
                                agg_df = q_hist.groupby('Year')[[hist_mt_col, 'Quantity Billed in L']].sum().reset_index()
                                agg_df['Financial Year'] = agg_df['Year'].apply(lambda y: f"FY {int(y)}-{int(y)+1}")
                                cols_agg = ['Financial Year', hist_mt_col, 'Quantity Billed in L']
                                st.markdown(f"**Summary: Total FY Q{curr_q} Sales by Year**")
                                st.dataframe(agg_df.sort_values(by='Year', ascending=False)[cols_agg], hide_index=True)
                            else:
                                st.info(f"No historical data found for FY Q{curr_q} in the previous 2 years (FY {curr_y-1}, FY {curr_y-2}).")
                                
                            # 3. Target Calculations (Quarterly & Yearly)
                            prev_year_q_data = c_hist[(c_hist['Quarter'] == curr_q) & (c_hist['Year'] == curr_y - 1)]
                            prev_year_q_mt = get_single_col(prev_year_q_data, hist_mt_col).sum() if not prev_year_q_data.empty else 0
                            target_q_l = (prev_year_q_mt * 1250) * 1.10
                            
                            prev_year_data = c_hist[c_hist['Year'] == curr_y - 1]
                            prev_year_mt = get_single_col(prev_year_data, hist_mt_col).sum() if not prev_year_data.empty else 0
                            target_y_l = (prev_year_mt * 1250) * 1.10
                            
                            # 4. Extract Historical Product Groups
                            hist_pg_col = get_col_name(hist_df, ['Product Group', 'Group'])
                            if hist_pg_col and not prev_year_q_data.empty:
                                pg_grouped = prev_year_q_data.groupby(hist_pg_col)[hist_mt_col].sum()
                                for pg, mt_vol in pg_grouped.items():
                                    if pd.notna(pg):
                                        hist_pg_vols[str(pg).strip()] = mt_vol * 1250
                    else:
                        st.info("No historical records found for this specific customer.")
                else:
                    st.error("Historical Excel is missing essential columns. Make sure you have: 'Customer Code', 'Month Year', and 'Quantity Billed in MT'.")
                
                # --- Get Active Product Groups & Map Products for Tooltips ---
                active_pgs = []
                pg_to_products = {}
                
                if items_pg_col and items_pn_col:
                    try:
                        pg_series = get_single_col(items_df, items_pg_col)
                        pn_series = get_single_col(items_df, items_pn_col)
                        
                        for pg_val, pn_val in zip(pg_series, pn_series):
                            if pd.notna(pg_val) and pd.notna(pn_val):
                                pg_key = str(pg_val).strip().lower()
                                pn_str = str(pn_val).strip()
                                
                                if pg_key not in pg_to_products: pg_to_products[pg_key] = []
                                if pn_str not in pg_to_products[pg_key]: pg_to_products[pg_key].append(pn_str)
                                    
                        active_pgs = list(pg_to_products.keys())
                    except Exception: active_pgs = []

                # --- Order Analytics Parsing (JSON) ---
                active_qty_units_q, active_l_q = 0, 0.0
                completed_qty_units_q, completed_l_q = 0, 0.0
                active_qty_units_y, active_l_y = 0, 0.0
                completed_qty_units_y, completed_l_y = 0, 0.0
                curr_pg_vols = {} 
                
                for o in st.session_state.placed_orders:
                    if cust_code == o['order_id'].split("_")[0]: 
                        o_date = pd.to_datetime(o['timestamp'], errors='coerce')
                        if pd.isna(o_date): o_fy, o_q = curr_y, curr_q
                        else:
                            o_m, o_y_cal = o_date.month, o_date.year
                            o_fy = o_y_cal if o_m >= 4 else o_y_cal - 1
                            o_q = (o_m - 4) // 3 + 1 if o_m >= 4 else 4
                        
                        is_curr_fy = (o_fy == curr_y)
                        is_curr_q = is_curr_fy and (o_q == curr_q)
                        
                        if is_curr_fy:
                            for item in o['items']:
                                vol_l = item.get('Qty', 0) * item.get('Pack Size', 1)
                                
                                pg = item.get('Product Group')
                                if not pg and items_pc_col:
                                    try:
                                        p_code = normalize_code(item.get('Product Code', ''))
                                        items_df['Temp_Code'] = get_single_col(items_df, items_pc_col).apply(normalize_code)
                                        f_row = items_df[items_df['Temp_Code'] == p_code]
                                        pg = get_row_val(f_row.iloc[0], ['Product Group'], 'Unknown') if not f_row.empty else 'Unknown'
                                    except: pg = 'Unknown'
                                elif not pg: pg = 'Unknown'
                                
                                if is_curr_q: curr_pg_vols[pg] = curr_pg_vols.get(pg, 0) + vol_l

                                if o['status'] == 'Active':
                                    active_qty_units_y += item.get('Qty', 0)
                                    active_l_y += vol_l
                                    if is_curr_q:
                                        active_qty_units_q += item.get('Qty', 0)
                                        active_l_q += vol_l
                                elif o['status'] == 'Completed':
                                    completed_qty_units_y += item.get('Qty', 0)
                                    completed_l_y += vol_l
                                    if is_curr_q:
                                        completed_qty_units_q += item.get('Qty', 0)
                                        completed_l_q += vol_l
                
                # --- Dashboards Rendering ---
                st.markdown("### 🎯 Current Quarter Order Analytics")
                m1, m2, m3, m4 = st.columns(4)
                target_q_txt = f"Target = FY Q{curr_q} of FY {curr_y-1} + 10%" if curr_y > 0 else "Based on same quarter last year + 10%"
                
                diff_l_q = target_q_l - (active_l_q + completed_l_q)
                delta_q_str = f"-{diff_l_q:,.2f} L" if diff_l_q > 0 else f"+{abs(diff_l_q):,.2f} L (Target Exceeded!)"
                
                m1.metric("QTR Target (L)", f"{target_q_l:,.2f} L", help=target_q_txt)
                m2.metric("Active Orders (QTD)", f"{active_qty_units_q} units", f"{active_l_q:,.2f} L Vol", delta_color="off")
                m3.metric("Completed Orders (QTD)", f"{completed_qty_units_q} units", f"{completed_l_q:,.2f} L Vol", delta_color="off")
                m4.metric("Gap to QTR Target", f"{abs(diff_l_q):,.2f} L", delta=delta_q_str)

                st.markdown("### 📅 Yearly Order Analytics (YTD)")
                ym1, ym2, ym3, ym4 = st.columns(4)
                target_y_txt = f"Target = Total FY {curr_y-1} + 10%" if curr_y > 0 else "Based on last year + 10%"
                
                diff_l_y = target_y_l - (active_l_y + completed_l_y)
                delta_y_str = f"-{diff_l_y:,.2f} L" if diff_l_y > 0 else f"+{abs(diff_l_y):,.2f} L (Target Exceeded!)"
                
                ym1.metric("Yearly Target (L)", f"{target_y_l:,.2f} L", help=target_y_txt)
                ym2.metric("Active Orders (YTD)", f"{active_qty_units_y} units", f"{active_l_y:,.2f} L Vol", delta_color="off")
                ym3.metric("Completed Orders (YTD)", f"{completed_qty_units_y} units", f"{completed_l_y:,.2f} L Vol", delta_color="off")
                ym4.metric("Gap to Yearly Target", f"{abs(diff_l_y):,.2f} L", delta=delta_y_str)

                # --- Category Breakdown (DEF vs Lubes) ---
                st.markdown("### 🛢️ Product Category Breakdown (Current Qtr)")
                cat1, cat2 = st.columns(2)
                
                # Find all DEF volume (handling case variations)
                def_vol_q = sum(vol for pg, vol in curr_pg_vols.items() if str(pg).strip().upper() == 'DEF')
                total_vol_q = active_l_q + completed_l_q
                lubes_vol_q = total_vol_q - def_vol_q
                
                cat1.metric("DEF Orders (QTD)", f"{def_vol_q:,.2f} L")
                cat2.metric("Lubes Orders (QTD)", f"{lubes_vol_q:,.2f} L", help="All orders excluding DEF")

                # --- AI Suggestions Implementation ---
                st.markdown("### 💡 Product Group Order Suggestions (Current Quarter)")
                if hist_pg_vols:
                    suggestions = []
                    discontinued_vol = 0.0
                    active_top_pgs = []
                    
                    for pg, h_vol in hist_pg_vols.items():
                        c_vol = curr_pg_vols.get(pg, 0.0)
                        gap = h_vol - c_vol
                        
                        if gap > 0:
                            pg_key = str(pg).strip().lower()
                            is_active = pg_key in active_pgs
                            
                            if is_active:
                                active_top_pgs.append(pg)
                                prod_list = pg_to_products.get(pg_key, [])
                                prod_tooltip_list = "\n".join([f"• {p}" for p in prod_list]) if prod_list else "No active products"
                                
                                suggestions.append({
                                    "Product Group": pg, 
                                    "Historical Qtr Vol (L)": round(h_vol, 2), 
                                    "Current Qtr Vol (L)": round(c_vol, 2), 
                                    "Suggested Target (L)": round(gap, 2),
                                    "Status": "✅ Active",
                                    "Available Products (Hover)": prod_tooltip_list
                                })
                            else:
                                discontinued_vol += gap
                                suggestions.append({
                                    "Product Group": f"⚠️ {pg}", 
                                    "Historical Qtr Vol (L)": round(h_vol, 2), 
                                    "Current Qtr Vol (L)": round(c_vol, 2), 
                                    "Suggested Target (L)": round(gap, 2),
                                    "Status": "❌ Discontinued",
                                    "Available Products (Hover)": "N/A"
                                })
                    
                    if suggestions:
                        suggestions = sorted(suggestions, key=lambda x: x['Suggested Target (L)'], reverse=True)
                        st.dataframe(
                            pd.DataFrame(suggestions), 
                            use_container_width=True, 
                            hide_index=True,
                            column_config={
                                "Available Products (Hover)": st.column_config.TextColumn(
                                    "Available Products (Hover to View)",
                                    help="Hover over any cell in this column to see the full bulleted list of orderable products.",
                                    width="large"
                                )
                            }
                        )
                        if discontinued_vol > 0:
                            st.warning(f"⚠️ **Attention:** **{discontinued_vol:,.2f} L** of your historical target comes from **Discontinued** product groups. To ensure targets are met, reallocate this lost volume by upselling active groups (e.g., **{', '.join(active_top_pgs[:3]) if active_top_pgs else 'other available products'}**).")
                        else:
                            st.info(f"👉 **Tip:** The customer historically orders **{suggestions[0]['Product Group']}** in this quarter, but is currently lagging. Consider suggesting an order for this group to hit your targets!")
                    else:
                        st.success("✅ Excellent! The customer has met or exceeded historical volumes for all typical product groups this quarter.")
                else:
                    st.info("No product group historicals found for the previous year's quarter to base suggestions on.")

        # --- END HISTORICAL SECTION ---
        
        if selected_customer not in st.session_state.drafts: st.session_state.drafts[selected_customer] = []
        current_draft = st.session_state.drafts[selected_customer]
        st.markdown("---")
        
        col_input, col_draft = st.columns([1, 1.2], gap="large")
        
        with col_input:
            st.subheader("1. Add Line Item")
            selected_product = st.selectbox("Search Product", product_list)
            
            c_qty, c_toggle = st.columns(2)
            with c_qty: quantity = st.number_input("Quantity", min_value=1, step=1, value=1)
            with c_toggle: apply_cd = st.toggle("Apply Cash Discount (CD)", value=True)
            
            if selected_product:
                item_code = selected_product.split(" - ")[0]
                items_df['Normalized_Code'] = get_single_col(items_df, items_pc_col).apply(normalize_code)
                item_row = items_df[items_df['Normalized_Code'] == item_code].iloc[0]
                
                pricing = calculate_price(item_row, quantity, st.session_state.slabs_df, apply_cd)
                st.caption(f"Group: {get_row_val(item_row, ['Product Group'], 'N/A')} | Pack Size: **{pricing['pack_size']}** | DBP/Ltr: ₹{pricing['dbp_per_ltr']} | **DBP/Pack: ₹{pricing['dbp_per_pack']}**")
                
                # --- NEW INVENTORY STOCK CHECK LOGIC ---
                inv_warning = ""
                if not st.session_state.inventory_df.empty:
                    inv_df = st.session_state.inventory_df
                    inv_pc_col = get_col_name(inv_df, ['Product Code', 'Item Code', 'Material', 'Code'])
                    inv_stock_col = get_col_name(inv_df, ['Stock', 'Inventory', 'Qty', 'Quantity', 'Balance'])
                    
                    if inv_pc_col and inv_stock_col:
                        inv_df['Norm_Code'] = get_single_col(inv_df, inv_pc_col).apply(normalize_code)
                        stock_match = inv_df[inv_df['Norm_Code'] == item_code]
                        
                        if not stock_match.empty:
                            try:
                                avail_stock = float(get_single_col(stock_match, inv_stock_col).iloc[0])
                                if quantity > avail_stock:
                                    inv_warning = f"⚠️ **Low Stock Alert:** You are ordering **{quantity}**, but current inventory is only **{avail_stock}** units!"
                            except: pass 
                        else:
                            inv_warning = "⚠️ **Stock Alert:** This product is not found in your currently uploaded Inventory list."
                
                if inv_warning:
                    st.warning(inv_warning)
                # --- END INVENTORY CHECK ---
                
                if pricing['is_slab_item'] == 'Yes':
                    if "Price = 0" in pricing['applied_slab']: st.error(f"⚠️ Item in Slab Master, but Qty {quantity} has no rule. Landing Price set to ₹0.")
                    else: st.success(f"**Slab Rule Forced ({pricing['applied_slab']}):** Landing Price mapped to ₹{pricing['landing_price_unit']}")
                else: st.info(f"**Standard Landing Price (No Slabs Found):** ₹{pricing['landing_price_unit']}")
                
                st.write(f"CD Applied: ₹{pricing['cash_discount_applied']}" if apply_cd else "CD: ₹0 (Toggled Off)")
                st.markdown(f"### Final Unit Price: ₹{pricing['final_price_per_unit']}  x  {quantity} qty  =  **Total: ₹{pricing['total_price']}**")
                
                if st.button("➕ Add to Draft", type="primary", use_container_width=True):
                    line_item = {
                        "Product Code": item_code, "Product Name": get_row_val(item_row, ['Product Name'], 'Unknown'),
                        "Product Group": get_row_val(item_row, ['Product Group'], 'Unknown'),
                        "Pack Size": pricing['pack_size'], "Qty": quantity, "CD Active": pricing['cd_active'],
                        "Slab Applied": pricing['applied_slab'], "Unit Price (Final)": pricing['final_price_per_unit'],
                        "Total Price": pricing['total_price'], "_pricing_data": pricing
                    }
                    st.session_state.drafts[selected_customer].append(line_item)
                    st.toast(f"✅ {quantity}x {line_item['Product Name']} added to draft!", icon="🛒")
                    st.rerun()

        with col_draft:
            st.subheader("2. Current Order Draft")
            if not current_draft:
                st.info("No items added for this customer yet.")
            else:
                display_draft = [{k: v for k, v in item.items() if k != '_pricing_data'} for item in current_draft]
                st.dataframe(pd.DataFrame(display_draft), use_container_width=True, hide_index=True)
                
                total_order_qty = sum(item['Qty'] for item in current_draft)
                total_order_value = sum(item['Total Price'] for item in current_draft)
                st.markdown(f"**Total Items:** {total_order_qty} | **Grand Total: ₹{round(total_order_value, 2)}**")
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("🗑️ Clear Draft", use_container_width=True):
                        st.session_state.drafts[selected_customer] = []
                        st.rerun()
                with col_btn2:
                    if st.button("✅ Submit Order", type="primary", use_container_width=True):
                        cust_code = selected_customer.split(" - ")[0]
                        order_id = f"{cust_code}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        
                        final_order = {
                            "order_id": order_id, "customer": selected_customer, "items": current_draft,
                            "total_value": round(total_order_value, 2), "total_qty": total_order_qty,
                            "status": "Active", "timestamp": str(datetime.datetime.now())
                        }
                        
                        st.session_state.placed_orders.append(final_order)
                        save_orders(st.session_state.placed_orders)
                        st.session_state.drafts[selected_customer] = []
                        st.toast(f"🚀 Order {order_id} Submitted Successfully!", icon="✅")
                        st.rerun()

# --- TAB 3: MANAGE PLACED ORDERS & EXPORT ---
with tab3:
    col_dash_title, col_export = st.columns([2, 1])
    with col_dash_title: st.header("Active Orders Dashboard")
    with col_export:
        if st.session_state.placed_orders:
            export_data = []
            for order in st.session_state.placed_orders:
                c_code = order['customer'].split(" - ")[0]
                c_name = " - ".join(order['customer'].split(" - ")[1:])
                for item in order['items']:
                    export_data.append({
                        "Customer Code": c_code, "Customer Name": c_name, "Order ID": order['order_id'],
                        "Status": order['status'], "Product Code": item['Product Code'], "Product Name": item['Product Name'],
                        "Product Group": item.get('Product Group', 'Unknown'),
                        "Pack Size": item.get('Pack Size', 1), "Quantity": item['Qty'], "CD Active": item['CD Active'], 
                        "Slab Applied": item['Slab Applied'], "Unit Price": item['Unit Price (Final)'], 
                        "Total Line Price": item['Total Price'], "Timestamp": order['timestamp']
                    })
            
            df_export = pd.DataFrame(export_data)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                grouped = df_export.groupby("Customer Code")
                for cust_code, group in grouped:
                    safe_sheet_name = str(cust_code)[:31].replace("/", "_").replace("\\", "_")
                    group.drop(columns=["Customer Code"]).to_excel(writer, index=False, sheet_name=safe_sheet_name)
                df_export.to_excel(writer, index=False, sheet_name="All Active Orders")
            st.download_button("📥 Export Segregated Excel", data=buffer.getvalue(), file_name=f"Orders_Export_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            
    st.markdown("---")
    if not st.session_state.placed_orders: st.info("No orders have been submitted yet.")
    else:
        for i, order in enumerate(reversed(st.session_state.placed_orders)):
            actual_idx = len(st.session_state.placed_orders) - 1 - i 
            with st.expander(f"📦 {order['order_id']} | Total: ₹{order['total_value']} ({order['status']})", expanded=True):
                display_items = [{k: v for k, v in item.items() if k != '_pricing_data'} for item in order['items']]
                st.dataframe(pd.DataFrame(display_items), use_container_width=True, hide_index=True)
                
                c1, c2, c3 = st.columns(3)
                with c1:
                    if order['status'] != "Completed":
                        if st.button("Mark as Completed", key=f"comp_{order['order_id']}"):
                            st.session_state.placed_orders[actual_idx]['status'] = "Completed"
                            save_orders(st.session_state.placed_orders)
                            st.toast(f"🎉 Order {order['order_id']} marked as completed!", icon="🎉")
                            st.rerun()
                with c2:
                    if st.button("Delete Order", key=f"del_{order['order_id']}"):
                        st.session_state.placed_orders.pop(actual_idx)
                        save_orders(st.session_state.placed_orders)
                        st.toast(f"🗑️ Order deleted!", icon="🗑️")
                        st.rerun()
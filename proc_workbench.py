import streamlit as st 
import pandas as pd 
import numpy as np 
import io
import re

# ==========================================
# 1. PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Procurement Workbench",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #f9f9f9; }
    .stButton>button { width: 100%; background-color: #005eb8; color: white; }
    .metric-card { background-color: white; padding: 15px; border-radius: 10px; box-shadow: 2px 2px 10px rgba(0,0,0,0.1) }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================

def check_mandatory(value):
    if pd.isna(value): return False
    s_val = str(value).strip()
    if s_val == "" or s_val.lower() == "nan": return False
    return True

def check_must_be_empty(value):
    if pd.isna(value): return True
    s_val = str(value).strip()
    if s_val == "" or s_val.lower() == "nan": return True
    return False

def check_greater_than_zero(value):
    try:
        if float(value) > 0: 
            return True
    except:
        pass
    return False

def check_incoterm_rules(row):
    errors = []
    incot = str(row.get('IncoT', '')).strip().upper()
    inco2 = str(row.get('Inco. 2', '')).strip()

    if not incot or incot == 'NAN': return errors
    
    valid_versions = ["Inc2020", "Inc2010", "Inc2000"]
    if not any(v.upper() in inco2.upper() for v in valid_versions):
        errors.append(f"Location (Inco. 2) must specify version (e.g. Inc2020). Found: '{inco2}'")
    
    if incot == 'DAT': errors.append("Incoterm 'DAT' is obsolete. Change to 'DPU'.")

    ship_to_group = ['DAF', 'DDU', 'DEQ', 'DES', 'DAP', 'DDP', 'DPU']
    exw_group = ['EXW']
    fca_allowed = ["supplier warehouse", "specified warehouse", "securiforce wareh"]

    inco2_lower = inco2.lower()
    if incot in ship_to_group:
        if not inco2_lower.startswith("ship-to address"):
            errors.append(f"For {incot}, Location must start with 'Ship-to address...'")
    elif incot in exw_group:
        if not inco2_lower.startswith("supplier warehouse"):
            errors.append(f"For {incot}, Location must start with 'Supplier warehouse...'")
    elif incot == 'FCA':
        if not any(inco2_lower.startswith(x) for x in fca_allowed):
            errors.append(f"For FCA, Location invalid prefix")

    return errors

def check_postal_code(country, postal_code):
    if not check_mandatory(postal_code): return None
    postal_str = str(postal_code).strip()
    rules = {
        'GB': (5, 9), 
        'JP': (7, 8), 
        'PT': (7, 8), 
        'CA': (6, 7), 
        'AU': (4, 6), 
        'CN': (6, 6), 'IN': (6, 6), 'SG': (6, 6), 'TW': (3, 6),
        'FR': (5, 5), 'ID': (5, 5), 'MX': (5, 5), 'MY': (5, 5),
        'BE': (4, 4)
    }
    
    if country in rules:
        min_len, max_len = rules[country]
        curr_len = len(postal_str)
        
        if not (min_len <= curr_len <= max_len):
            if min_len == max_len:
                return f"Postal code for {country} must be {min_len} chars. Found: {curr_len}"
            else:
                return f"Postal code for {country} must be between {min_len}-{max_len} chars. Found: {curr_len}"
    return None

def get_duplicates_df(df):
    """Checks duplicates"""
    dupes = pd.DataFrame()
    
    # 1. Synertrade Logic (Complex)
    if 'Synertrade Supplier ID' in df.columns:
        df['__syn_clean'] = df['Synertrade Supplier ID'].astype(str).str.strip()
        mask = (df['__syn_clean'].ne('nan') & df['__syn_clean'].ne('') & df['__syn_clean'].ne('0'))
        valid_df = df[mask].copy()
        
        if 'Vendor' in df.columns:
            # Check if one SynID is used by multiple DIFFERENT Vendor IDs
            counts = valid_df.groupby('__syn_clean')['Vendor'].nunique()
            bad_syn_ids = counts[counts > 1].index.tolist()
            if bad_syn_ids:
                s_dupes = df[df['__syn_clean'].isin(bad_syn_ids)].copy()
                s_dupes['Duplicate_Reason'] = 'Synertrade ID used by multiple Vendors'
                dupes = pd.concat([dupes, s_dupes])

    # 2. Vendor ID Logic (Strict duplicates in file)
    if 'Vendor' in df.columns:
         df['__ven_clean'] = df['Vendor'].astype(str).str.strip()

    if '__syn_clean' in df.columns: del df['__syn_clean']
    if '__ven_clean' in df.columns: del df['__ven_clean']
    return dupes

def get_primary_id(row):
    if check_mandatory(row.get('Vendor')): return str(row['Vendor'])
    if check_mandatory(row.get('Synertrade Supplier ID')): return str(row['Synertrade Supplier ID'])
    return "N/A"

# ==========================================
# 3. SMD ANALYSIS LOGIC (Hybrid Engine)
# ==========================================

def to_excel_download_smd(full_df, df_errors, duplicates_df, metrics_dict, error_breakdown_df, bad_cells):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#005eb8', 'font_color': 'white', 'border': 1})
        bold_format = workbook.add_format({'bold': True})

        # --- Sheet 1: Dashboard ---
        worksheet0 = workbook.add_worksheet('Dashboard_Summary')
        worksheet0.write('B2', "High Level Summary", header_format)
        worksheet0.write('C2', "Count", header_format)
        worksheet0.write('B3', "Total Records", bold_format)
        worksheet0.write('C3', metrics_dict['Total'])
        worksheet0.write('B4', "Correct Records", bold_format)
        worksheet0.write('C4', metrics_dict['Correct'])
        worksheet0.write('B5', "Records with Errors", bold_format)
        worksheet0.write('C5', metrics_dict['Errors'])
        if 'Duplicates' in metrics_dict:
            worksheet0.write('B6', "Duplicates Found", bold_format)
            worksheet0.write('C6', metrics_dict['Duplicates'])

        worksheet0.write('B9', "Errors by Category", header_format)
        worksheet0.write('C9', "Count", header_format)
        categories = [
            ("Purchasing Info Issues", metrics_dict.get('Purchasing', 0)),
            ("Org Structure & Finance Issues", metrics_dict.get('Org_Finance', 0)),
            ("Master Data ID Issues", metrics_dict.get('Master_Data', 0)),
            ("General Issues", metrics_dict.get('General', 0))
        ]
        row_num = 9
        for cat, count in categories:
            worksheet0.write(row_num, 1, cat)
            worksheet0.write(row_num, 2, count)
            row_num += 1

        if not error_breakdown_df.empty:
            start_row = row_num + 3
            worksheet0.write(start_row, 1, "Top 5 Specific Errors", header_format)
            worksheet0.write(start_row, 2, "Count", header_format)
            top_5 = error_breakdown_df.head(5)
            for i, row in top_5.iterrows():
                worksheet0.write(start_row + i + 1, 1, row['Error Description'])
                worksheet0.write(start_row + i + 1, 2, row['Count'])
        worksheet0.set_column(1, 1, 60)

        # --- Sheet 2: Error Rows Summary ---
        if not df_errors.empty: 
            cols_to_keep = ['Row_Index', 'Primary_ID', 'Error_Details']
            if 'Name 1' in df_errors.columns: cols_to_keep.insert(2, 'Name 1')
            
            clean_view = df_errors[cols_to_keep].copy()
            clean_view.to_excel(writer, index=False, sheet_name='Error_Rows_Summary')
            ws1 = writer.sheets['Error_Rows_Summary']
            for i, col in enumerate(clean_view.columns): ws1.write(0, i, col, header_format)
            ws1.set_column(0, 3, 25)
            ws1.set_column(3, 3, 80)
            ws1.freeze_panes(1, 2)
        
        # --- Sheet 3: Full Raw Data ---
        cols_to_drop = ['Row_Index', 'Vendor_ID', 'Primary_ID', 'Name_1', 'Company_Code', 'Synertrade_ID', 'Error_Details']
        drop_list = [c for c in cols_to_drop if c in full_df.columns]
        raw_sheet_df = full_df.drop(columns=drop_list)

        raw_sheet_df.to_excel(writer, index=False, sheet_name='Full_Raw_Data')
        worksheet2 = writer.sheets['Full_Raw_Data']
        for col_num, value in enumerate(raw_sheet_df.columns.values): 
            worksheet2.write(0, col_num, value, header_format)

        col_map = {name: i for i, name in enumerate(raw_sheet_df.columns)}
        for row_idx, col_name in bad_cells:
            if col_name in col_map:
                excel_col_idx = col_map[col_name]
                excel_row_idx = row_idx + 1
                try:
                    cell_value = raw_sheet_df.iat[row_idx, excel_col_idx]
                    if pd.isna(cell_value):
                        worksheet2.write_blank(excel_row_idx, excel_col_idx, None, red_format)
                    else:
                        worksheet2.write(excel_row_idx, excel_col_idx, cell_value, red_format)
                except IndexError: pass
        worksheet2.freeze_panes(1, 4) 

        # --- Priority Tabs ---
        priority_config = [
            ('Purchasing_Issues', 'Purchasing_Validation'),
            ('Org_Finance_Issues', 'Org_Finance_Validation'),
            ('Master_Data_Issues', 'Master_Data_Validation')
        ]
        disp_cols = ['Row_Index', 'Primary_ID', 'Name_1', 'Company_Code']

        for col_name, sheet_name in priority_config:
            subset = full_df[full_df[col_name] != ""]
            if not subset.empty:
                valid_disp = [c for c in disp_cols if c in full_df.columns]
                final_cols = valid_disp + [col_name]
                subset[final_cols].to_excel(writer, index=False, sheet_name=sheet_name)

                # apply header format 
                ws = writer.sheets[sheet_name]
                for i, col in enumerate(final_cols):
                    ws.write(0, i, col, header_format)
                ws.set_column(0, len(final_cols)-1, 25) 

        # --- General Split Tabs ---
        gen_tabs = [
            ('missing|Required', 'Gen_Mandatory_Missing'),
            ('must be empty', 'Gen_Should_Be_Empty')
        ]

        if 'General_Errors' in full_df.columns:
            valid_disp = [c for c in disp_cols if c in full_df.columns]
            final_cols = valid_disp + ['General_Errors']

            # Mandatory & empty
            for keyword, sheet_name in gen_tabs:
                mask = full_df['General_Errors'].str.contains(keyword, case=False, na=False)
                sub_df = full_df[mask]
                if not sub_df.empty: 
                    sub_df[final_cols].to_excel(writer, index=False, sheet_name=sheet_name)
                    ws = writer.sheets[sheet_name]
                    for i, col in enumerate(final_cols): ws.write(0, i, col, header_format)
                    ws.set_column(0, len(final_cols)-1, 25)

            # Formatting             
            fmt_mask = (full_df['General_Errors'] != "") & \
            (~full_df['General_Errors'].str.contains("missing|Required", case=False, na=False)) & \
            (~full_df['General_Errors'].str.contains("must be empty", case=False, na=False))

            fmt_df = full_df[fmt_mask]
            if not fmt_df.empty:
                fmt_df[final_cols].to_excel(writer, index=False, sheet_name='Gen_Formatting_Logic')
                ws = writer.sheets['Gen_Formatting_Logic']
                for i, col in enumerate(final_cols): ws.write(0, i, col, header_format)
                ws.set_column(0, len(final_cols)-1, 25)

        # --- Duplicates ---
        if not duplicates_df.empty:
            priority_cols = ['Duplicate_Reason', 'Vendor', 'Name 1', 'Synertrade Supplier ID']
            existing_cols = duplicates_df.columns.tolist()
            new_order = [c for c in priority_cols if c in existing_cols] + [c for c in existing_cols if c not in priority_cols]
            duplicates_df = duplicates_df[new_order]

            duplicates_df.to_excel(writer, index=False, sheet_name='Potential_Duplicates')
            ws_dupe = writer.sheets['Potential_Duplicates']
            for i, col in enumerate(duplicates_df.columns): ws_dupe.write(0, i, col, header_format)
            ws_dupe.freeze_panes(1, 4)
            ws_dupe.set_column(0, 0, 30)

    return output.getvalue()

def run_smd_analysis(df, requirements_df, target_cocd, target_porg, region):
    df_out = df.copy()
    df.columns = df.columns.str.strip()

    # --- target porg (split by comma) ---
    # example: "0001, 0002" -> ['0001', '0002']
    valid_porg_list = []
    if target_porg: 
        valid_porg_list = [p.strip() for p in str(target_porg).split(',') if p.strip()]

    # --- 1. DYNAMIC CONFIGURATION (Load Rules from Excel) ---
    req_dict = {'Mandatory': [], 'Empty': []}
    
    if requirements_df is not None:
        # Standardize headers
        requirements_df.columns = [c.strip().title() for c in requirements_df.columns]
        
        # Check for required columns
        if 'Field' in requirements_df.columns and 'Rule' in requirements_df.columns:
            for idx, row in requirements_df.iterrows():
                field = str(row['Field']).strip()
                rule = str(row['Rule']).strip().lower()
                
                # Get Region & Category safely
                rule_region = str(row['Region']).strip().upper() if 'Region' in requirements_df.columns else 'ALL'
                cat = str(row['Category']).strip() if 'Category' in requirements_df.columns else 'General'
                
                # Apply if Region matches
                if rule_region == 'ALL' or rule_region == region.upper():
                    rule_obj = {'field': field, 'cat': cat}
                    
                    if 'mandatory' in rule: 
                        req_dict['Mandatory'].append(rule_obj) # Append Dictionary
                    if 'empty' in rule: 
                        req_dict['Empty'].append(rule_obj)     # Append Dictionary

    # --- 2. HARDCODED LOGIC & COLUMN MAPPING ---
    valid_vendor_ids = set()
    if 'Vendor' in df.columns: valid_vendor_ids = set(df['Vendor'].astype(str).str.strip())
    
    col_payt_fin = 'PayT'
    col_payt_purch = 'PayT.1' if 'PayT.1' in df.columns else 'PayT'
    if region == 'EU':
        col_payt_fin = 'PayT C.Co' if 'PayT C.Co' in df.columns else 'PayT'
        col_payt_purch = 'PayT POrg' if 'PayT POrg' in df.columns else 'PayT'

    all_error_details, p1_list_col, p2_list_col, p3_list_col, gen_list_col, bad_cells = [], [], [], [], [], []

    def log_err(category_list, msg, col_name=None, idx=None):
        category_list.append(msg)
        if col_name and col_name in df.columns: bad_cells.append((idx, col_name))

    # --- 3. EXECUTION ---
    for index, row in df.iterrows():
        p1, p2, p3, gen = [], [], [], [] 
        country = str(row.get('Cty','')).strip().upper()
        is_local = (country == 'MY')

        # --- A. Apply Dynamic Rules (Excel) ---
        def get_target_list(cat_name):
            c = cat_name.lower()
            if 'purchasing' in c: return p1
            if 'org' in c or 'finance' in c: return p2
            if 'master' in c or 'vendor' in c: return p3
            return gen # default
        
        for item in req_dict['Mandatory']:
            target_list = get_target_list(item['cat'])
            if item['field'] in df.columns and not check_mandatory(row[item['field']]):
                log_err(target_list, f"{item['field']} is missing", item['field'], index)
        for item in req_dict['Empty']:
            if item['field'] in df.columns and not check_must_be_empty(row[item['field']]):
                log_err(target_list, f"{item['field']} must be empty", item['field'], index)

        # --- B. Hardcoded Business Logic (Things too complex for simple Excel rules) ---

        # 1. Currency (Check all variations)
        crcy_cols = [c for c in df.columns if 'crcy' in c.lower() or 'currency' in c.lower()]
        if crcy_cols:
            has_currency = any(check_mandatory(row[c]) for c in crcy_cols)
            if not has_currency:
                log_err(p1, "Currency missing (All Cols)", crcy_cols[0], index)

        if col_payt_purch in df.columns and not check_mandatory(row[col_payt_purch]): log_err(p1, "Purch PayT Missing", col_payt_purch, index)
        
        # 2. Incoterms
        if 'IncoT' in df.columns:
            if not check_mandatory(row['IncoT']): log_err(p1, "Incoterms missing", 'IncoT', index)
            elif 'Inco. 2' in df.columns:
                if not check_mandatory(row['Inco. 2']): log_err(p1, "Inco. 2 missing", 'Inco. 2', index)
                else: p1.extend(check_incoterm_rules(row))

        # 3. Payment Terms
        if col_payt_fin in df.columns and col_payt_purch in df.columns:
            fin, pur = str(row[col_payt_fin]).strip(), str(row[col_payt_purch]).strip()
            if fin != pur: log_err(p2, f"PayT Mismatch ({fin} vs {pur})", col_payt_fin, index)
        
        # 4. Org Check
        if 'CoCd' in df.columns and str(row['CoCd']).strip() != target_cocd: log_err(p2, f"CoCd != {target_cocd}", 'CoCd', index)
        if 'POrg' in df.columns:
            porg = str(row['POrg']).strip()
            if valid_porg_list:
                if check_mandatory(porg):
                    if porg not in valid_porg_list: 
                        log_err(p2, f"POrg != {valid_porg_list}", 'POrg', index)
                else:
                    log_err(p2, "POrg is missing", 'POrg', index)

        # 5. Tax Logic (At least 1 required)
        tax_candidates = [c for c in df.columns if ('tax' in c.lower() or 'vat' in c.lower()) 
                          and 'identification' not in c.lower() and 'liable' not in c.lower() and 'equal' not in c.lower()]
        if tax_candidates:
            has_tax = any(check_mandatory(row[c]) for c in tax_candidates)
            if not has_tax: 
                gen.append("At least one Tax ID Required")
                bad_cells.append((index, tax_candidates[0]))

        # 6. Postal Code
        if 'Postl Code' in df.columns:
            postal_err = check_postal_code(country, row['Postl Code'])
            if postal_err: log_err(gen, postal_err, 'Postl Code', index)

        # 7. Duplicate Logic (Alt Payee Scope)
        if 'AltPayeeAc' in df.columns:
            alt = str(row['AltPayeeAc']).strip()
            if check_mandatory(alt) and alt not in valid_vendor_ids:
                log_err(p3, "AltPayee Not in Scope", 'AltPayeeAc', index)
        
        # 8. Telephone 1 logic
        if 'Telephone 1' in df.columns:
            phone = str(row['Telephone 1']).strip()

            # Rule 1: Mandatory for Global
            if not check_mandatory(phone): 
                log_err(gen, "Tel 1 missing", 'Telephone 1', index)
            
            # Rule 2: "+" symbol only for CoCd 3072
            elif str(target_cocd).strip() == '3072' and "+" not in phone: 
                log_err(gen, "Tel 1 missing '+'", 'Telephone 1', index)
                    
        # 9. Synertrade Supplier ID
        if 'Synertrade Supplier ID' in df.columns:
                syn_id = str(row['Synertrade Supplier ID'])
                if not check_mandatory(syn_id):
                    log_err(p3, "Synertrade ID missing", 'Synertrade Supplier ID', index)

        # Consolidate
        all_errs = p1 + p2 + p3 + gen
        all_error_details.append(" | ".join(all_errs))
        p1_list_col.append(" | ".join(p1))
        p2_list_col.append(" | ".join(p2))
        p3_list_col.append(" | ".join(p3))
        gen_list_col.append(" | ".join(gen))

    df_out.insert(0, 'General_Errors', gen_list_col)
    df_out.insert(0, 'Master_Data_Issues', p3_list_col)
    df_out.insert(0, 'Org_Finance_Issues', p2_list_col)
    df_out.insert(0, 'Purchasing_Issues', p1_list_col)
    df_out.insert(0, 'Error_Details', all_error_details)
    
    df_out['Row_Index'] = df_out.index + 2
    df_out['Primary_ID'] = df_out.apply(get_primary_id, axis=1)
    df_out['Vendor_ID'] = df_out.get('Vendor', 'N/A')
    df_out['Name_1'] = df_out.get('Name 1', 'N/A')
    df_out['Company_Code'] = df_out.get('CoCd', 'N/A')
    
    return df_out, bad_cells

# ==========================================
# 4. EMAIL ANALYSIS & MAIN UI
# ==========================================

def run_email_analysis(df):
    df.columns = df.columns.str.strip()
    req_cols = ['Vendor', 'Communication link notes', 'ID', '#']
    missing = [c for c in req_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns: {missing}")
        return pd.DataFrame()
    df_out = df.copy()

    # convert "ID" to numbers, turn errors into NaN, fill NaN with a huge number so they are not "min"
    df_out['ID Numeric'] = pd.to_numeric(df_out['ID'], errors='coerce').fillna(9999999)

    vendor_groups = df_out.groupby('Vendor')
    email_errors = []

    for idx, row in df_out.iterrows():
        errs = []
        vendor = row['Vendor']
        vendor_data = vendor_groups.get_group(vendor)

        # Notes check
        notes = str(row['Communication link notes']).strip()
        if not check_mandatory(notes): errs.append("Comm. Notes Empty")

        # Smallest ID logic 
        # find the mathematical minimum ID
        min_id = vendor_data['ID Numeric'].min()
        current_id = row['ID Numeric']

        # check if row is marked (X or 1)
        flag = str(row['#']).strip().upper()
        is_marked = (flag == 'X' or flag == '1')

        # Scenario: if it is the smallest ID, but not marked
        if current_id == min_id:
            if not is_marked: 
                errs.append("Smallest ID not marked Default (#)")

        elif current_id != min_id:
            if is_marked: 
                errs.append("Non-smallest ID marked as Default")
                
        email_errors.append(" | ".join(errs))

    final_error_col = []
    for idx, row in df_out.iterrows():
        final_error_col.append(email_errors[idx])

    df_out.insert(0, 'Email_Validation_Errors', final_error_col)

    # cleanup temp column
    if 'ID Numeric' in df_out.columns: del df_out['ID Numeric']

    return df_out

def to_excel_email_download(df_result, metrics_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#005eb8', 'font_color': 'white', 'border': 1})
        bold_format = workbook.add_format({'bold': True})
        
        # Dashboard summary 
        ws0 = workbook.add_worksheet('Dashboard Summary')

        # High level metrics
        ws0.write('B2', "Email Validation Summary", header_format)
        ws0.write('C2', "Count", header_format)
        ws0.write('B3', "Total Records", bold_format)
        ws0.write('C3', metrics_dict['Total'])
        ws0.write('B4', "Correct Records", bold_format)
        ws0.write('C4', metrics_dict['Correct'])
        ws0.write('B5', "Records with Issues", bold_format)
        ws0.write('C5', metrics_dict['Errors'])

        ws0.set_column(1, 2, 25)

        # Error Summary
        df_err = df_result[df_result['Email_Validation_Errors'] != ""]
        if not df_err.empty:
            df_err.to_excel(writer, index=False, sheet_name='Error_Summary')
            ws1 = writer.sheets['Error_Summary']
            for i, col in enumerate(df_err.columns): ws1.write(0, i, col, header_format)
            ws1.set_column(0, 0, 50)
            ws1.set_column(1, len(df_err.columns)-1, 15)
        
        # Full email data
        df_result.to_excel(writer, index=False, sheet_name='Full_Email_Data')
        ws2 = writer.sheets['Full_Email_Data']
        (max_row, max_col) = df_result.shape

        for i, col in enumerate(df_result.columns): ws2.write(0, i, col, header_format)

        # conditional formatting applied
        ws2.conditional_format(1, 0, max_row, max_col - 1, {'type': 'formula', 'criteria': '=$A2<>""', 'format': red_format})
        ws2.set_column(0, 0, 50)

    return output.getvalue()

# ==========================================
# 5. PO ANALYSIS LOGIC 
# ==========================================

def load_po_config(config_file):
    """Reads the Excel config and returns dictionaries/sets of rules"""
    config = {}
    try: 
        xls = pd.ExcelFile(config_file)

        # Load logic matrix 
        if 'PO_Logic_Matrix' in xls.sheet_names:
            df_matrix = pd.read_excel(xls, 'PO_Logic_Matrix').fillna('')
            config['matrix'] = df_matrix.to_dict('records')

        # Load settings (single values)
        if 'Settings' in xls.sheet_names: 
            df_set = pd.read_excel(xls, 'Settings')
            # convert two columns into a dictionary: {parameter: value}
            config['settings'] = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))

            # Load Doc Type Rules
            if 'Doc_Type_Rules' in xls.sheet_names:
                df_doc = pd.read_excel(xls, 'Doc_Type_Rules')
                # create list of types that require material 
                config['req_material'] = set(df_doc[df_doc['Requires_Material'] == 'Yes'].iloc[:,0].astype(str).str.strip())
                # create list of types that must not have material 
                config['no_material'] = set(df_doc[df_doc['Requires_Material'] == 'No'].iloc[:,0].astype(str).str.strip())

            # Standard Rules 
            if 'Standard_Rules' in xls.sheet_names:
                df_std = pd.read_excel(xls, 'Standard_Rules')
                config['standard'] = df_std.to_dict('records')
            
            # Reference lists 
            # Load lists
            list_map = {
                'PCN List': 'valid_pcn',
                'UNSPSC List': 'valid_unspsc',
                'UOM Master': 'valid_uom',
                'Active List': 'active_vendors', 
                'Suppress Vendors': 'suppress_vendors',
                'Payment Terms': 'valid_payt'
            }
            for sheet, key in list_map.items(): 
                if sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet)
                    config[key] = set(df.iloc[:, 0].dropna().astype(str).str.strip())
                
    except Exception as e: 
        st.error(f"Error reading PO config: {e}")
        
    return config

def check_special_characters(text, banned_chars_list):
    if pd.isna(text):
        return None

    text = str(text)

    for char in banned_chars_list:
        if char == '&':
            continue # skip '&' here to run the custom logic below
        if char in text: 
            return f"Contains banned character: '{char}'"
    if '</' in text or '/>' in text: 
        return "Invalid XML tags"
    if '&' in text: 
        # Rule: Only flag if the text is trapped between 2 &s without spaces (e.g. &text&)
        # Regex explanation: 
        # &     : Literal &
        # [^ ]+ : One or more characters that have no spaces
        # &     : Literal &
        if re.search(r'&[^ ]+&', text): 
            return "Invalid format '&text&'"
    if '•' in text:
        return "Contains Bullet point"
    if '→' in text: 
        return "Contains Arrow symbol"
    return None 

def check_intercompany_vendor(vendor_id):
    if pd.isna(vendor_id): 
        return False
    return bool(re.match(r'^A\d{4}', str(vendor_id)))

def run_po_analysis_dynamic(df, config_file):
    # --- 1. SETTINGS & RULES ---
    rules = load_po_config(config_file)
    settings = rules.get('settings', {})
    matrix = rules.get('matrix', [])
    
    # Parse Settings
    no_dec = set(str(settings.get('No_Decimal_Currencies', 'JPY,KRW,IDR')).split(','))
    special_chars = str(settings.get('Banned_Chars', '</,/>,&')).split(',')
    banned_reqs = set([x.strip() for x in str(settings.get('Banned_Requestors', '')).split(',') if x.strip()])
    
    try: small_val_limit = float(settings.get('Small Value Limit', 10.0))
    except: small_val_limit = 10.0

    max_short = int(float(settings.get('Max_Short_Text_Length', 40)))
    max_req = int(float(settings.get('Max_Requestor_Length', 12)))
    max_prep = int(float(settings.get('Max_Preparer_Length', 12)))
    max_unload = int(float(settings.get('Max_Unloading_Pt_Length', 25)))

    # --- 2. INTELLIGENT COLUMN MAPPING ---
    def get_actual_col(potentials):
        for p in potentials:
            if p in df.columns: return p
        clean_map = {c.lower().replace(' ', '').replace('_', '').replace('.', ''): c for c in df.columns}
        for p in potentials:
            cp = p.lower().replace(' ', '').replace('_', '').replace('.', '')
            if cp in clean_map: return clean_map[cp]
        return None

    # Map every column required by your specific checks
    col_map = {
        'gr': get_actual_col(['GR', 'G/R', 'Goods Receipt']),
        'material': get_actual_col(['Material', 'Material Number', 'Mat. No.']),
        'type': get_actual_col(['Type', 'Doc Type']),
        'net_price': get_actual_col(['Net Price_Ori', 'Net price_ori', 'Net Price']),
        'still_del': get_actual_col(['Still to Del Quantity', 'Still to Deliver']),
        'still_pay_qty': get_actual_col(['Still to Pay Quantity', 'Still to Pay']),
        'still_pay_amt': get_actual_col(['Still to pay amt ori', 'Still to Pay Amount']),
        'ir_exist': get_actual_col(['IR_Exist', 'IR Indicator']),
        'd_item': get_actual_col(['D-Item', 'Deletion Indicator']),
        'incomplete': get_actual_col(['Incomplete']),
        'rel_ind': get_actual_col(['Rel', 'Release Indicator']),
        'dci': get_actual_col(['DCI', 'Delivery Complete']),
        'fin': get_actual_col(['FIN', 'Final Invoice']),
        'rebate': get_actual_col(['R', 'Rebate', 'Return Item']),
        'matl_group': get_actual_col(['Matl Group', 'Material Group']),
        'vendor': get_actual_col(['Vendor', 'Supplier']),
        'oun': get_actual_col(['OUn', 'Order Unit']),
        'po_uom_ext': get_actual_col(['PO UOM - Ext']),
        'u_price_ord': get_actual_col(['Order Price Unit (Purchasing)']),
        'requestor': get_actual_col(['Requestor']),
        'preparer': get_actual_col(['Preparer']),
        'saa': get_actual_col(['SAA', 'Split']),
        'curr_net': get_actual_col(['Curr.', 'Curency']),
        'curr_unit': get_actual_col(['Crcy']),
        'unit_price': get_actual_col(['Unit Price']),
        'schd': get_actual_col(['Schd.', 'Schedule Line']),
        'slm': get_actual_col(['Supplier SLMID', 'SLM ID']),
        'unload': get_actual_col(['Unloading Point - Ext']),
        'payt': get_actual_col(['PayT', 'Payment Term']),
        'amt_eur': get_actual_col(['Still to pay amt eur', 'Still to Pay Amt Eur']),
        'incot': get_actual_col(['IncoT', 'Incoterm']),
        'per': get_actual_col(['Per']),
        'short_text': get_actual_col(['Short Text']),
        'ven_mat': get_actual_col(['Vendor Material Number'])
    }

    # --- 3. DATA PREPARATION ---
    records = df.to_dict('records')
    category_list = ['PCN', 'Unit of Measurement', 'Requestor', 'Preparer', 'Split Accounting', 
                     'Text', 'Currency', 'Schedule Line', 'Vendor', 'Unloading Point', 
                     'Doc Type', 'Payment Term', 'FOC', 'Logic Checks', 'Additional Pricing', 'Incoterm']
    
    cat_remarks_collector = {cat: [] for cat in category_list}
    po_categories, po_statuses, po_remarks, error_details = [], [], [], []
    bad_cells = []

    def safe_f(val):
        if val is None or val == '' or val != val: return 0.0
        try: return float(str(val).replace(',', ''))
        except: return 0.0

    # --- 4. THE OPTIMIZED LOOP ---
    for idx, row in enumerate(records):
        row_cat_errors = {cat: [] for cat in category_list}
        
        def get_v(key):
            actual = col_map.get(key)
            return row.get(actual) if actual else ""

        def log_e(cat, msg, key):
            row_cat_errors[cat].append(msg)
            actual = col_map.get(key)
            if actual: bad_cells.append((idx, actual))

        # Core logic variables
        gr_val = str(get_v('gr')).strip().upper()
        mat_raw = str(get_v('material')).strip()
        mat_val = "" if mat_raw in ['0', '00000000', 'nan', ''] else mat_raw
        net_price = safe_f(get_v('net_price'))
        still_del = safe_f(get_v('still_del'))
        still_pay_qty = safe_f(get_v('still_pay_qty'))
        still_pay_amt = safe_f(get_v('still_pay_amt'))
        ir_exist_raw = str(get_v('ir_exist')).strip().upper()
        ir_clean = ir_exist_raw.replace('.', '').replace(' ', '')

        # PO Category Determination
        p_cat = "Direct PO" if mat_val != "" else ("Material PO" if gr_val == 'X' else "Service PO")
        
        # Matrix Logic (Simplified evaluation)
        p_stat, p_rem, rule_found = "Review", "No matching logic found", False
        for rule in matrix:
            match = True
            if str(rule.get('Category', '')).strip() and str(rule.get('Category', '')).strip() != p_cat: match = False
            if match:
                r_gr = str(rule.get('GR', '')).strip().upper()
                if r_gr == 'X' and gr_val != 'X': match = False
                elif r_gr == 'EMPTY' and gr_val != '': match = False
            if match:
                r_mat = str(rule.get('Material', '')).strip().upper()
                if r_mat == 'FILLED' and mat_val == '': match = False
                elif r_mat == 'EMPTY' and mat_val != '': match = False
            # Condition parsing
            if match:
                r_cond = str(rule.get('Conditions', '')).upper().replace(' ', '')
                if r_cond:
                    if 'PRICE=0' in r_cond and net_price != 0: match = False
                    if 'PRICE>0' in r_cond and net_price <= 0: match = False
                    if 'DEL=0' in r_cond and still_del != 0: match = False
                    if 'DEL>0' in r_cond and still_del <= 0: match = False
                    if 'PAYQTY=0' in r_cond and still_pay_qty != 0: match = False
                    if 'PAYQTY>0' in r_cond and still_pay_qty <= 0: match = False
                    if 'PAYAMT=0' in r_cond and still_pay_amt != 0: match = False
                    if 'IR_EXIST=' in r_cond:
                        target = r_cond.split('IR_EXIST=')[1].split(',')[0].replace('.', '').replace(' ', '')
                        if ir_clean != target: match = False
            if match:
                p_stat, p_rem, rule_found = rule.get('Status'), rule.get('Remark'), True
                break

        # Hardcoded Filter Rule Overrides
        if str(get_v('d_item')) == 'L': p_stat, p_rem = 'Closed', 'Deleted Item. PO line item closed.'
        elif str(get_v('incomplete')) == 'X': p_stat, p_rem = 'Closed', 'Incomplete/on hold item.'
        elif str(get_v('rel_ind')) in ['Z', 'P']: p_stat, p_rem = 'Closed', 'Blocked item.'
        elif get_v('dci') == 'X' or get_v('fin') == 'X': p_stat, p_rem = 'Closed', 'PO closed due to DCI/FIN ticked.'
        elif get_v('rebate') == 'X': p_stat, p_rem = 'Closed', 'Rebate or Return Item.'

        # --- RESTORED SPECIFIC CATEGORY CHECKS ---
        if p_stat in ['Open', 'Check with Local']:
            # 1. PCN
            m_group = str(get_v('matl_group')).strip()
            vnd = str(get_v('vendor')).strip()
            if check_intercompany_vendor(vnd) and m_group != 'I9999': log_e('PCN', "Intercompany PO not using PCN I9999", 'matl_group')
            if 'valid_pcn' in rules and m_group not in rules['valid_pcn']: log_e('PCN', "PCN not in PCN tool", 'matl_group')
            if 'valid_unspsc' in rules and m_group not in rules['valid_unspsc']: log_e('PCN', "PCN not in UNSPSC List", 'matl_group')

            # 2. UOM
            oun = str(get_v('oun')).strip()
            uom_ext = str(get_v('po_uom_ext')).strip()
            u_ord = str(get_v('u_price_ord')).strip()
            if oun != uom_ext: log_e('Unit of Measurement', "OUn != PO UOM - Ext", 'oun')
            if uom_ext != u_ord: log_e('Unit of Measurement', "PO UOM - Ext != Order Price Unit", 'po_uom_ext')
            if 'valid_uom' in rules and uom_ext not in rules['valid_uom']: log_e('Unit of Measurement', "UOM not in MyBuy", 'po_uom_ext')

            # 3. Requestor
            req = str(get_v('requestor')).strip()
            if req in banned_reqs: log_e('Requestor', "Requestor is PROC SSC colleague", 'requestor')
            if len(req) > max_req: log_e('Requestor', f"User ID exceeding {max_req} chars", 'requestor')

            # 4. Preparer
            prep = str(get_v('preparer')).strip()
            if len(prep) > max_prep: log_e('Preparer', f"User ID exceeding {max_prep} chars", 'preparer')

            # 5. Split Accounting
            saa = safe_f(get_v('saa'))
            if saa > 1:
                msg = "Material SAA > 1" if mat_val != "" else "Service SAA > 1"
                log_e('Split Accounting', msg, 'saa')

            # 6. Text
            for c_name in df.columns:
                if 'comment' in c_name.lower():
                    txt = str(row[c_name])
                    if len(txt) > 4000: log_e('Text', f"{c_name} > 4000 chars", c_name)
                    err_c = check_special_characters(txt, special_chars)
                    if err_c: log_e('Text', f"{c_name}: {err_c}", c_name)
            
            short = str(get_v('short_text'))
            if len(short) > max_short: log_e('Text', f"Exceeds {max_short} chars", 'short_text')
            err_short = check_special_characters(short, special_chars)
            if err_short: log_e('Text', f"Short Text: {err_short}", 'short_text')

            # 7. Currency
            curr1 = str(get_v('curr_net')).strip().upper()
            if curr1 in no_dec and net_price % 1 != 0: log_e('Currency', "Currency with decimal error", 'net_price')
            
            curr2 = str(get_v('curr_unit')).strip().upper()
            u_price = safe_f(get_v('unit_price'))
            if curr2 in no_dec and u_price % 1 != 0: log_e('Currency', "Currency with decimal error", 'unit_price')

            # 8. Schedule Line
            if safe_f(get_v('schd')) > 1: log_e('Schedule Line', "Schedule line more than one per item", 'schd')

            # 9. Vendor
            if 'active_vendors' in rules and vnd not in rules['active_vendors']: log_e('Vendor', "Invalid supplier", 'vendor')
            if check_intercompany_vendor(vnd): log_e('Vendor', "Intercompany PO", 'vendor')
            slm = str(get_v('slm')).strip()
            if not check_mandatory(vnd) or not check_mandatory(slm): log_e('Vendor', "No supplier SLM ID", 'slm')
            if 'suppress_vendors' in rules and vnd in rules['suppress_vendors']: log_e('Vendor', "Suppress PO supplier", 'vendor')

            # 10. Unloading Point
            unload = str(get_v('unload')).strip()
            if not unload or unload.lower() == 'nan': log_e('Unloading Point', "Empty unloading point", 'unload')
            elif len(unload) > max_unload: log_e('Unloading Point', f"Exceeds {max_unload} chars", 'unload')

            # 11. Doc Type
            if po_type := str(get_v('type')).strip():
                if po_type in rules.get('req_material', set()) and mat_val == "": log_e('Doc Type', "Direct PO DOC Type without material", 'type')
                if po_type in rules.get('no_material', set()) and mat_val != "": log_e('Doc Type', "Indirect PO DOC Type with material", 'type')

            # 12. Payment Term
            payt = str(get_v('payt')).strip()
            if 'valid_payt' in rules and payt not in rules['valid_payt']: log_e('Payment Term', "Payment term not in MyBuy", 'payt')

            # 13. FOC
            if ir_clean == 'FOC' and still_pay_qty < 1: log_e('FOC', "FOC Service item < 1", 'still_pay_qty')

            # 14. Logic Checks
            if still_pay_amt < 0: log_e('Logic Checks', "Negative still to pay amount", 'still_pay_amt')
            if mat_val == "" and still_pay_qty == 0 and still_pay_amt > 0: log_e('Logic Checks', "Open amount but no quantity (Service)", 'still_pay_amt')
            amt_eur = safe_f(get_v('amt_eur'))
            if 0 < amt_eur <= small_val_limit: log_e('Logic Checks', f"PO value < {small_val_limit} EUR", 'amt_eur')

            # 15. Additional Pricing
            if safe_f(get_v('per')) > 1: log_e('Additional Pricing', "Additional pricing (Per > 1)", 'per')

            # 16. Incoterm
            if not check_mandatory(str(get_v('incot'))): log_e('Incoterm', "Incoterm is missing", 'incot')

        # --- 5. CONSOLIDATE ---
        po_categories.append(p_cat); po_statuses.append(p_stat); po_remarks.append(p_rem)
        row_joined_errs = []
        for cat in category_list:
            msg = " | ".join(row_cat_errors[cat])
            cat_remarks_collector[cat].append(msg)
            if msg: row_joined_errs.append(msg)
        error_details.append(" | ".join(row_joined_errs))

    # --- 6. FINAL BUILD ---
    df_out = df.copy()
    df_out.insert(0, 'PO Category', po_categories)
    df_out.insert(1, 'PO Status', po_statuses)
    df_out.insert(2, 'Remarks', po_remarks)
    df_out.insert(3, 'Error_Details', error_details)
    for cat in category_list: df_out[f"{cat}_Remarks"] = cat_remarks_collector[cat]
    
    return df_out, bad_cells, category_list

def to_excel_po_download(full_df, bad_cells, category_list):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: 
        workbook = writer.book
        red_foramt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#005eb8', 'font_color': 'white', 'border': 1})
        bold_format = workbook.add_format({'bold': True})

        # metrics calculation 
        total = len(full_df)
        errors = len(full_df[full_df['Error_Details'] != ""])

        # Dashboard 
        ws0 = workbook.add_worksheet('Dashboard_Summary')
        ws0.write('B2', "PO Analysis Summary", header_format)
        ws0.write('C2', "Count", header_format)
        ws0.write('B3', "Total Records", bold_format)
        ws0.write('C3', total)
        ws0.write('B4', "Records with Errors", bold_format)
        ws0.write('C4', errors)

        # Breakdown
        r = 7
        ws0.write('B6', "Errors by Category", header_format)
        ws0.write('C6', "Count", header_format)
        for cat in category_list: 
            col_name = f"{cat}_Remarks"
            if col_name in full_df.columns: 
                count = len(full_df[full_df[col_name].str.strip() != ""])
                ws0.write(r, 1, cat)
                ws0.write(r, 2, count)
                r += 1
        
        # Breakdown by status
        r += 2
        ws0.write(r, 1, "Breakdown by Status", header_format)
        ws0.write(r, 2, "Count", header_format)
        r += 1
        if 'PO Status' in full_df.columns:
            status_counts = full_df['PO Status'].value_counts()
            for status, count in status_counts.items():
                ws0.write(r, 1, status)
                ws0.write(r, 2, count)
                r += 1

        ws0.set_column(1, 1, 40)

        # Raw Data
        # drop helper columns
        drop_cols = ['Error_Details'] + [f"{c}_Remarks" for c in category_list]
        clean_df = full_df.drop(columns=[c for c in drop_cols if c in full_df.columns])
        clean_df.to_excel(writer, index=False, sheet_name='Full_Raw_Data')
        ws1 = writer.sheets['Full_Raw_Data']

        for i, col in enumerate(clean_df.columns): ws1.write(0, i, col, header_format)

        # highlight cells 
        col_map = {name: i for i, name in enumerate(clean_df.columns)}
        for row_idx, col_name in bad_cells:
            if col_name in col_map: 
                excel_col_idx = col_map[col_name]
                excel_row_idx = row_idx + 1
                try: 
                    val = clean_df.iat[row_idx, excel_col_idx]
                    if pd.isna(val): ws1.write_blank(excel_row_idx, excel_col_idx, None, red_foramt)
                    else: ws1.write(excel_row_idx, excel_col_idx, val, red_foramt)
                except: pass
        ws1.freeze_panes(1, 0)

        # Other sheets
        # Errors Categories Tabs
        for cat in category_list: 
            col_name = f"{cat}_Remarks"
            if col_name in full_df.columns: 
                subset = full_df[full_df[col_name] != ""]
                if not subset.empty: 
                    # Raw data + ONLY the specific remark column
                    # Remove other internal columns 
                    base_data = subset.drop(columns=[c for c in drop_cols if c != col_name and c in subset.columns])

                    # Move remark to front
                    cols = [col_name] + [c for c in base_data.columns if c != col_name]
                    final_view = base_data[cols]

                    final_view.to_excel(writer, index=False, sheet_name=cat[:30]) # Sheet name limit 31 chars

                    # Header format
                    ws = writer.sheets[cat[:30]]
                    for i, c in enumerate(final_view.columns): ws.write(0, i, c, header_format)
                    ws.set_column(0, 0, 50)

        # Status Tabs
        if 'PO Status' in full_df.columns:
            unique_statuses = full_df['PO Status'].unique()
            for status in unique_statuses:
                # Filter rows with this status 
                status_subset = clean_df[clean_df['PO Status'] == status]

                if not status_subset.empty:
                    # clean sheet name (remove invalid characters)
                    sheet_name = f"Status_{str(status)[:20]}".replace('/', '_')

                    status_subset.to_excel(writer, index=False, sheet_name=sheet_name)

                    # Apply formatting
                    ws_stat = writer.sheets[sheet_name]
                    for i, c in enumerate(status_subset.columns): ws_stat.write(0, i, c, header_format)
                    ws_stat.set_column(0, len(status_subset.columns)-1, 15)
                    ws_stat.freeze_panes(1, 3)

    return output.getvalue()


# =================================
# USER INTERFACE
# =================================

def main(): 
    with st.sidebar:
        st.title("🛡️ Workbench")
        st.write("v9.0 - Hybrid Rule Engine")
        st.markdown("---")

        # --- 1. Module Switches ---
        with st.expander("Enable/Disable Modules", expanded=True):
            show_smd = st.toggle("SMD Analysis", value=True)
            show_email = st.toggle("Email Validation", value=True)
            show_po = st.toggle("PO Analysis", value=True)

        # --- 2. Dynamic Navigation ---
        # Build the menu list based on switches
        available_pages = ["Home"]
        if show_smd: available_pages.append("SMD Analysis")
        if show_email: available_pages.append("Email Validation")
        if show_po: available_pages.append("PO Analysis")

        st.markdown("Navigation")
        task = st.radio("Go to:", available_pages)
        st.markdown("---")
        
        target_cocd = "3072"
        target_porg = "3072"
        region_mode = "APAC"
        
        if task == "SMD Analysis":
            st.header("⚙️ Configuration")
            region_mode = st.radio("Region:", ["APAC", "EU"], horizontal=True)
            target_cocd = st.text_input("Target CoCd:", value="3072" if region_mode == "APAC" else "1040")
            target_porg = st.text_input("Target POrg:", value="3072" if region_mode == "APAC" else "1040", help="Enter multiple POrgs separated by commas.")

    # ================================================
    # PAGE LOGIC
    # ================================================

    if task == "Home":
        st.title("Procurement Workbench")
        st.markdown("""
                    Welcome!
                    Use the sidebar to enable or disable specific analysis modules.
                    
                    Available Modules:
                    - SMD Analysis: Validate Supplier Master Data against global and regional rules.
                    - Email Validation: Check vendor email lists for missing contacts or format errors. 
                    - PO Analysis: Analyze Purchase Orders using a dynamic logic matrix.""")
        st.info("Open the sidebar to get started.")

    # --- SMD ANALYSIS ---
    elif task == "SMD Analysis": 
        st.title(f"SMD Validation ({region_mode})")
        
        st.subheader("1. Upload Rules Config (Optional)")
        req_file = st.file_uploader("Upload 'SMD_Rules_Config.xlsx'", type=['xlsx'], key='smd_req')
        
        req_df = None
        if req_file:
            req_df = pd.read_excel(req_file)
            st.success("Custom Rules Loaded")
            with st.expander("View Rules"): st.dataframe(req_df)

        st.subheader("2. Upload Raw Data")
        uploaded_file = st.file_uploader("Upload Raw Data", type=['xlsx'], key='smd_raw')

        if uploaded_file and st.button("Run Analysis", type="primary"):
            with st.spinner(f"Analyzing..."):
                # Read all sheets from Excel file
                xls = pd.ExcelFile(uploaded_file)
                all_dfs = []
                for sheet_name in xls.sheet_names:
                    try:
                        sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, dtype=str)
                        all_dfs.append(sheet_df)
                    except: pass
                
                if not all_dfs:
                    st.error("No data found in the file.")
                else:
                    df = pd.concat(all_dfs, ignore_index=True)

                    results, bad_cells = run_smd_analysis(df, req_df, target_cocd, target_porg, region_mode)
                    duplicates_df = get_duplicates_df(df)
                    
                    df_errors_only = results[results['Error_Details'] != ""]

                    p1 = len(results[results['Purchasing_Issues'] != ""])
                    p2 = len(results[results['Org_Finance_Issues'] != ""])
                    p3 = len(results[results['Master_Data_Issues'] != ""])
                    gen = len(results[results['General_Errors'] != ""])
                
                    metrics = {
                        'Total': len(results), 'Correct': len(results) - len(df_errors_only), 'Errors': len(df_errors_only),
                        'Duplicates': len(duplicates_df),
                        'Purchasing': p1, 'Org_Finance': p2, 'Master_Data': p3, 'General': gen
                    }
                
                    error_bkdown = pd.DataFrame()
                    if not df_errors_only.empty:
                        error_bkdown = df_errors_only['Error_Details'].str.split(' \| ').explode().value_counts().reset_index()
                        error_bkdown.columns = ['Error Description', 'Count']

                    st.metric("Total Errors", len(df_errors_only))
                    fname = f"SMD_Report_{target_cocd}.xlsx"
                    data = to_excel_download_smd(results, df_errors_only, duplicates_df, metrics, error_bkdown, bad_cells)
                    st.download_button("Download Report", data, fname)

    # --- PO ANALYSIS ----
    elif task == "PO Analysis": 
        st.title("Purchase Order Analysis")

        st.subheader("1. Upload PO Rules")
        po_rules_file = st.file_uploader("Upload 'PO_Rules_Config.xlsx'", type=['xlsx'], key='po_rules')

        st.subheader("2. Upload PO Data")
        po_raw_file = st.file_uploader("Upload Raw PO Data", type=['xlsx'], key='po_raw')

        if po_rules_file and po_raw_file: 
            if st.button("Run PO Check", type="primary"):
                with st.spinner("Analyzing..."):
                    xls = pd.ExcelFile(po_raw_file)
                    all_dfs = []
                    for sheet_name in xls.sheet_names:
                        try: 
                            sheet_df = pd.read_excel(po_raw_file, sheet_name=sheet_name, dtype=str, keep_default_na=False, na_values=None)
                            all_dfs.append(sheet_df)
                        except: pass
                    
                    if not all_dfs:
                        st.error("No data found in PO file.")
                        st.stop()

                    df_po = pd.concat(all_dfs, ignore_index=True)

                    # Run dynamic PO engine
                    res_po, bad_cells, cat_list = run_po_analysis_dynamic(df_po, po_rules_file)

                    err_count = len(res_po[res_po['Error_Details'] != ""])
                    st.metric("PO Lines with Errors", err_count)

                    data = to_excel_po_download(res_po, bad_cells, cat_list)
                    st.download_button("Download PO Report", data, "PO_Analysis_Report.xlsx")
    
    # --- EMAIL ---
    elif task == "Email Validation": 
        st.title("Vendor Email Validation")
        uploaded_email = st.file_uploader("Upload Email List", type=['xlsx'])

        if uploaded_email:
            if st.button("Run Email Check", type="primary"):
                with st.spinner("Analyzing..."): 
                    xls = pd.ExcelFile(uploaded_email)
                    all_dfs = []
                    for sheet_name in xls.sheet_names: 
                        try: 
                            sheet_df = pd.read_excel(uploaded_email, sheet_name=sheet_name, dtype=str)
                            all_dfs.append(sheet_df)
                        except: pass
                    
                    if not all_dfs: 
                        st.error("No data found in Email file.")
                        st.stop()

                    df_email = pd.concat(all_dfs, ignore_index=True)
                    res_email = run_email_analysis(df_email)

                    # filter errors
                    err_df = res_email[res_email['Email_Validation_Errors'] != ""]

                    # Calculate counts
                    error_count = len(err_df)
                    metrics_dict = {
                        'Total': len(res_email),
                        'Correct': len(res_email) - error_count,
                        'Errors': error_count
                    }

                    if not err_df.empty:
                        st.metric("Issues Found", error_count)

                        # Generate download
                        data = to_excel_email_download(res_email, metrics_dict)
                        st.download_button("Download Email Report", data, "Email_Validation.xlsx")
                    else: st.success("Valid!")

if __name__ == "__main__":
    main()


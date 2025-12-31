import streamlit as st 
import pandas as pd 
import numpy as np 
import io

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
        # red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        # header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        bold_format = workbook.add_format({'bold': True})

        # --- Sheet 1: Dashboard ---
        worksheet0 = workbook.add_worksheet('Dashboard_Summary')
        worksheet0.write('B2', "High Level Summary", header_format)
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

    # --- 1. DYNAMIC CONFIGURATION (Load Rules from Excel) ---
    req_dict = {'Mandatory': [], 'Empty': []}
    
    if requirements_df is not None:
        requirements_df.columns = [c.strip().title() for c in requirements_df.columns]
        # Check required columns
        if 'Field' in requirements_df.columns and 'Rule' in requirements_df.columns:
            for idx, row in requirements_df.iterrows():
                field = str(row['Field']).strip()
                rule = str(row['Rule']).strip().lower()
                
                # Check Region Scope if column exists (Default to ALL if missing)
                rule_region = str(row['Region']).strip().upper() if 'Region' in requirements_df.columns else 'ALL'
                
                # Apply if Region matches selected mode OR is ALL
                if rule_region == 'ALL' or rule_region == region.upper():
                    if 'mandatory' in rule: req_dict['Mandatory'].append(field)
                    if 'empty' in rule: req_dict['Empty'].append(field)

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
        for field in req_dict['Mandatory']:
            if field in df.columns and not check_mandatory(row[field]):
                log_err(gen, f"{field} is missing", field, index)
        for field in req_dict['Empty']:
            if field in df.columns and not check_must_be_empty(row[field]):
                log_err(gen, f"{field} must be empty", field, index)

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
            if target_porg and check_mandatory(porg) and porg != target_porg: log_err(p2, f"POrg != {target_porg}", 'POrg', index)

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
        
        # 8. Telephone 1 & "+" for MY ONLY
        if 'Telephone 1' in df.columns:
            phone = str(row['Telephone 1']).strip()
            if is_local:  # Logic: IF Country is 'MY'
                if not check_mandatory(phone): 
                    log_err(gen, "Tel 1 missing", 'Telephone 1', index)
                elif "+" not in phone: 
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
    vendor_groups = df_out.groupby('Vendor')
    email_errors = []
    for idx, row in df_out.iterrows():
        errs = []
        vendor = row['Vendor']
        vendor_data = vendor_groups.get_group(vendor)
        notes = str(row['Communication link notes']).strip()
        if not check_mandatory(notes): errs.append("Comm. Notes Empty")
        min_id = vendor_data['ID'].min()
        if row['ID'] == min_id:
            flag = str(row['#']).strip().upper()
            if flag != 'X' and flag != '1': errs.append("Smallest ID not marked Default (#)")
        email_errors.append(" | ".join(errs))
    final_error_col = []
    for idx, row in df_out.iterrows():
        final_error_col.append(email_errors[idx])
    df_out.insert(0, 'Email_Validation_Errors', final_error_col)
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

def run_po_analysis(df, requirements_df): 
    """
    Dynamic PO Validation based on Uploaded Rules. 
    Rules structure: Field, Rule, Category, Error_Message 
    """

    df_out = df.copy()
    df.columns = df.columns.str.strip()

    # Parse rules
    req_dict = {'Mandatory': [], 'Numeric': [], 'NoSpecial': []}

    if requirements_df is not None: 
        requirements_df.columns = [c.strip().title() for c in requirements_df.columns]
        if 'Field' in requirements_df.columns and 'Rule' in requirements_df.columns: 
            for idx, row in requirements_df.iterrows(): 
                field = str(row['Field']).strip()
                rule = str(row['Rule']).strip()
                cat = str(row['Category']).strip() if 'Category' in requirements_df.columns else 'General'
                msg = str(row['Error_Message']).strip() if 'Category' in requirements_df.columns else f"{field} Error"

                rule_obj = {'field': field, 'cat': cat, 'msg': msg}

                if 'mandatory' in rule: req_dict['Mandatory'].append(rule_obj)
                if 'numeric' in rule or 'greaterthan' in rule: 
                    req_dict['Numeric'].append(rule_obj)
                    # add more logic if needed

    # Iterate rows 
    all_error_details = []
    bad_cells = []

    # Dynamic category dictionaries to store errors per category
    category_map = {}

    for index, row in df.iterrows(): 
        row_errors = []

        # helper
        def log_po_err(msg, col, cat): 
            row_errors.append(msg)
            bad_cells.append((index, col))
            # store in category specific list for later split 
            if index not in category_map: 
                category_map[index] = {}
            if cat not in category_map[index]: 
                category_map[index][cat] = []
                category_map[index][cat].append(msg)
        
        # Mandatory Checks 
        for item in req_dict['Mandatory']: 
            if item['field'] in df.columns and not check_mandatory(row[item['field']]): 
                log_po_err(item['msg'], item['field'], item['cat'])
        
        # Numeric checks (>0)
        for item in req_dict['Numeric']: 
            if item['field'] in df.columns:
                val = row[item['field']]
                if check_mandatory(val): 
                    try: 
                        if float(val) <= 0: 
                            log_po_err(item['msg'], item['field'], item['cat'])
                    except: 
                        log_po_err(f"{item['field']} must be numeric", item['field'], item['cat'])
        
        all_error_details.append(" | ".join(row_errors))

    # Add columns to DF
    # need to find all unique categories used 
    all_cats = set()
    if requirements_df is not None and 'Category' in requirements_df.columns: 
        all_cats = set(requirements_df['Category'].dropna().unique())
    else: 
        all_cats = {'General'}
    
    for cat in all_cats:
        cat_col_name = f"{cat}_Errors"
        cat_values = []
        for i in range(len(df)):
            if i in category_map and cat in category_map[i]: 
                cat_values.append(" | ".join(category_map[i][cat]))
            else: 
                cat_values.append("")
        df_out.insert(0, cat_col_name, cat_values)
    
    return df_out, bad_cells, list(all_cats)

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
        correct = total - errors

        # Dashboard 
        ws0 = workbook.add_worksheet('Dashboard_Summary')
        ws0.write('B2', "PO Analysis Summary", header_format)
        ws0.write('B3', "Total Records", bold_format)
        ws0.write('C3', total)
        ws0.write('B4', "Records with Errors", bold_format)
        ws0.write('C4', errors)

        # Breakdown
        ws0.write('B7', "Errors by Category", header_format)
        ws0.write('C7', "Count", header_format)
        r = 8 
        for cat in category_list: 
            col_name = f"{cat}_Errors"
            if col_name in full_df.columns: 
                count = len(full_df[full_df[col_name] != ""])
                ws0.write(r, 1, cat)
                ws0.write(r, 2, count)
                r += 1
        ws0.set_column(1, 1, 40)

        # Raw Data
        # drop helper columns
        drop_cols = ['Error_Details'] + [f"{c}_Errors" for c in category_list]
        clean_df = full_df.drop(columns=[c for c in drop_cols if c in full_df.columns])

        clean_df.to_excel(writer, index=False, sheet_name='Full_Raw_Data')
        ws2 = writer.sheets['Full_Raw_Data']

        for i, col in enumerate(clean_df.columns): ws2.write(0, i, col, header_format)

        # highlight cells 
        col_map = {name: i for i, name in enumerate(clean_df.columns)}
        for row_idx, col_name in bad_cells:
            if col_name in col_map: 
                excel_col_idx = col_map[col_name]
                excel_row_idx = row_idx + 1
                try: 
                    val = clean_df.iat[row_idx, excel_col_idx]
                    if pd.isna(val): ws2.write_blank(excel_row_idx, excel_col_idx, None, red_foramt)
                    else: ws2.write(excel_row_idx, excel_col_idx, val, red_foramt)
                except: pass
        ws2.freeze_panes(1, 0)

        # Categories
        for cat in category_list: 
            col_name = f"{cat}_Errors"
            if col_name in full_df.columns: 
                subset = full_df[full_df[col_name] != ""]
                if not subset.empty: 
                    # show the first 3 cols of raw data + error column
                    preview_cols = list(clean_df.columns[:3]) + [col_name]
                    subset[preview_cols].to_excel(writer, index=False, sheet_name=f"{cat}_Errors")
    return output.getvalue()

def main(): 
    with st.sidebar:
        st.title("🛡️ Workbench")
        st.write("v8.0 - Hybrid Rule Engine")
        st.markdown("---")
        task = st.radio("Select Module", ["SMD Analysis", "Email Validation", "PO Analysis"])
        st.markdown("---")
        
        target_cocd = "3072"
        target_porg = "3072"
        region_mode = "APAC"
        
        if task == "SMD Analysis":
            st.header("⚙️ Configuration")
            region_mode = st.radio("Region:", ["APAC", "EU"], horizontal=True)
            target_cocd = st.text_input("Target CoCd:", value="3072" if region_mode == "APAC" else "1040")
            target_porg = st.text_input("Target POrg:", value="3072" if region_mode == "APAC" else "1040")

    if task == "Home":
        st.title("Procurement Workbench")
        st.info("Select a module from the sidebar.")

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
                            sheet_df = pd.read_excel(po_raw_file, sheet_name=sheet_name, dtype=str)
                            all_dfs.append(sheet_df)
                        except: pass
                    
                    if not all_dfs:
                        st.error("No data found in PO file.")
                        st.stop()

                    df_po = pd.concat(all_dfs, ignore_index=True)
                    rules_df = pd.read_excel(po_rules_file)

                    # Run dynamic PO engine
                    res_po, bad_cells, cat_list = run_po_analysis(df_po, rules_df)

                    err_count = len(res_po[res_po['Error_Details'] != ""])
                    st.metric("PO Lines with Errors", err_count)

                    if err_count > 0:
                        st.dataframe(res_po[res_po['Error_Details'] != ""].head())

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

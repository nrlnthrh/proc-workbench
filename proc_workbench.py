import streamlit as st 
import pandas as pd 
import numpy as np 
import io
import re
import xlsxwriter

# ==========================================
# 1. PAGE CONFIGURATION - Layout
# ==========================================
st.set_page_config(
    page_title="PROCleans",
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

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# ===========================================================
# 2. HELPER FUNCTIONS - Mandatory functions for Analysis
# ===========================================================

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

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# ==========================================
# 3. SMD ANALYSIS LOGIC (Hybrid Engine)
# ==========================================

def check_postal_code(country, postal_code, rules_dict=None):
    if not check_mandatory(postal_code): return None
    postal_str = str(postal_code).strip()
    
    if country in rules_dict:
        # Handle tuple (min, max) from Excel loader
        val = rules_dict[country]
        if isinstance(val, tuple):
            min_len, max_len = val
        else:
            # Fallback if dictionary has single int
            min_len = max_len = int(val)

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
                s_dupes['Duplicate Reason'] = 'Synertrade ID used by multiple Vendors'
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

def load_smd_config(config_file):
    config = {
        'field_rules': [],
        'incoterms': {},
        'postal_codes': {},
        'reference_lists': {} # dynamic storage for lists
    }
    
    if not config_file: return config

    try:
        xls = pd.ExcelFile(config_file)
        
        # 1. Field Rules
        if 'Field_Rules' in xls.sheet_names:
            df_rules = pd.read_excel(xls, 'Field_Rules')
            df_rules.columns = [c.strip().title() for c in df_rules.columns]
            config['field_rules'] = df_rules.to_dict('records')
        
        # 2. Incoterm Rules (Optional)
        if 'Incoterm_Rules' in xls.sheet_names:
            df_inco = pd.read_excel(xls, 'Incoterm_Rules')
            # Normalize headers
            df_inco.columns = [c.strip().title() for c in df_inco.columns] 
            
            # Store as list of dictionaries
            config['incoterms'] = []
            for _, row in df_inco.iterrows():
                config['incoterms'].append({
                    'code': str(row.get('Incoterm', '')).strip().upper(),
                    'rule': str(row.get('Rule', '')).strip().lower(),
                    'val': str(row.get('Value', '')).strip(),
                    'msg': str(row.get('Error_Message', 'Incoterm Error')).strip()
                })
            
        # 3. Postal Codes
        if 'Postal_Codes' in xls.sheet_names:
            df_postal = pd.read_excel(xls, 'Postal_Codes')
            # Assuming format: Country | Min | Max
            for _, row in df_postal.iterrows():
                try:
                    c = str(row.iloc[0]).strip().upper()
                    mn = int(row.iloc[1])
                    mx = int(row.iloc[2])
                    config['postal_codes'][c] = (mn, mx)
                except: pass

        # 4. Reference Lists (Dynamic Loading)
        # Any column found here becomes a validation list!
        if 'Reference_Lists' in xls.sheet_names:
            df_ref = pd.read_excel(xls, 'Reference_Lists')
            for col in df_ref.columns:
                # Clean column name (e.g., "REF_PAYT")
                key = col.strip()
                # Store set of valid values
                config['reference_lists'][key] = set(df_ref[col].dropna().astype(str).str.strip())
        
        # 5. Column Mapping Sheet
        if 'Column_Mapping' in xls.sheet_names:
            df_map = pd.read_excel(xls, 'Column_Mapping')
            for _, row in df_map.iterrows():
                std = str(row['Standard_Column']).strip()
                # Split possible headers by comma
                possibles = [x.strip() for x in str(row['Possible_Headers']).split(',')]
                config['column_mapping'][std] = possibles
                
    except Exception as e:
        st.error(f"Error reading SMD Config: {e}")
    return config

def run_smd_analysis(df, config_file, target_cocd, target_porg):
    df_out = df.copy()
    df.columns = df.columns.str.strip()

    # --- Rules ---
    rules_data = load_smd_config(config_file)

    # Header Standardization
    col_map_config = rules_data.get('column_mapping', {})
    current_cols_lower = {c.lower(): c for c in df_out.columns}
    rename_map = {}

    for std_col, possible_names in col_map_config.items():
        if std_col in df_out.columns: continue
        for alias in possible_names:
            if alias.lower() in current_cols_lower:
                rename_map[current_cols_lower[alias.lower()]] = std_col
                break
    
    if rename_map:
        df_out = df_out.rename(columns=rename_map)
    
    # Update df reference for processing
    df = df_out

    field_rules = rules_data.get('field_rules', [])
    ref_lists = rules_data.get('reference_lists', {})
    postal_rules = rules_data.get('postal_codes', {})

    # --- target porg (split by comma) ---
    # example: "0001, 0002" -> ['0001', '0002']
    valid_porg_list = []
    if target_porg: 
        valid_porg_list = [p.strip() for p in str(target_porg).split(',') if p.strip()]

    # --- 1. DYNAMIC CONFIGURATION (Load Rules from Excel) ---
    req_dict = {'Mandatory': [], 'Empty': [], 'InList': [], 'Contains': []}
    
    for rule in field_rules:
        field = str(rule.get('Field')).strip()
        rtype = str(rule.get('Rule')).strip().lower()
        cat = str(rule.get('Category', 'General')).strip()
        val = str(rule.get('Value', '')).strip()

        rule_obj = {'field': field, 'cat': cat, 'val': val}
        if 'mandatory' in rtype: req_dict['Mandatory'].append(rule_obj)
        elif 'empty' in rtype: req_dict['Empty'].append(rule_obj)
        elif 'inlist' in rtype: req_dict['InList'].append(rule_obj)
        elif 'contains' in rtype: req_dict['Contains'].append(rule_obj)

    # --- 2. HARDCODED LOGIC & COLUMN MAPPING ---
    valid_vendor_ids = set()
    if 'Vendor' in df.columns: valid_vendor_ids = set(df['Vendor'].astype(str).str.strip())
    
    col_payt_fin = 'PayT'
    col_payt_purch = 'PayT.1'

    all_error_details, p1_list_col, p2_list_col, p3_list_col, gen_list_col = [], [], [], [], []

    bad_cells = []
    def log_err(category_list, msg, col_name, idx, specific_cat_name):
        category_list.append(msg)
        if col_name and col_name in df.columns: bad_cells.append((idx, col_name, specific_cat_name))

    # --- 3. EXECUTION ---
    for index, row in df.iterrows():
        p1, p2, p3, gen = [], [], [], [] 
        country = str(row.get('Cty','')).strip().upper()
        is_local = (country == 'MY')

        # --- A. Apply Dynamic Rules (Excel) ---
        def get_target_list(cat_name):
            c = cat_name.lower()
            if 'purchasing' in c: return p1, 'Purchasing'
            if 'org' in c or 'finance' in c: return p2, 'Org Finance'
            if 'master' in c or 'vendor' in c: return p3, 'Master Data'
            return gen, 'General' # default
        
        for item in req_dict['Mandatory']:
            t_list, t_name = get_target_list(item['cat'])
            if item['field'] in df.columns and not check_mandatory(row[item['field']]):
                log_err(t_list, f"{item['field']} is missing", item['field'], index, t_name)
        for item in req_dict['Empty']:
            if item['field'] in df.columns and not check_must_be_empty(row[item['field']]):
                log_err(t_list, f"{item['field']} must be empty", item['field'], index, t_name)
        
        # InList Check (Checks against Reference Lists)
        for item in req_dict['InList']:
            f_name = item['field']
            ref_key = item['val'] # e.g. "REF_PAYT"
            if f_name in df.columns:
                val = str(row[f_name]).strip()
                if check_mandatory(val):
                    # Check if the list exists in our config
                    if ref_key in ref_lists:
                        if val not in ref_lists[ref_key]:
                            log_err(t_list, f"{f_name} invalid (not in {ref_key})", f_name, index, t_name)
        
        # Contains Check
        for item in req_dict['Contains']:
            f_name = item['field']
            chars = [c.strip() for c in item['val'].split(',')]
            if f_name in df.columns:
                val = str(row[f_name]).strip()
                if check_mandatory(val) and not any(c in val for c in chars):
                    log_err(t_list, f"{f_name} missing required value", f_name, index, t_name)

        # --- B. Hardcoded Business Logic (Things too complex for simple Excel rules) ---
        # 1. Currency (Check all variations)
        crcy_cols = [c for c in df.columns if 'crcy' in c.lower() or 'currency' in c.lower()]
        if crcy_cols:
            has_currency = any(check_mandatory(row[c]) for c in crcy_cols)
            if not has_currency:
                log_err(p1, "Currency missing (All Cols)", crcy_cols[0], index, 'Purchasing')

        if col_payt_purch in df.columns and not check_mandatory(row[col_payt_purch]): log_err(p1, "Purch PayT Missing", col_payt_purch, index, 'Purchasing')
        
        # 2. Incoterms
        if 'IncoT' in df.columns:
            incot = str(row.get('IncoT', '')).strip().upper()
            inco2 = str(row.get('Inco. 2', '')).strip().lower()

            if check_mandatory(incot):
                # Iterate through rules loaded from Excel
                for rule in rules_data.get('incoterms', []):
                    target_code = rule['code']
                    rtype = rule['rule']
                    rval = rule['val'].lower()
                    msg = rule['msg']

                    # Apply rule if code matches OR rule is for ALL
                    if target_code == 'ALL' or target_code == incot:
                        
                        if rtype == 'contains':
                            # Check if ANY of the values exist in Inco 2
                            opts = [x.strip() for x in rval.split(',')]
                            if not any(opt in inco2 for opt in opts):
                                log_err(p1, msg, 'Inco. 2', index, 'Purchasing')

                        elif rtype == 'startswith':
                            # Check if Inco 2 starts with any option
                            opts = [x.strip() for x in rval.split(',')]
                            if not any(inco2.startswith(opt) for opt in opts):
                                log_err(p1, msg, 'Inco. 2', index, 'Purchasing')

                        elif rtype == 'equals' and rval == 'obsolete':
                            # Flag the IncoT itself
                            log_err(p1, msg, 'IncoT', index, 'Purchasing')

            else:
                log_err(p1, "IncoT missing", 'IncoT', index, 'Purchasing')

        # 3. Payment Terms
        if col_payt_fin in df.columns and col_payt_purch in df.columns:
            fin, pur = str(row[col_payt_fin]).strip(), str(row[col_payt_purch]).strip()
            if fin != pur: log_err(p2, f"PayT Mismatch ({fin} vs {pur})", col_payt_fin, index, 'Org Finance')
        
        # 4. Org Check
        if 'CoCd' in df.columns and str(row['CoCd']).strip() != target_cocd: log_err(p2, f"CoCd != {target_cocd}", 'CoCd', index, 'Org Finance')
        if 'POrg' in df.columns:
            porg = str(row['POrg']).strip()
            if valid_porg_list:
                if check_mandatory(porg):
                    if porg not in valid_porg_list: 
                        log_err(p2, f"POrg != {valid_porg_list}", 'POrg', index, 'Org Finance')
                else:
                    log_err(p2, "POrg is missing", 'POrg', index, 'Org Finance')

        # 5. Tax Logic (At least 1 required)
        tax_candidates = [c for c in df.columns if ('tax' in c.lower() or 'vat' in c.lower()) 
                          and 'identification' not in c.lower() and 'liable' not in c.lower() and 'equal' not in c.lower()]
        if tax_candidates:
            has_tax = any(check_mandatory(row[c]) for c in tax_candidates)
            if not has_tax: 
                gen.append("At least one Tax ID Required")
                bad_cells.append((index, tax_candidates[0], 'General'))

        # 6. Postal Code
        if 'Postl Code' in df.columns:
            postal_err = check_postal_code(country, row['Postl Code'], postal_rules)
            if postal_err: log_err(gen, postal_err, 'Postl Code', index, 'General')

        # 7. Duplicate Logic (Alt Payee Scope)
        if 'AltPayeeAc' in df.columns:
            alt = str(row['AltPayeeAc']).strip()
            if check_mandatory(alt) and alt not in valid_vendor_ids:
                log_err(p3, "AltPayee Not in Scope", 'AltPayeeAc', index, 'Master Data')
        
        # 8. Telephone 1 logic
        if 'Telephone 1' in df.columns:
            phone = str(row['Telephone 1']).strip()

            # Rule 1: Mandatory for Global
            if not check_mandatory(phone): 
                log_err(gen, "Tel 1 missing", 'Telephone 1', index, 'General')
            
            # Rule 2: "+" symbol only for CoCd 3072
            elif str(target_cocd).strip() == '3072' and "+" not in phone: 
                log_err(gen, "Tel 1 missing '+'", 'Telephone 1', index, 'General')
                    
        # 9. Synertrade Supplier ID
        if 'Synertrade Supplier ID' in df.columns:
                syn_id = str(row['Synertrade Supplier ID'])
                if not check_mandatory(syn_id):
                    log_err(p3, "Synertrade ID missing", 'Synertrade Supplier ID', index, 'Master Data')

        # Consolidate
        all_errs = p1 + p2 + p3 + gen
        all_error_details.append(" | ".join(all_errs))
        p1_list_col.append(" | ".join(p1))
        p2_list_col.append(" | ".join(p2))
        p3_list_col.append(" | ".join(p3))
        gen_list_col.append(" | ".join(gen))

    df_out.insert(0, 'General Errors', gen_list_col)
    df_out.insert(0, 'Master Data Issues', p3_list_col)
    df_out.insert(0, 'Org Finance Issues', p2_list_col)
    df_out.insert(0, 'Purchasing Issues', p1_list_col)
    df_out.insert(0, 'Error Details', all_error_details)
    
    df_out['Row Index'] = df_out.index + 2
    df_out['Primary ID'] = df_out.apply(get_primary_id, axis=1)
    df_out['Vendor ID'] = df_out.get('Vendor', 'N/A')
    df_out['Name 1'] = df_out.get('Name 1', 'N/A')
    df_out['Company Code'] = df_out.get('CoCd', 'N/A')
    
    return df_out, bad_cells

def to_excel_download_smd(full_df, df_errors, duplicates_df, metrics_dict, error_breakdown_df, bad_cells):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#005eb8', 'font_color': 'white', 'border': 1})
        bold_format = workbook.add_format({'bold': True})
        
        # Helper for highlight
        def write_with_highlight(dataframe, sheet_name, target_category='ALL', freeze_col=0):
            dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]

            # Format headers
            for i, col in enumerate(dataframe.columns):
                ws.write(0, i, col, header_format)
            
            # Map columns
            col_map = {name: i for i, name in enumerate(dataframe.columns)}

            # Iterate rows and highlight errors
            for excel_row_offset, original_idx in enumerate(dataframe.index):
                excel_row = excel_row_offset + 1 

                # check for columns present in specific sheet
                for row_idx, col_name, cat in bad_cells:
                    if row_idx == original_idx and col_name in col_map:

                        if target_category == 'ALL' or target_category == cat:
                            col_idx = col_map[col_name]
                            val = dataframe.at[original_idx, col_name]

                            if pd.isna(val): ws.write_blank(excel_row, col_idx, None, red_format)
                            else: ws.write(excel_row, col_idx, val, red_format)
            # Final formatting
            ws.freeze_panes(1, freeze_col + 1)

        # --- Sheet 1: Dashboard ---
        worksheet0 = workbook.add_worksheet('Dashboard Summary')
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
            ("Org Structure & Finance Issues", metrics_dict.get('Org Finance', 0)),
            ("Master Data ID Issues", metrics_dict.get('Master Data', 0)),
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

        # Helper to clean the internal columns
        cols_to_show = ['Purchasing Issues', 'Org Finance Issues', 'Master Data Issues', 'General Errors']
        internal_cols = ['Row Index', 'Vendor ID', 'Primary ID', 'Company Code', 'Synertrade ID', 'Error Details']

        def clean_df(df_subset, specific_error=None):
            drop_list = [c for c in internal_cols if c in df_subset.columns]

            # Drop other error if focus on specific error
            if specific_error:
                other_errs = [c for c in cols_to_show if c != specific_error]
                drop_list.extend([c for c in other_errs if c in df_subset.columns])

            clean_view = df_subset.drop(columns=drop_list)

            # Move the specific error to the front
            if specific_error and specific_error in clean_view.columns:
                cols = [specific_error] + [c for c in clean_view.columns if c != specific_error]
                clean_view = clean_view[cols]
            return clean_view

        # --- Sheet 2: Error Rows Summary ---
        if not df_errors.empty: 
            cols_to_keep = ['Row Index', 'Primary ID', 'Error Details']
            if 'Name 1' in df_errors.columns: cols_to_keep.insert(2, 'Name 1')

            summary_view = df_errors[cols_to_keep].copy()
            summary_view.to_excel(writer, index=False, sheet_name='Error Summary')
            ws1 = writer.sheets['Error Summary']
            for i, col, in enumerate(summary_view.columns): ws1.write(0, i, col, header_format)
            ws1.set_column(0, 3, 25)
            ws1.set_column(len(summary_view.columns)-1, len(summary_view.columns)-1, 80)
            ws1.freeze_panes(1, 1)
        
        # --- Sheet 3: Full Raw Data ---
        drop_list = [c for c in internal_cols if c in full_df.columns]
        raw_sheet_df = full_df.drop(columns=drop_list)

        # reorder columns
        existing_cols = [c for c in cols_to_show if c in raw_sheet_df.columns]
        other_cols = [c for c in raw_sheet_df.columns if c not in existing_cols]
        final_order = existing_cols + other_cols
        raw_sheet_df = raw_sheet_df[final_order]

        write_with_highlight(raw_sheet_df, 'Full Raw Data', target_category='ALL', freeze_col=3)

        # --- Priority Tabs ---
        priority_config = [
            ('Purchasing Issues', 'Purchasing Validation', 'Purchasing'),
            ('Org Finance Issues', 'Org Finance Validation', 'Org Finance'),
            ('Master Data Issues', 'Master Data Validation', 'Master Data')
        ]

        for col_name, sheet_name, cat_key in priority_config:
            subset = full_df[full_df[col_name] != ""]
            if not subset.empty:
                final_view = clean_df(subset, col_name)
                write_with_highlight(final_view, sheet_name, target_category=cat_key, freeze_col=0) 

        # --- General Split Tabs ---
        if 'General Errors' in full_df.columns:
            # Mandatory
            mand_mask = full_df['General Errors'].str.contains("missing|Required", case=False, na=False)
            mand_df = full_df[mand_mask]
            if not mand_df.empty:
                final_view = clean_df(mand_df, 'General Errors')
                write_with_highlight(final_view, 'Gen Mandatory Missing', target_category='General', freeze_col=0)
            
            # Empty
            empty_mask = full_df['General Errors'].str.contains("must be empty", case=False, na=False)
            emp_df = full_df[empty_mask]
            if not emp_df.empty:
                final_view = clean_df(emp_df, 'General Errors')
                write_with_highlight(final_view, 'Gen Should Be Empty', target_category='General', freeze_col=0)

            # Formatting             
            fmt_mask = (full_df['General Errors'] != "") & \
            (~full_df['General Errors'].str.contains("missing|Required", case=False, na=False)) & \
            (~full_df['General Errors'].str.contains("must be empty", case=False, na=False))

            fmt_df = full_df[fmt_mask]
            if not fmt_df.empty:
                final_view = clean_df(fmt_df, 'General Errors')
                write_with_highlight(final_view, 'Gen Formatting Logic', target_category='General', freeze_col=0)

        # --- Duplicates ---
        if not duplicates_df.empty:
            priority_cols = ['Duplicate Reason', 'Vendor', 'Name 1', 'Synertrade Supplier ID']
            existing_cols = duplicates_df.columns.tolist()
            new_order = [c for c in priority_cols if c in existing_cols] + [c for c in existing_cols if c not in priority_cols]
            duplicates_df = duplicates_df[new_order]

            duplicates_df.to_excel(writer, index=False, sheet_name='Potential Duplicates')
            ws_dupe = writer.sheets['Potential Duplicates']
            for i, col in enumerate(duplicates_df.columns): ws_dupe.write(0, i, col, header_format)
            ws_dupe.freeze_panes(1, 4)
            ws_dupe.set_column(0, 0, 30)

    return output.getvalue()

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
        is_marked = (flag == 'X')

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

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    # Identify "comment" columns 
    comment_cols = [c for c in df.columns if 'comment' in c.lower()]
    # Identify all needed columns
    needed_cols = [v for v in col_map.values() if v] + comment_cols
    # Create a small view for processing
    df_small = df[list(set(needed_cols))]

    category_list = ['PCN', 'Unit of Measurement', 'Requestor', 'Preparer', 'Split Accounting', 
                     'Text', 'Currency', 'Schedule Line', 'Vendor', 'Unloading Point', 
                     'Doc Type', 'Payment Term', 'FOC', 'Logic Checks', 'Additional Pricing', 'Incoterm']
    
    results = {
        'PO Category': [], 'PO Status': [], 'Remarks': [], 'Error_Details': []
    }
    for cat in category_list: results[f"{cat}_Remarks"] = []
    
    bad_cells = []

    def safe_f(val):
        if pd.isna(val) or str(val).strip() == '': return 0.0
        try:
            # This handles both value; "2350.00" and "2,350.00"
            clean_val = str(val).replace(',', '').strip()
            return float(clean_val)
        except: return 0.0

    # --- 4. THE OPTIMIZED LOOP ---
    for idx, row in df_small.iterrows():
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
        p_stat, p_rem, rule_found = "Review", "No matching logic found", False

        if p_cat == "Direct PO":
            p_stat = "Excluded"
            p_rem = "No further action needed."
            rule_found = True
        else:
            p_cat = "Material PO" if gr_val == 'X' else "Service PO"
        
        # Matrix Logic (Simplified evaluation)
        for rule in matrix:
            match = True
            # Check Category match
            r_c = str(rule.get('Category', '')).strip()
            if r_c and r_c != p_cat: match = False
            # Check GR Flag
            if match:
                r_gr = str(rule.get('GR', '')).strip().upper()
                if r_gr == 'X' and gr_val != 'X': match = False
                elif r_gr == 'EMPTY' and gr_val != '': match = False
            # Check Material Flag
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
                    if 'PAYQTY<0' in r_cond and still_pay_qty >= 0: match = False
                    if 'PAYAMT=0' in r_cond and still_pay_amt != 0: match = False
                    if 'PAYAMT>0' in r_cond and still_pay_amt <= 0: match = False
                    if 'PAYAMT<0' in r_cond and still_pay_amt >= 0: match = False
                    if 'IR_EXIST=' in r_cond:
                        target = r_cond.split('IR_EXIST=')[1].split(',')[0].replace('.', '').replace(' ', '')
                        if ir_clean != target: match = False
            if match:
                p_stat, p_rem, rule_found = rule.get('Status'), rule.get('Remark'), True
                break

        # Hardcoded Filter Rule Overrides
        if str(get_v('d_item')) == 'L': p_stat, p_rem = 'Closed', 'Deleted Item. PO line item closed. No further action is required.'
        elif str(get_v('incomplete')) == 'X': p_stat, p_rem = 'Closed', 'Incomplete/on hold item. PO line item closed. No further action is required.'
        elif str(get_v('rel_ind')) in ['Z', 'P']: p_stat, p_rem = 'Closed', 'Blocked item. PO line item closed. No further action is required.'
        elif get_v('dci') == 'X' and get_v('fin') == 'X': p_stat, p_rem = 'Closed', 'PO closed due to final invoice ticked and delivery completed ticked, no further action required.'
        elif get_v('dci') == 'X' : p_stat, p_rem = 'Closed', 'Delivery Completed ticked, no further action required.'
        elif get_v('fin') == 'X': p_stat, p_rem = 'Closed', 'Final invoice ticked, no further action required.'
        elif get_v('rebate') == 'X': p_stat, p_rem = 'Closed', 'Rebate or Return Item, no further action required.'

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

            ven_mat = str(get_v('ven_mat'))
            err_ven = check_special_characters(ven_mat, special_chars)
            if err_ven: log_e('Text', f"Vendor Material Number: {err_ven}", 'ven_mat')

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
            if mat_val == "" and still_pay_qty == 0 and still_pay_amt > 0: log_e('Logic Checks', "Have open amount, but without open quantity (Service)", 'still_pay_amt')
            if mat_val == "" and ir_exist_raw not in ['FOC', 'F.O.C.']: 
                if still_del > 0 and still_pay_qty > 0 and still_pay_amt < 0: log_e('Logic Checks', "Have open amount, but without open quantity (Material)", 'still_pay_amt')
            amt_eur = safe_f(get_v('amt_eur'))
            if 0 < amt_eur <= small_val_limit: log_e('Logic Checks', f"PO value < {small_val_limit} EUR", 'amt_eur')

            # 15. Additional Pricing
            if safe_f(get_v('per')) > 1: log_e('Additional Pricing', "Additional pricing (Per > 1)", 'per')

            # 16. Incoterm
            if not check_mandatory(str(get_v('incot'))): log_e('Incoterm', "Incoterm is missing", 'incot')

        # --- 5. CONSOLIDATE ---
        results['PO Category'].append(p_cat)
        results['PO Status'].append(p_stat)
        results['Remarks'].append(p_rem)

        row_joined_errs = []
        for cat in category_list:
            msg = " | ".join(row_cat_errors[cat])
            results[f"{cat}_Remarks"].append(msg)
            if msg: row_joined_errs.append(msg)
        results['Error_Details'].append(" | ".join(row_joined_errs))

    # --- 6. FINAL BUILD ---
    df_results = pd.DataFrame(results)
    df_results = df_results.reset_index(drop=True)
    df = df.reset_index(drop=True)
    # Concate results to original DF 
    df_out = pd.concat([df_results, df], axis=1)

    return df_out, bad_cells, category_list

def to_excel_po_download(full_df, bad_cells, category_list):
    output = io.BytesIO()

    # Create view of the columns want to keep
    analysis_headers = ['PO Category', 'PO Status', 'Remarks']
    drop_cols = ['Error_Details'] + [f"{c}_Remarks" for c in category_list]
    cols_to_keep = [c for c in full_df.columns if c not in (analysis_headers) + (drop_cols)]
    display_cols = analysis_headers + cols_to_keep
    clean_df = full_df[display_cols].fillna('').astype(str)
    
    writer_options = {'options': {'strings_to_numbers': False, 'constant_memory': True}}
    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs=writer_options) as writer: 
        workbook = writer.book
        
        direct_po_format = workbook.add_format({'bg_color': "#D4D436", 'font_color': "#000000"})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#005eb8', 'font_color': 'white', 'border': 1})
        bold_format = workbook.add_format({'bold': True})

        # define the precision masks
        red_base = {'bg_color': '#FFC7CE', 'font_color': '#9C0006'}
        masks = {
            'text': None, 
            'id': '0', 
            '1dec': '#,##0.0',
            '2dec': '#,##0.00', 
            '3dec': '#,##0.000'
        }
        formats = {}
        for name, mask in masks.items():
            props = {'num_format': mask} if mask else {}
            formats[name] = workbook.add_format(props)
            formats[f"red_{name}"] = workbook.add_format({**props, **red_base})

        # metrics calculation 
        total = len(full_df)
        errors = len(full_df[full_df['Error_Details'] != ""])

        # --- Dashboard ---
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

        # --- Raw Data ---
        ws1 = workbook.add_worksheet('Full_Raw_Data')

        # Precision detection loop
        col_map = {}
        for c_idx, col_name in enumerate(clean_df.columns):
            ws1.write_string(0, c_idx, col_name, header_format)

            # scan first 100 rows of this column to detect the precision style
            sample = clean_df[col_name].astype(str).sample(n=min(len(clean_df), 100))
            precision = 'text'
            for val in sample:
                if val == '' or val.lower() == 'nan': continue
                try:
                    float(val.replace(',', ''))
                    if '.' in val: 
                        decimals = len(val.split('.')[1])
                        if decimals >= 3: precision = '3dec'; break
                        elif decimals == 2: 
                            if precision != '3dec': precision = '2dec'
                        elif decimals == 1:
                            if precision not in ['2dec', '3dec']: precision = '1dec'
                    elif precision not in ['1dec', '2dec', '3dec']: precision = 'id'
                except: 
                    precision = 'text'; break
        
            col_map[c_idx] = precision
            ws1.set_column(c_idx, c_idx, 18)

        # Data loop
        error_lookup = set(bad_cells)

        for r_idx, row in enumerate(clean_df.itertuples(index=False)):
            excel_row = r_idx + 1
            for c_idx, val in enumerate(row):
                col_name = clean_df.columns[c_idx]
                style_key = col_map[c_idx]

                # check if this cell has an error
                is_error = (r_idx, col_name) in error_lookup
                fmt = formats[f"red_{style_key}" if is_error else style_key]

                # Write logic: empty strings or text stay as strings
                if style_key == 'text' or str(val) == '':
                    ws1.write_string(excel_row, c_idx, str(val), fmt)
                else:
                    try:
                        # Convert to number so Excel applies the .000 mask
                        num = float(str(val).replace(',', ''))
                        ws1.write_number(excel_row, c_idx, num, fmt)
                    except:
                        ws1.write_string(excel_row, c_idx, str(val), fmt)
        
        # Direct PO
        try:
            cat_col_idx = clean_df.columns.get_loc('PO Category')
            col_letter = xlsxwriter.utility.xl_col_to_name(cat_col_idx)
            ws1.conditional_format(1, 0, len(clean_df), len(clean_df.columns)-1, {
                'type':     'formula',
                'criteria': f'=${col_letter}2="Direct PO"',
                'format':   direct_po_format
            })
        except: pass

        ws1.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
        ws1.freeze_panes(1, 3)

        # Other sheets
        # Errors Categories Tabs
        for cat in category_list: 
            col_name = f"{cat}_Remarks"
            if col_name in full_df.columns: 
                mask = full_df[col_name].str.strip() != ""
                subset = full_df[mask].copy()
                if not subset.empty: 
                    ws_err = workbook.add_worksheet(cat[:30])
                    # Reorder: Remark first, then raw data
                    err_display = [col_name] + display_cols
                    err_df = subset[err_display].fillna('').astype(str)
                    
                    for c_idx, c_name in enumerate(err_df.columns):
                        ws_err.write_string(0, c_idx, c_name, header_format)
                    for r_idx, row in enumerate(err_df.itertuples(index=False)):
                        for c_idx, val in enumerate(row):
                            if c_idx == 0:
                                ws_err.write_string(r_idx+1, c_idx, str(val))
                            else:
                                style_key = col_map[c_idx - 1]
                                fmt = formats[style_key]

                                if style_key == 'text' or str(val) == '':
                                    ws_err.write_string(r_idx+1, c_idx, str(val), fmt)
                                else:
                                    try:
                                        ws_err.write_number(r_idx+1, c_idx, float(str(val).replace(',', '')), fmt)
                                    except:
                                        ws_err.write_string(r_idx+1, c_idx, str(val), fmt)

                    ws_err.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
                    ws_err.set_column(0, 0, 50)
                    ws_err.set_column(1, len(err_df.columns)-1, 18)

        # Status Tabs
        if 'PO Status' in full_df.columns:
            for status in full_df['PO Status'].unique():
                stat_df = clean_df[clean_df['PO Status'] == status]
                if not stat_df.empty:
                    sheet_name = f"Status_{str(status)[:20]}".replace('/', '_')
                    ws_stat = workbook.add_worksheet(sheet_name)
                    for c_idx, c_name in enumerate(stat_df.columns):
                        ws_stat.write_string(0, c_idx, c_name, header_format)
                    for r_idx, row in enumerate(stat_df.itertuples(index=False)):
                        for c_idx, val in enumerate(row):
                            style_key = col_map[c_idx]
                            fmt = formats[style_key]

                            if style_key == 'text' or str(val) == '':
                                ws_stat.write_string(r_idx+1, c_idx, str(val), fmt)
                            else:
                                try:
                                    ws_stat.write_number(r_idx+1, c_idx, float(str(val).replace(',', '')), fmt)
                                except:
                                    ws_stat.write_string(r_idx+1, c_idx, str(val), fmt)
                    ws_stat.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
                    ws_stat.set_column(0, len(stat_df.columns)-1, 18)
                    ws_stat.freeze_panes(1, 3)

    return output.getvalue()

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# =================================
# USER INTERFACE - Web Display
# =================================

def main():
    # Session state setup
    if 'smd_on' not in st.session_state: st.session_state.smd_on = False
    if 'po_on' not in st.session_state: st.session_state.po_on = False
    if 'email_on' not in st.session_state: st.session_state.email_on = False

    # Callbacks
    def toggle_smd():
        if st.session_state.smd_on:
            st.session_state.po_on = False
            st.session_state.email_on = False

    def toggle_po():
        if st.session_state.po_on:
            st.session_state.smd_on = False
            st.session_state.email_on = False

    def toggle_email():
        if st.session_state.email_on:
            st.session_state.smd_on = False
            st.session_state.po_on = False

    with st.sidebar:
        st.title("🛡️ PROCleans")
        st.caption("Hybrid Rule Engine")
        st.markdown("---")

        # --- 1. Navigation ---
        st.subheader("Select Modules")

        # Switch button to on/off modules
        show_smd = st.toggle("SMD Analysis", key="smd_on", on_change=toggle_smd)
        show_po = st.toggle("PO Analysis", key="po_on", on_change=toggle_po)
        show_email = st.toggle("Email Validation", key="email_on", on_change=toggle_email)

        st.markdown("---")

        if show_smd:
            st.header("Configuration")
            target_cocd = st.text_input("Company Code:", placeholder="e.g., 3072")
            target_porg = st.text_input("Purchase Organization:", placeholder="e.g., 3072, 3050", help="Enter multiple Purchase Organizations separated by commas.")
        
    # ================================================
    # PAGE LOGIC
    # ================================================

    tabs_to_show = []
    if show_smd: tabs_to_show.append("SMD Analysis")
    if show_po: tabs_to_show.append("PO Analysis")
    if show_email: tabs_to_show.append("Email Validation")

    if not tabs_to_show:
        st.title("PROCleans")
        st.markdown("""
                    <div style='background-color: #e6f3ff; padding: 20px; border-radius: 10px; border-left: 5px solid #005eb8;'>
                        <h3>Welcome!
                            Use the sidebar to enable or disable specific analysis modules.</h3>
                        <h4>Available Modules:</h4>
                        <p>- SMD Analysis: Validate Supplier Master Data against global and regional rules.</p>
                        <p>- PO Analysis: Analyze Purchase Orders using a dynamic logic matrix.</p>
                        <p>- Email Validation: Check vendor email lists for missing contacts or format errors.</p>
                    </div>
        """, unsafe_allow_html=True)
        return

    # ==============================================================================================================
    # --- SMD ANALYSIS ---
    # ____________________

    elif show_smd:
        st.title("Supplier Master Data Analysis")

        with st.container(border=True):
            st.subheader("1. Upload Rules Config")
            req_file = st.file_uploader("Upload 'SMD_Rules_Config.xlsx'", type=['xlsx'], key='smd_req')

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

                    results, bad_cells = run_smd_analysis(df, req_file, target_cocd, target_porg)
                    duplicates_df = get_duplicates_df(df)
                    
                    df_errors_only = results[results['Error Details'] != ""]

                    p1 = len(results[results['Purchasing Issues'] != ""])
                    p2 = len(results[results['Org Finance Issues'] != ""])
                    p3 = len(results[results['Master Data Issues'] != ""])
                    gen = len(results[results['General Errors'] != ""])
                
                    metrics = {
                        'Total': len(results), 'Correct': len(results) - len(df_errors_only), 'Errors': len(df_errors_only),
                        'Duplicates': len(duplicates_df),
                        'Purchasing': p1, 'Org Finance': p2, 'Master Data': p3, 'General': gen
                    }
                
                    error_bkdown = pd.DataFrame()
                    if not df_errors_only.empty:
                        error_bkdown = df_errors_only['Error Details'].str.split(' \| ').explode().value_counts().reset_index()
                        error_bkdown.columns = ['Error Description', 'Count']

                    st.metric("Total Errors", len(df_errors_only))
                    fname = f"SMD_Report_{target_cocd}.xlsx"
                    data = to_excel_download_smd(results, df_errors_only, duplicates_df, metrics, error_bkdown, bad_cells)
                    st.download_button("Download Report", data, fname)
    # ==============================================================================================================

    # ==============================================================================================================
    # --- PO ANALYSIS ----
    # _____________________

    elif show_po:
        st.title("Purchase Order Analysis")

        with st.container(border=True):
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
                            sheet_df = pd.read_excel(po_raw_file, sheet_name=sheet_name, keep_default_na=False, na_values=None)
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
    # ==============================================================================================================

    # ==============================================================================================================
    # --- EMAIL ---
    # ________________

    elif show_email:
        st.title("Vendor Email Validation")

        with st.container(border=True):
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
                    
    # ==============================================================================================================

if __name__ == "__main__":
    main()

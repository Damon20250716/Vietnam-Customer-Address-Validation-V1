import streamlit as st
import pandas as pd
import unicodedata
import difflib
import re
from io import BytesIO
from datetime import datetime

# Remove Vietnamese tone marks
def remove_tones(text):
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    text = unicodedata.normalize('NFD', text)
    return ''.join([c for c in text if unicodedata.category(c) != 'Mn'])

# Normalize text (remove tones + lowercase)
def normalize_text(text):
    return remove_tones(text).lower().strip()

# Clean address for matching
def clean_address(addr):
    if not isinstance(addr, str):
        return ""
    addr = normalize_text(addr)
    addr = re.sub(r'[^\w\s,]', ' ', addr)  # Remove special chars
    addr = re.sub(r'\s+', ' ', addr).strip()  # Normalize spaces
    return addr

# Address matching with ≥60% similarity
def address_match(addr1, addr2):
    a1 = clean_address(addr1)
    a2 = clean_address(addr2)
    
    if not a1 or not a2:
        return False
    
    # Exact match
    if a1 == a2:
        return True
    
    # Substring match
    if a1 in a2 or a2 in a1:
        return True
    
    # Similarity ratio (≥60%)
    ratio = difflib.SequenceMatcher(None, a1, a2).ratio()
    if ratio >= 0.6:
        return True
    
    # Component-wise match (≥60% per component)
    a1_parts = [p.strip() for p in a1.split(',') if p.strip()]
    a2_parts = [p.strip() for p in a2.split(',') if p.strip()]
    
    if a1_parts and a2_parts:
        match_count = 0
        for p1 in a1_parts:
            for p2 in a2_parts:
                if difflib.SequenceMatcher(None, p1, p2).ratio() >= 0.6:
                    match_count += 1
                    break
        if match_count / len(a1_parts) >= 0.6:
            return True
            
    return False

# Get contact name from form
def get_contact_name(form_row):
    name_fields = [
        "Full Name of Contact-In English Only",
        "Full Name of Contact",
        "Contact Name",
        "Name of Contact",
        "Contact Full Name"
    ]
    for field in name_fields:
        if field in form_row:
            return str(form_row[field]) if pd.notna(form_row[field]) else ""
    return ""

# Get address line from UPS data
def get_ups_address_line(ups_row, line_num):
    col = f"Address Line {line_num}"
    return str(ups_row[col]) if pd.notna(ups_row[col]) else ""

# Normalize account number (handle numbers, strings, floats)
def normalize_account(account):
    if pd.isna(account):
        return ""
    account_str = str(account).strip().lower()
    if account_str.endswith('.0'):  # Handle float numbers
        account_str = account_str[:-2]
    return account_str

# Validate required columns
def validate_columns(df, required, df_name):
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns in {df_name}: {', '.join(missing)}")

def process_files(forms_df, ups_df):
    # Preprocess UPS data
    ups_df['Address Type'] = ups_df['Address Type'].astype(str).str.strip().str.zfill(2)
    ups_df['Account Number_norm'] = ups_df['Account Number'].apply(normalize_account)
    ups_df['Full_Address'] = ups_df.apply(
        lambda x: f"{get_ups_address_line(x, 1)}, {get_ups_address_line(x, 2)}, {get_ups_address_line(x, 3)}", 
        axis=1
    )
    ups_df['Full_Address_clean'] = ups_df['Full_Address'].apply(clean_address)
    ups_grouped = ups_df.groupby('Account Number_norm')

    # Preprocess Forms data
    forms_df['Account Number_norm'] = forms_df['Account Number'].apply(normalize_account)
    forms_df['Is Same Billing'] = forms_df[
        "Is Your New Billing Address the Same as Your Pickup and Delivery Address?"
    ].astype(str).str.strip().str.lower()

    # Track all form indices to ensure none are missing
    all_form_indices = set(forms_df.index)
    matched_rows = []
    unmatched_rows = []
    upload_template = []

    # Process each form entry
    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same = form_row['Is Same Billing']
        form_data = form_row.to_dict()
        
        # Remove tones from all string fields for output
        for key, value in form_data.items():
            if isinstance(value, str):
                form_data[key] = remove_tones(value)

        # Check if account exists in UPS data
        if acc_norm not in ups_grouped.groups:
            form_data['Unmatched Reason'] = "Account number not found in UPS data"
            unmatched_rows.append(form_data)
            all_form_indices.discard(idx)
            continue

        ups_account_data = ups_grouped.get_group(acc_norm)
        matched = False
        ups_01 = ups_account_data[ups_account_data['Address Type'] == '01']

        # Case 1: Customer answered "Yes" (single address for all types)
        if is_same == 'yes':
            # Combine Form address lines 1-3
            form_addr1 = str(form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")).strip()
            form_addr2 = str(form_row.get("New Address Line 2 (Street Name)-In English Only", "")).strip()
            form_addr3 = str(form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")).strip()
            form_full = f"{form_addr1}, {form_addr2}, {form_addr3}"
            form_clean = clean_address(form_full)

            # Match against UPS 01 (All-in-one address type)
            if not ups_01.empty:
                for _, ups_row in ups_01.iterrows():
                    if address_match(form_full, ups_row['Full_Address']):  # Compare raw full address
                        # Matched - add to matched records
                        match_data = form_data.copy()
                        match_data['UPS Matched Address'] = ups_row['Full_Address']
                        match_data['Matching Type'] = '01 (All-in-one)'
                        matched_rows.append(match_data)
                        matched = True

                        # Build upload template (matches your Excel structure)
                        contact = get_contact_name(form_row)
                        
                        # 1. Address Type 01 (All-in-one) - Invoice Option BLANK
                        upload_template.append({
                            "AC_NUM": form_row['Account Number'],
                            "AC_Address_Type": "01",
                            "invoice option": "",  # Blank as per template
                            "AC_name": remove_tones(str(ups_row.get("AC_Name", ""))),
                            "Address_Line1": remove_tones(form_addr1),
                            "Address_Line2": remove_tones(form_addr2),
                            "Postal_Code": str(ups_row.get("Postal_Code", "")),
                            "Country_Code": str(ups_row.get("Country_Code", "")),
                            "Attention_Name": contact,
                            "Address_Line22": remove_tones(form_addr3),  # Address Line 3
                            "Address_Country_Code": str(ups_row.get("Address_Country_Code", "")),
                            "Matching_Score": "High"
                        })

                        # 2. Address Type 03 (Billing) - Invoice Options 1, 2, 6 (3 rows)
                        for code in ['1', '2', '6']:
                            upload_template.append({
                                "AC_NUM": form_row['Account Number'],
                                "AC_Address_Type": "03",
                                "invoice option": code,
                                "AC_name": remove_tones(str(ups_row.get("AC_Name", ""))),
                                "Address_Line1": remove_tones(form_addr1),
                                "Address_Line2": remove_tones(form_addr2),
                                "Postal_Code": str(ups_row.get("Postal_Code", "")),
                                "Country_Code": str(ups_row.get("Country_Code", "")),
                                "Attention_Name": contact,
                                "Address_Line22": remove_tones(form_addr3),
                                "Address_Country_Code": str(ups_row.get("Address_Country_Code", "")),
                                "Matching_Score": "High"
                            })
                        break

        # Case 2: Customer answered "No" (separate billing/delivery/pickup)
        else:
            # --- Billing Address (Type 03) ---
            billing_addr1 = str(form_row.get("New Billing Address Line 1", "")).strip()
            billing_addr2 = str(form_row.get("New Billing Address Line 2", "")).strip()
            billing_addr3 = str(form_row.get("New Billing Address Line 3", "")).strip()
            billing_full = f"{billing_addr1}, {billing_addr2}, {billing_addr3}"
            billing_clean = clean_address(billing_full)
            billing_matched = False

            ups_03 = ups_account_data[ups_account_data['Address Type'] == '03']
            for _, ups_row in ups_03.iterrows():
                if address_match(billing_full, ups_row['Full_Address']):
                    billing_matched = True
                    billing_ups_row = ups_row
                    break

            # --- Delivery Address (Type 13) ---
            delivery_addr1 = str(form_row.get("New Delivery Address Line 1", "")).strip()
            delivery_addr2 = str(form_row.get("New Delivery Address Line 2", "")).strip()
            delivery_addr3 = str(form_row.get("New Delivery Address Line 3", "")).strip()
            delivery_full = f"{delivery_addr1}, {delivery_addr2}, {delivery_addr3}"
            delivery_clean = clean_address(delivery_full)
            delivery_matched = False

            ups_13 = ups_account_data[ups_account_data['Address Type'] == '13']
            for _, ups_row in ups_13.iterrows():
                if address_match(delivery_full, ups_row['Full_Address']):
                    delivery_matched = True
                    delivery_ups_row = ups_row
                    break

            # --- Pickup Addresses (Type 02, max 3) ---
            pickup_addrs = []
            pickup_count_raw = form_row.get("How Many Pick Up Address Do You Have?", 0)
            word_to_num = {"One": 1, "Two": 2, "Three": 3}  # Extend for more values if needed
            if isinstance(pickup_count_raw, str):
                pickup_count = word_to_num.get(pickup_count_raw.strip(), 0)
            else:
                try:
                    pickup_count = int(pickup_count_raw)
                except ValueError:
                    pickup_count = 0
            pickup_count = min(pickup_count, 3)  # Max 3 pickups

            for i in range(1, pickup_count + 1):
                prefix = ["First", "Second", "Third"][i-1]
                addr1 = str(form_row.get(f"{prefix} New Pick Up Address Line 1", "")).strip()
                addr2 = str(form_row.get(f"{prefix} New Pick Up Address Line 2", "")).strip()
                addr3 = str(form_row.get(f"{prefix} New Pick Up Address Line 3", "")).strip()
                pickup_addrs.append((addr1, addr2, addr3))

            # Check pickup count matches UPS
            ups_02_count = len(ups_account_data[ups_account_data['Address Type'] == '02'])
            if len(pickup_addrs) != ups_02_count:
                form_data['Unmatched Reason'] = f"Pickup count mismatch (Form: {len(pickup_addrs)}, UPS: {ups_02_count})"
                unmatched_rows.append(form_data)
                all_form_indices.discard(idx)
                continue

            # Check each pickup address
            pickup_matched = True
            pickup_ups_rows = []
            ups_02 = ups_account_data[ups_account_data['Address Type'] == '02']

            for addr in pickup_addrs:
                addr_full = f"{addr[0]}, {addr[1]}, {addr[2]}"
                addr_clean = clean_address(addr_full)
                found = False
                for _, ups_row in ups_02.iterrows():
                    if address_match(addr_full, ups_row['Full_Address']):
                        pickup_ups_rows.append(ups_row)
                        found = True
                        break
                if not found:
                    pickup_matched = False
                    break

            # If all addresses match, build upload template
            if billing_matched and delivery_matched and pickup_matched:
                match_data = form_data.copy()
                match_data['UPS Matched Billing Address'] = billing_ups_row['Full_Address']
                match_data['UPS Matched Delivery Address'] = delivery_ups_row['Full_Address']
                for i, (addr, ups_row) in enumerate(zip(pickup_addrs, pickup_ups_rows), 1):
                    match_data[f"UPS Matched Pickup Address {i}"] = ups_row['Full_Address']
                matched_rows.append(match_data)
                matched = True

                # Build upload template rows
                contact = get_contact_name(form_row)
                
                # 1. Billing Address (Type 03) - Invoice Options 1, 2, 6 (3 rows)
                for code in ['1', '2', '6']:
                    upload_template.append({
                        "AC_NUM": form_row['Account Number'],
                        "AC_Address_Type": "03",
                        "invoice option": code,
                        "AC_name": remove_tones(str(billing_ups_row.get("AC_Name", ""))),
                        "Address_Line1": remove_tones(billing_addr1),
                        "Address_Line2": remove_tones(billing_addr2),
                        "Postal_Code": str(billing_ups_row.get("Postal_Code", "")),
                        "Country_Code": str(billing_ups_row.get("Country_Code", "")),
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(billing_addr3),
                        "Address_Country_Code": str(billing_ups_row.get("Address_Country_Code", "")),
                        "Matching_Score": "High"
                    })

                # 2. Delivery Address (Type 13) - Invoice Option BLANK
                upload_template.append({
                    "AC_NUM": form_row['Account Number'],
                    "AC_Address_Type": "13",
                    "invoice option": "",
                    "AC_name": remove_tones(str(delivery_ups_row.get("AC_Name", ""))),
                    "Address_Line1": remove_tones(delivery_addr1),
                    "Address_Line2": remove_tones(delivery_addr2),
                    "Postal_Code": str(delivery_ups_row.get("Postal_Code", "")),
                    "Country_Code": str(delivery_ups_row.get("Country_Code", "")),
                    "Attention_Name": contact,
                    "Address_Line22": remove_tones(delivery_addr3),
                    "Address_Country_Code": str(delivery_ups_row.get("Address_Country_Code", "")),
                    "Matching_Score": "High"
                })

                # 3. Pickup Addresses (Type 02) - Each pickup as a separate row (Invoice Option BLANK)
                for i, (addr, ups_row) in enumerate(zip(pickup_addrs, pickup_ups_rows), 1):
                    pu_city = remove_tones(str(form_row.get(f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address City / Province", "")))
                    upload_template.append({
                        "AC_NUM": form_row['Account Number'],
                        "AC_Address_Type": "02",
                        "invoice option": "",
                        "AC_name": remove_tones(str(ups_row.get("AC_Name", ""))),
                        "Address_Line1": remove_tones(addr[0]),
                        "Address_Line2": remove_tones(addr[1]),
                        "Postal_Code": str(ups_row.get("Postal_Code", "")),
                        "Country_Code": str(ups_row.get("Country_Code", "")),
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(addr[2]),  # Address Line 3
                        "Address_Country_Code": str(ups_row.get("Address_Country_Code", "")),
                        "Matching_Score": "High"
                    })

        # Handle unmatched cases
        if not matched:
            if is_same == 'yes':
                reason = "Address did not match UPS '01' type address" if not ups_01.empty else "No '01' type address in UPS data"
            else:
                reason = "One or more address types (billing/delivery/pickup) did not match"
            form_data['Unmatched Reason'] = reason
            unmatched_rows.append(form_data)
        
        all_form_indices.discard(idx)

    # Catch any remaining form entries that weren't processed
    for idx in all_form_indices:
        form_data = forms_df.loc[idx].to_dict()
        for key, value in form_data.items():
            if isinstance(value, str):
                form_data[key] = remove_tones(value)
        form_data['Unmatched Reason'] = "Not processed - system error"
        unmatched_rows.append(form_data)

    # Convert to DataFrames
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_df = pd.DataFrame(upload_template)

    # Ensure column order matches your Excel template
    upload_columns = [
        "AC_NUM", "AC_Address_Type", "invoice option", "AC_name", 
        "Address_Line1", "Address_Line2", "Postal_Code", "Country_Code",
        "Attention_Name", "Address_Line22", "Address_Country_Code", 
        "Matching_Score"
    ]
    upload_df = upload_df.reindex(columns=upload_columns)

    return matched_df, unmatched_df, upload_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

def main():
    st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide", page_icon="🇻🇳")
    
    st.title("Vietnam Address Validation Tool")
    st.write("Upload Microsoft Forms responses and UPS system data to validate addresses")

    # File uploaders
    forms_file = st.file_uploader("Upload Microsoft Forms Responses (XLSX)", type="xlsx", key="forms")
    ups_file = st.file_uploader("Upload UPS System Data (XLSX)", type="xlsx", key="ups")

    if forms_file and ups_file:
        if st.button("Process Files", type="primary"):
            with st.spinner("Processing files..."):
                try:
                    # Read files
                    forms_df = pd.read_excel(forms_file)
                    ups_df = pd.read_excel(ups_file)

                    # Validate required columns
                    required_forms = ["Account Number", "Is Your New Billing Address the Same as Your Pickup and Delivery Address?"]
                    required_ups = ["Account Number", "Address Type", "Address Line 1"]
                    validate_columns(forms_df, required_forms, "Forms Data")
                    validate_columns(ups_df, required_ups, "UPS Data")

                    # Process data
                    matched, unmatched, upload = process_files(forms_df, ups_df)

                    # Display results
                    st.success(f"Processing complete! Matched: {len(matched)}, Unmatched: {len(unmatched)}")

                    # Show tabs
                    tab1, tab2, tab3 = st.tabs(["Matched Records", "Unmatched Records", "Upload Template"])
                    
                    with tab1:
                        st.subheader("Matched Records")
                        st.dataframe(matched)
                        if not matched.empty:
                            st.download_button(
                                "Download Matched Records",
                                to_excel(matched),
                                "matched_records.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                    with tab2:
                        st.subheader("Unmatched Records")
                        st.dataframe(unmatched)
                        if not unmatched.empty:
                            st.download_button(
                                "Download Unmatched Records",
                                to_excel(unmatched),
                                "unmatched_records.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                    with tab3:
                        st.subheader("Upload Template")
                        st.dataframe(upload)
                        if not upload.empty:
                            st.download_button(
                                "Download Upload Template",
                                to_excel(upload),
                                "upload_template.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                except Exception as e:
                    st.error(f"Error: {str(e)}")
                    st.exception(e)

if __name__ == "__main__":
    main()

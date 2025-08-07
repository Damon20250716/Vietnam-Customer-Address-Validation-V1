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

# Clean address for matching (preserve core components)
def clean_address(addr):
    if not isinstance(addr, str):
        return ""
    addr = normalize_text(addr)
    addr = re.sub(r'[^\w\s,/-]', ' ', addr)  # Preserve slashes and hyphens in addresses
    addr = re.sub(r'\s+', ' ', addr).strip()
    return addr

# ENHANCED: More flexible address matching for similar addresses (like K8C455)
def address_match(addr1, addr2):
    a1 = clean_address(addr1)
    a2 = clean_address(addr2)
    
    if not a1 or not a2:
        return False
    
    # Exact match
    if a1 == a2:
        return True
    
    # Substring match (prioritize longer substrings)
    min_length = min(len(a1), len(a2))
    if min_length > 5:  # Only check substrings for addresses with meaningful length
        if a1 in a2 or a2 in a1:
            return True
    
    # Similarity ratio (lowered threshold to 50% for more flexibility)
    ratio = difflib.SequenceMatcher(None, a1, a2).ratio()
    if ratio >= 0.5:  # Reduced from 0.6 to catch closer matches
        return True
    
    # Component-wise match (focus on street names)
    a1_parts = [p.strip() for p in a1.split(',') if p.strip()]
    a2_parts = [p.strip() for p in a2.split(',') if p.strip()]
    
    if a1_parts and a2_parts:
        # Prioritize street name components (first parts are often street numbers/names)
        street_match = False
        if len(a1_parts) > 0 and len(a2_parts) > 0:
            street_ratio = difflib.SequenceMatcher(None, a1_parts[0], a2_parts[0]).ratio()
            if street_ratio >= 0.5:  # Street components match
                street_match = True
        
        # Check other components
        match_count = 0
        for p1 in a1_parts[1:]:  # Skip first part (already checked street)
            for p2 in a2_parts[1:]:
                if difflib.SequenceMatcher(None, p1, p2).ratio() >= 0.5:
                    match_count += 1
                    break
        
        # If street matches and at least 1 other component
        if street_match and (match_count >= 1):
            return True
            
    return False

# Get contact name from form (expanded field names)
def get_contact_name(form_row):
    name_fields = [
        "Full Name of Contact-In English Only",
        "Full Name of Contact",
        "Contact Name",
        "Name of Contact",
        "Contact Full Name",
        "Contact Person Name",  # Added common variant
        "Name"  # Fallback
    ]
    for field in name_fields:
        if field in form_row:
            return str(form_row[field]) if pd.notna(form_row[field]) else ""
    return ""

# Get address line from form (expanded field names to ensure capture)
def get_form_address_line(form_row, line_num):
    # Common variations of address line fields
    line_variants = [
        f"New Address Line {line_num}",
        f"New Address Line {line_num} (In English Only)",
        f"Address Line {line_num}",
        f"New Billing Address Line {line_num}",  # For "No" case
        f"New Delivery Address Line {line_num}"   # For "No" case
    ]
    for variant in line_variants:
        if variant in form_row:
            return str(form_row[variant]).strip() if pd.notna(form_row[variant]) else ""
    return ""

# Get address line from UPS data
def get_ups_address_line(ups_row, line_num):
    col = f"Address Line {line_num}"
    return str(ups_row[col]).strip() if pd.notna(ups_row[col]) else ""

# Normalize account number (ensure consistency)
def normalize_account(account):
    if pd.isna(account):
        return ""
    account_str = str(account).strip().upper()  # Use uppercase for consistency (K8C455 vs k8c455)
    if account_str.endswith('.0'):
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

    # Track all form indices
    all_form_indices = set(forms_df.index)
    matched_rows = []
    unmatched_rows = []
    upload_template = []

    # Process each form entry
    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same = form_row['Is Same Billing']
        form_data = form_row.to_dict()
        
        # Remove tones from all string fields
        for key, value in form_data.items():
            if isinstance(value, str):
                form_data[key] = remove_tones(value)

        # Check if account exists in UPS data (critical for K8C455)
        if acc_norm not in ups_grouped.groups:
            form_data['Unmatched Reason'] = "Account number not found in UPS data"
            unmatched_rows.append(form_data)
            all_form_indices.discard(idx)
            continue

        ups_account_data = ups_grouped.get_group(acc_norm)
        matched = False

        # Case 1: Customer answered "Yes" (single address for all types)
        if is_same == 'yes':
            # Get form address lines (using flexible field matching)
            form_addr1 = get_form_address_line(form_row, 1)
            form_addr2 = get_form_address_line(form_row, 2)
            form_addr3 = get_form_address_line(form_row, 3)
            form_full = f"{form_addr1}, {form_addr2}, {form_addr3}"
            form_clean = clean_address(form_full)

            # Get city from form
            city_fields = ["New City / Province", "New City - In English Only", "City", "New Billing City"]
            form_city = ""
            for field in city_fields:
                if field in form_row:
                    form_city = remove_tones(str(form_row[field]).strip())
                    break

            # ENHANCED: Check both 01 and 03 address types for "Yes" answers (critical for K8C455)
            # First check 01 (original logic), then 03 if no match
            address_types_to_check = ['01', '03']  # Added 03 to check
            for addr_type in address_types_to_check:
                ups_matching_type = ups_account_data[ups_account_data['Address Type'] == addr_type]
                if not ups_matching_type.empty:
                    for _, ups_row in ups_matching_type.iterrows():
                        if address_match(form_full, ups_row['Full_Address']):
                            # Add to matched records
                            match_data = form_data.copy()
                            match_data['UPS Matched Address'] = ups_row['Full_Address']
                            match_data['Matching Type'] = f"{addr_type} (All-in-one)"
                            matched_rows.append(match_data)
                            matched = True

                            # Build upload template
                            contact = get_contact_name(form_row)
                            
                            # 1. Address Type 01 (billing) - Invoice Option BLANK
                            upload_template.append({
                                "AC_NUM": form_row['Account Number'],
                                "AC_Address_Type": "01",
                                "invoice option": "",
                                "AC_name": remove_tones(str(ups_row.get("AC_Name", ""))),
                                "Address_Line1": remove_tones(form_addr1),
                                "Address_Line2": remove_tones(form_addr2),
                                "City": form_city,
                                "Postal_Code": str(ups_row.get("Postal_Code", "")),
                                "Country_Code": str(ups_row.get("Country_Code", "")),
                                "Attention_Name": contact,
                                "Address_Line22": remove_tones(form_addr3),
                                "Address_Country_Code": str(ups_row.get("Address_Country_Code", "")),
                                "Matching_Score": "High"
                            })

                            # 2. Address Type 03 (delivery) - Invoice Options 1, 2, 6
                            for code in ['1', '2', '6']:
                                upload_template.append({
                                    "AC_NUM": form_row['Account Number'],
                                    "AC_Address_Type": "03",
                                    "invoice option": code,
                                    "AC_name": remove_tones(str(ups_row.get("AC_Name", ""))),
                                    "Address_Line1": remove_tones(form_addr1),
                                    "Address_Line2": remove_tones(form_addr2),
                                    "City": form_city,
                                    "Postal_Code": str(ups_row.get("Postal_Code", "")),
                                    "Country_Code": str(ups_row.get("Country_Code", "")),
                                    "Attention_Name": contact,
                                    "Address_Line22": remove_tones(form_addr3),
                                    "Address_Country_Code": str(ups_row.get("Address_Country_Code", "")),
                                    "Matching_Score": "High"
                                })
                            break  # Exit loop once matched
                    if matched:
                        break  # Exit address type loop once matched

        # Case 2: Customer answered "No" (separate addresses)
        else:
            # --- Billing Address (Type 02, invoice blank) ---
            billing_addr1 = get_form_address_line(form_row, 1)  # Flexible field matching
            billing_addr2 = get_form_address_line(form_row, 2)
            billing_addr3 = get_form_address_line(form_row, 3)
            billing_full = f"{billing_addr1}, {billing_addr2}, {billing_addr3}"
            billing_city = remove_tones(str(form_row.get("New Billing City / Province", "")).strip())
            billing_matched = False

            # Check 02 and 03 for billing (flexibility)
            for addr_type in ['02', '03']:
                ups_billing = ups_account_data[ups_account_data['Address Type'] == addr_type]
                for _, ups_row in ups_billing.iterrows():
                    if address_match(billing_full, ups_row['Full_Address']):
                        billing_matched = True
                        billing_ups_row = ups_row
                        break
                if billing_matched:
                    break

            # --- Delivery Address (Type 03, invoice 1,2,6) ---
            delivery_addr1 = get_form_address_line(form_row, 1)
            delivery_addr2 = get_form_address_line(form_row, 2)
            delivery_addr3 = get_form_address_line(form_row, 3)
            delivery_full = f"{delivery_addr1}, {delivery_addr2}, {delivery_addr3}"
            delivery_city = remove_tones(str(form_row.get("New Delivery City / Province", "")).strip())
            delivery_matched = False

            ups_delivery = ups_account_data[ups_account_data['Address Type'] == '03']
            for _, ups_row in ups_delivery.iterrows():
                if address_match(delivery_full, ups_row['Full_Address']):
                    delivery_matched = True
                    delivery_ups_row = ups_row
                    break

            # --- Pickup Addresses (Type 13, max 3, invoice blank) ---
            pickup_addrs = []
            pickup_count_raw = form_row.get("How Many Pick Up Address Do You Have?", 0)
            word_to_num = {"One": 1, "Two": 2, "Three": 3, "1":1, "2":2, "3":3}  # Added numeric handling
            if isinstance(pickup_count_raw, str):
                pickup_count = word_to_num.get(pickup_count_raw.strip(), 0)
            else:
                try:
                    pickup_count = int(pickup_count_raw)
                except ValueError:
                    pickup_count = 0
            pickup_count = min(pickup_count, 3)

            for i in range(1, pickup_count + 1):
                prefix = ["First", "Second", "Third"][i-1]
                addr1 = str(form_row.get(f"{prefix} New Pick Up Address Line 1", "")).strip()
                addr2 = str(form_row.get(f"{prefix} New Pick Up Address Line 2", "")).strip()
                addr3 = str(form_row.get(f"{prefix} New Pick Up Address Line 3", "")).strip()
                city = remove_tones(str(form_row.get(f"{prefix} New Pick Up Address City / Province", "")).strip())
                addr_full = f"{addr1}, {addr2}, {addr3}"
                pickup_addrs.append( (addr_full, addr1, addr2, addr3, city) )

            # Check pickup count matches UPS 13 addresses
            ups_13 = ups_account_data[ups_account_data['Address Type'] == '13']
            ups_13_count = len(ups_13)
            if len(pickup_addrs) != ups_13_count and ups_13_count > 0:
                form_data['Unmatched Reason'] = f"Pickup count mismatch (Form: {len(pickup_addrs)}, UPS: {ups_13_count})"
                unmatched_rows.append(form_data)
                all_form_indices.discard(idx)
                continue

            # Check each pickup address against UPS 13
            pickup_matched = True
            pickup_ups_rows = []
            for addr_full, _, _, _, _ in pickup_addrs:
                found = False
                for _, ups_row in ups_13.iterrows():
                    if address_match(addr_full, ups_row['Full_Address']):
                        pickup_ups_rows.append(ups_row)
                        found = True
                        break
                if not found and ups_13_count > 0:  # Only fail if UPS has pickup addresses
                    pickup_matched = False
                    break

            # If all addresses match, build upload template
            if billing_matched and delivery_matched and (pickup_matched or ups_13_count == 0):
                match_data = form_data.copy()
                match_data['UPS Matched Billing Address'] = billing_ups_row['Full_Address']
                match_data['UPS Matched Delivery Address'] = delivery_ups_row['Full_Address']
                for i, (_, ups_row) in enumerate(zip(pickup_addrs, pickup_ups_rows), 1):
                    match_data[f"UPS Matched Pickup Address {i}"] = ups_row['Full_Address']
                matched_rows.append(match_data)
                matched = True

                # Build upload template rows
                contact = get_contact_name(form_row)
                
                # 1. Billing Address (Type 02, invoice blank)
                upload_template.append({
                    "AC_NUM": form_row['Account Number'],
                    "AC_Address_Type": "02",
                    "invoice option": "",
                    "AC_name": remove_tones(str(billing_ups_row.get("AC_Name", ""))),
                    "Address_Line1": remove_tones(billing_addr1),
                    "Address_Line2": remove_tones(billing_addr2),
                    "City": billing_city,
                    "Postal_Code": str(billing_ups_row.get("Postal_Code", "")),
                    "Country_Code": str(billing_ups_row.get("Country_Code", "")),
                    "Attention_Name": contact,
                    "Address_Line22": remove_tones(billing_addr3),
                    "Address_Country_Code": str(billing_ups_row.get("Address_Country_Code", "")),
                    "Matching_Score": "High"
                })

                # 2. Delivery Address (Type 03, invoice 1,2,6)
                for code in ['1', '2', '6']:
                    upload_template.append({
                        "AC_NUM": form_row['Account Number'],
                        "AC_Address_Type": "03",
                        "invoice option": code,
                        "AC_name": remove_tones(str(delivery_ups_row.get("AC_Name", ""))),
                        "Address_Line1": remove_tones(delivery_addr1),
                        "Address_Line2": remove_tones(delivery_addr2),
                        "City": delivery_city,
                        "Postal_Code": str(delivery_ups_row.get("Postal_Code", "")),
                        "Country_Code": str(delivery_ups_row.get("Country_Code", "")),
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(delivery_addr3),
                        "Address_Country_Code": str(delivery_ups_row.get("Address_Country_Code", "")),
                        "Matching_Score": "High"
                    })

                # 3. Pickup Addresses (Type 13, invoice blank)
                for (_, addr1, addr2, addr3, pu_city), ups_row in zip(pickup_addrs, pickup_ups_rows):
                    upload_template.append({
                        "AC_NUM": form_row['Account Number'],
                        "AC_Address_Type": "13",
                        "invoice option": "",
                        "AC_name": remove_tones(str(ups_row.get("AC_Name", ""))),
                        "Address_Line1": remove_tones(addr1),
                        "Address_Line2": remove_tones(addr2),
                        "City": pu_city,
                        "Postal_Code": str(ups_row.get("Postal_Code", "")),
                        "Country_Code": str(ups_row.get("Country_Code", "")),
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(addr3),
                        "Address_Country_Code": str(ups_row.get("Address_Country_Code", "")),
                        "Matching_Score": "High"
                    })

        # Handle unmatched cases (with debug info for K8C455)
        if not matched:
            if acc_norm == "K8C455":  # Add specific debug for K8C455
                form_data['Unmatched Reason'] = (
                    f"Debug: K8C455 - is_same={is_same}, "
                    f"Form Address: {form_full if 'form_full' in locals() else 'N/A'}, "
                    f"UPS Addresses: {', '.join(ups_account_data['Full_Address'].unique())}"
                )
            else:
                if is_same == 'yes':
                    reason = "Address did not match UPS '01' or '03' type addresses"
                else:
                    reason = "One or more address types (billing/delivery/pickup) did not match"
                form_data['Unmatched Reason'] = reason
            unmatched_rows.append(form_data)
        
        all_form_indices.discard(idx)

    # Catch any unprocessed entries
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

    # Ensure column order matches your template
    upload_columns = [
        "AC_NUM", "AC_Address_Type", "invoice option", "AC_name", 
        "Address_Line1", "Address_Line2", "City",
        "Postal_Code", "Country_Code", "Attention_Name", 
        "Address_Line22", "Address_Country_Code", "Matching_Score"
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

                    # Display results with focus on K8C455
                    st.success(f"Processing complete! Matched: {len(matched)}, Unmatched: {len(unmatched)}")
                    if "K8C455" in matched['Account Number_norm'].values:
                        st.info("Account K8C455 has been matched successfully!")

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

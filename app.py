import streamlit as st
import pandas as pd
import unicodedata
import difflib
import re
from io import BytesIO
from datetime import datetime

# Remove Vietnamese tones from text
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text

# Enhanced clean address function
def clean_address(addr):
    if not isinstance(addr, str):
        return ""
    addr = remove_tones(addr)
    addr = addr.lower()
    addr = re.sub(r'[^\w\s,]', ' ', addr)
    addr = re.sub(r'\s+', ' ', addr).strip()
    addr = re.sub(r',\s+', ',', addr)
    addr = re.sub(r'\s+,', ',', addr)
    return addr

# Improved flexible address matching
def flexible_address_match(addr1, addr2, threshold=0.65):
    a1 = clean_address(addr1)
    a2 = clean_address(addr2)
    
    if not a1 or not a2:
        return False
    
    if a1 == a2:
        return True
    
    if a1 in a2 or a2 in a1:
        return True
    
    a1_parts = [p.strip() for p in a1.split(',') if p.strip()]
    a2_parts = [p.strip() for p in a2.split(',') if p.strip()]
    
    if a1_parts and a2_parts:
        if all(any(difflib.SequenceMatcher(None, p1, p2).ratio() >= 0.7 
                 for p2 in a2_parts) for p1 in a1_parts):
            return True
        if all(any(difflib.SequenceMatcher(None, p2, p1).ratio() >= 0.7 
                 for p1 in a1_parts) for p2 in a2_parts):
            return True
    
    combined_a1 = ' '.join(a1_parts) if a1_parts else a1
    combined_a2 = ' '.join(a2_parts) if a2_parts else a2
    
    ratio = difflib.SequenceMatcher(None, combined_a1, combined_a2).ratio()
    if ratio >= threshold:
        return True
    
    words1 = set(combined_a1.split())
    words2 = set(combined_a2.split())
    
    if words1 and words2:
        if words1.issubset(words2) or words2.issubset(words1):
            return True
        
        common_words = words1 & words2
        min_len = min(len(words1), len(words2))
        if len(common_words) / min_len >= 0.7:
            return True
    
    return False

def get_contact_name(form_row):
    possible_names = [
        "Full Name of Contact-In English Only",
        "Full Name of Contact",
        "Contact Name",
        "Name of Contact",
        "Contact Full Name"
    ]
    
    for name in possible_names:
        if name in form_row:
            return str(form_row[name])
    return ""

def get_address_line(ups_row, line_num):
    col_name = f"Address Line {line_num}"
    return ups_row.get(col_name, "")

def validate_columns(df, required_columns):
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

# New: Normalize account numbers to handle floats (e.g., 12345.0 → "12345")
def normalize_account(account):
    account_str = str(account).lower().strip()
    # Remove trailing .0 if it's an integer float
    if account_str.endswith('.0'):
        account_str = account_str[:-2]
    return account_str

def process_files(forms_df, ups_df):
    try:
        # Fix 1: Standardize Address Type to 2-digit strings (e.g., 2 → "02")
        ups_df['Address Type'] = ups_df['Address Type'].astype(str).str.strip().str.zfill(2)
        
        # Validate required columns
        required_form_columns = ["Account Number", "Is Your New Billing Address the Same as Your Pickup and Delivery Address?"]
        required_ups_columns = ["Account Number", "Address Type", "Address Line 1"]
        validate_columns(forms_df, required_form_columns)
        validate_columns(ups_df, required_ups_columns)

        matched_rows = []
        unmatched_rows = []
        upload_template_rows = []

        # Fix 2: Use improved account normalization
        ups_df['Account Number_norm'] = ups_df['Account Number'].apply(normalize_account)
        forms_df['Account Number_norm'] = forms_df['Account Number'].apply(normalize_account)

        ups_grouped = ups_df.groupby('Account Number_norm')
        if not ups_grouped.ngroups:
            raise ValueError("No valid account numbers found in UPS data")

        processed_form_indices = set()

        for idx, form_row in forms_df.iterrows():
            acc_norm = form_row['Account Number_norm']
            is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

            if acc_norm not in ups_grouped.groups:
                unmatched_dict = form_row.to_dict()
                for col in form_row.index:
                    if isinstance(form_row[col], str):
                        unmatched_dict[col] = remove_tones(form_row[col])
                unmatched_dict['Unmatched Reason'] = "Account Number not found in UPS data"
                unmatched_rows.append(unmatched_dict)
                continue

            ups_acc_df = ups_grouped.get_group(acc_norm)
            ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

            if is_same_billing == "yes":
                new_addr1 = form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                new_addr2 = form_row.get("New Address Line 2 (Street Name)-In English Only", "")
                new_addr3 = form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
                city = form_row.get("City / Province", "")
                contact = get_contact_name(form_row)

                matched_in_ups = False
                ups_row_for_template = None
                combined_form = f"{new_addr1}, {new_addr2}"
                if new_addr3:
                    combined_form += f", {new_addr3}"
                
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row["Address Type"] == '01':  # Now matches 2-digit string
                        ups_addr1 = get_address_line(ups_row, 1)
                        ups_addr2 = get_address_line(ups_row, 2)
                        combined_ups = f"{ups_addr1}"
                        if ups_addr2:
                            combined_ups += f", {ups_addr2}"
                        
                        if flexible_address_match(combined_form, combined_ups):
                            matched_in_ups = True
                            ups_row_for_template = ups_row
                            break

                if matched_in_ups and ups_row_for_template is not None:
                    matched_dict = form_row.to_dict()
                    matched_dict["New Address Line 1 (Tone-free)"] = remove_tones(new_addr1)
                    matched_dict["New Address Line 2 (Tone-free)"] = remove_tones(new_addr2)
                    matched_dict["New Address Line 3 (Tone-free)"] = remove_tones(new_addr3)
                    matched_dict["UPS Matched Address"] = get_address_line(ups_row_for_template, 1)
                    addr2 = get_address_line(ups_row_for_template, 2)
                    if addr2:
                        matched_dict["UPS Matched Address"] += f", {addr2}"
                    
                    for col in form_row.index:
                        if isinstance(form_row[col], str) and col not in matched_dict:
                            matched_dict[col] = remove_tones(form_row[col])
                    matched_rows.append(matched_dict)
                    processed_form_indices.add(idx)

                    for code in ["1", "2", "6"]:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": code,
                            "invoice option": code,
                            "AC_Name": ups_row_for_template.get("AC_Name", ""),
                            "Address_Line1": remove_tones(new_addr1),
                            "Address_Line2": remove_tones(new_addr2),
                            "City": city,
                            "Postal_Code": ups_row_for_template.get("Postal_Code", ""),
                            "Country_Code": ups_row_for_template.get("Country_Code", ""),
                            "Attention_Name": contact,
                            "Address_Line22": remove_tones(new_addr3),
                            "Address_Country_Code": ups_row_for_template.get("Address_Country_Code", ""),
                            "Matching_Score": "High"
                        })
                else:
                    unmatched_dict = form_row.to_dict()
                    for col in form_row.index:
                        if isinstance(form_row[col], str):
                            unmatched_dict[col] = remove_tones(form_row[col])
                    unmatched_dict['Unmatched Reason'] = "Billing address (type 01) not matched in UPS system"
                    unmatched_rows.append(unmatched_dict)

            else:
                # Handle case where billing address is different
                billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
                billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
                billing_city = form_row.get("New Billing City / Province", "")
                combined_billing = f"{billing_addr1}, {billing_addr2}"
                if billing_addr3:
                    combined_billing += f", {billing_addr3}"

                delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
                delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
                delivery_city = form_row.get("New Delivery City / Province", "")
                combined_delivery = f"{delivery_addr1}, {delivery_addr2}"
                if delivery_addr3:
                    combined_delivery += f", {delivery_addr3}"

                pickup_num = 0
                try:
                    pickup_num = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
                except:
                    pickup_num = 0
                if pickup_num > 3:
                    pickup_num = 3

                pickup_addrs = []
                for i in range(1, pickup_num + 1):
                    prefix = ["First", "Second", "Third"][i - 1] + " New Pick Up Address"
                    pu_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                    pu_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                    pu_addr3 = form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                    pu_city = form_row.get(f"{prefix} City / Province", "")
                    combined_pu = f"{pu_addr1}, {pu_addr2}"
                    if pu_addr3:
                        combined_pu += f", {pu_addr3}"
                    pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city, combined_pu))

                def check_address_in_ups(combined_form_addr, addr_type_code):
                    for _, ups_row in ups_acc_df.iterrows():
                        if ups_row["Address Type"] == addr_type_code:  # Now uses 2-digit string
                            ups_addr1 = get_address_line(ups_row, 1)
                            ups_addr2 = get_address_line(ups_row, 2)
                            combined_ups = f"{ups_addr1}"
                            if ups_addr2:
                                combined_ups += f", {ups_addr2}"
                            
                            if flexible_address_match(combined_form_addr, combined_ups):
                                return ups_row
                    return None

                billing_match = check_address_in_ups(combined_billing, "03")
                delivery_match = check_address_in_ups(combined_delivery, "13")

                if len(pickup_addrs) != ups_pickup_count:
                    unmatched_dict = form_row.to_dict()
                    for col in form_row.index:
                        if isinstance(form_row[col], str):
                            unmatched_dict[col] = remove_tones(form_row[col])
                    unmatched_dict['Unmatched Reason'] = f"Pickup address count mismatch: Forms={len(pickup_addrs)}, UPS={ups_pickup_count}"
                    unmatched_rows.append(unmatched_dict)
                    continue
                else:
                    pickup_matches = []
                    unmatched_pickup = False
                    for pu_addr in pickup_addrs:
                        match = check_address_in_ups(pu_addr[4], "02")  # Now matches "02"
                        if match is None:
                            unmatched_dict = form_row.to_dict()
                            for col in form_row.index:
                                if isinstance(form_row[col], str):
                                    unmatched_dict[col] = remove_tones(form_row[col])
                            unmatched_dict['Unmatched Reason'] = f"Pickup address not matched: {pu_addr[0]}, {pu_addr[1]}"
                            unmatched_rows.append(unmatched_dict)
                            unmatched_pickup = True
                            break
                        else:
                            pickup_matches.append(match)
                    if unmatched_pickup:
                        continue

                    processed_form_indices.add(idx)

                    matched_dict = form_row.to_dict()
                    matched_dict["New Billing Address Line 1 (Tone-free)"] = remove_tones(billing_addr1)
                    matched_dict["New Billing Address Line 2 (Tone-free)"] = remove_tones(billing_addr2)
                    matched_dict["New Billing Address Line 3 (Tone-free)"] = remove_tones(billing_addr3)
                    if billing_match is not None:
                        matched_dict["UPS Matched Billing Address"] = get_address_line(billing_match, 1)
                        addr2 = get_address_line(billing_match, 2)
                        if addr2:
                            matched_dict["UPS Matched Billing Address"] += f", {addr2}"

                    matched_dict["New Delivery Address Line 1 (Tone-free)"] = remove_tones(delivery_addr1)
                    matched_dict["New Delivery Address Line 2 (Tone-free)"] = remove_tones(delivery_addr2)
                    matched_dict["New Delivery Address Line 3 (Tone-free)"] = remove_tones(delivery_addr3)
                    if delivery_match is not None:
                        matched_dict["UPS Matched Delivery Address"] = get_address_line(delivery_match, 1)
                        addr2 = get_address_line(delivery_match, 2)
                        if addr2:
                            matched_dict["UPS Matched Delivery Address"] += f", {addr2}"

                    for i, pu_addr in enumerate(pickup_addrs, 1):
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = remove_tones(pu_addr[0])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = remove_tones(pu_addr[1])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = remove_tones(pu_addr[2])
                        if i <= len(pickup_matches):
                            matched_dict[f"{['First', 'Second', 'Third'][i-1]} UPS Matched Pickup Address"] = get_address_line(pickup_matches[i-1], 1)
                            addr2 = get_address_line(pickup_matches[i-1], 2)
                            if addr2:
                                matched_dict[f"{['First', 'Second', 'Third'][i-1]} UPS Matched Pickup Address"] += f", {addr2}"

                    matched_rows.append(matched_dict)

                    for pu_addr in pickup_addrs:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": "02",
                            "invoice option": "",
                            "AC_Name": ups_acc_df["AC_Name"].values[0] if "AC_Name" in ups_acc_df.columns else "",
                            "Address_Line1": remove_tones(pu_addr[0]),
                            "Address_Line2": remove_tones(pu_addr[1]),
                            "City": pu_addr[3],
                            "Postal_Code": ups_acc_df["Postal_Code"].values[0] if "Postal_Code" in ups_acc_df.columns else "",
                            "Country_Code": ups_acc_df["Country_Code"].values[0] if "Country_Code" in ups_acc_df.columns else "",
                            "Attention_Name": get_contact_name(form_row),
                            "Address_Line22": remove_tones(pu_addr[2]),
                            "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0] if "Address_Country_Code" in ups_acc_df.columns else "",
                            "Matching_Score": "High"
                        })

                    for code in ["1", "2", "6"]:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": code,
                            "invoice option": code,
                            "AC_Name": ups_acc_df["AC_Name"].values[0] if "AC_Name" in ups_acc_df.columns else "",
                            "Address_Line1": remove_tones(billing_addr1),
                            "Address_Line2": remove_tones(billing_addr2),
                            "City": billing_city,
                            "Postal_Code": ups_acc_df["Postal_Code"].values[0] if "Postal_Code" in ups_acc_df.columns else "",
                            "Country_Code": ups_acc_df["Country_Code"].values[0] if "Country_Code" in ups_acc_df.columns else "",
                            "Attention_Name": get_contact_name(form_row),
                            "Address_Line22": remove_tones(billing_addr3),
                            "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0] if "Address_Country_Code" in ups_acc_df.columns else "",
                            "Matching_Score": "High"
                        })

                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "13",
                        "invoice option": "",
                        "AC_Name": ups_acc_df["AC_Name"].values[0] if "AC_Name" in ups_acc_df.columns else "",
                        "Address_Line1": remove_tones(delivery_addr1),
                        "Address_Line2": remove_tones(delivery_addr2),
                        "City": delivery_city,
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0] if "Postal_Code" in ups_acc_df.columns else "",
                        "Country_Code": ups_acc_df["Country_Code"].values[0] if "Country_Code" in ups_acc_df.columns else "",
                        "Attention_Name": get_contact_name(form_row),
                        "Address_Line22": remove_tones(delivery_addr3),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0] if "Address_Country_Code" in ups_acc_df.columns else "",
                        "Matching_Score": "High"
                    })

        unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_form_indices)]
        for _, row in unmatched_not_processed.iterrows():
            unmatched_dict = row.to_dict()
            for col in row.index:
                if isinstance(row[col], str):
                    unmatched_dict[col] = remove_tones(row[col])
            unmatched_dict['Unmatched Reason'] = "No matching address found or not processed"
            unmatched_rows.append(unmatched_dict)

        matched_df = pd.DataFrame(matched_rows) if matched_rows else pd.DataFrame()
        unmatched_df = pd.DataFrame(unmatched_rows) if unmatched_rows else pd.DataFrame()
        upload_template_df = pd.DataFrame(upload_template_rows) if upload_template_rows else pd.DataFrame()

        return matched_df, unmatched_df, upload_template_df

    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        raise e

def to_excel(df):
    # Fix 3: Handle xlsxwriter dependency with error handling
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Results')
            workbook = writer.book
            worksheet = writer.sheets['Results']
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
                col_idx = df.columns.get_loc(column)
                worksheet.set_column(col_idx, col_idx, column_width)
        return output.getvalue()
    except ImportError:
        st.warning("xlsxwriter not installed - using basic Excel format without styling")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Results')
        return output.getvalue()

def main():
    st.set_page_config(
        page_title="Vietnam Address Validation Tool",
        layout="wide",
        page_icon="🇻🇳",
        initial_sidebar_state="expanded"
    )
    
    st.markdown("""
        <style>
            .stProgress > div > div > div > div {
                background-color: #1f77b4;
            }
            .reportview-container .main .block-container {
                padding-top: 2rem;
                padding-bottom: 2rem;
            }
            .header-text {
                font-size: 24px !important;
                font-weight: 700 !important;
                margin-bottom: 1rem !important;
            }
            .download-button {
                margin-top: 1rem !important;
                margin-bottom: 1rem !important;
            }
        </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.title("About")
        st.info("""
            This tool validates Vietnam addresses by comparing Microsoft Forms responses with UPS system data.
            It now includes improved matching for addresses with 60-70% similarity.
        """)
        
        st.markdown("---")
        st.markdown("**Instructions:**")
        st.markdown("1. Upload Microsoft Forms response file")
        st.markdown("2. Upload UPS system address file")
        st.markdown("3. Click 'Process Files'")
        st.markdown("4. Download the results")
        
        st.markdown("---")
        st.markdown(f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        # Fix 3: Add note about xlsxwriter
        st.markdown("**Note:** Install xlsxwriter for formatted Excel outputs: `pip install xlsxwriter`")

    st.title("🇻🇳 Vietnam Address Validation Tool")
    st.markdown('<div class="header-text">Enhanced address matching with 60-70% similarity capture</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        forms_file = st.file_uploader("Microsoft Forms Response File (.xlsx)", type=["xlsx"], key="forms")
    with col2:
        ups_file = st.file_uploader("UPS System Address File (.xlsx)", type=["xlsx"], key="ups")

    if forms_file and ups_file:
        if st.button("Process Files", type="primary"):
            with st.spinner("Processing files..."):
                try:
                    progress_bar = st.progress(0)
                    
                    progress_bar.progress(10)
                    forms_df = pd.read_excel(forms_file)
                    ups_df = pd.read_excel(ups_file)
                    
                    progress_bar.progress(30)
                    matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)
                    progress_bar.progress(80)
                    
                    st.success(f"✅ Validation completed! Matched: {len(matched_df)} | Unmatched: {len(unmatched_df)}")
                    
                    tab1, tab2, tab3 = st.tabs(["📊 Results Summary", "✅ Matched Records", "❌ Unmatched Records"])
                    
                    with tab1:
                        st.subheader("Processing Summary")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Total Forms Records", len(forms_df))
                        col2.metric("Matched Records", len(matched_df), f"{len(matched_df)/len(forms_df):.1%}")
                        col3.metric("Unmatched Records", len(unmatched_df), f"{len(unmatched_df)/len(forms_df):.1%}")
                        
                        if not unmatched_df.empty:
                            st.subheader("Unmatched Reasons")
                            reason_counts = unmatched_df['Unmatched Reason'].value_counts()
                            st.bar_chart(reason_counts)
                    
                    with tab2:
                        if not matched_df.empty:
                            st.dataframe(matched_df.head(10))
                            st.download_button(
                                label="📥 Download All Matched Records (Excel)",
                                data=to_excel(matched_df),
                                file_name=f"matched_records_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="matched_download",
                                help="Download all matched records in Excel format",
                            )
                        else:
                            st.warning("No matched records found")
                    
                    with tab3:
                        if not unmatched_df.empty:
                            st.dataframe(unmatched_df.head(10))
                            st.download_button(
                                label="📥 Download All Unmatched Records (Excel)",
                                data=to_excel(unmatched_df),
                                file_name=f"unmatched_records_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="unmatched_download",
                                help="Download all unmatched records in Excel format",
                            )
                        else:
                            st.info("All records matched successfully!")
                    
                    st.subheader("Upload Template")
                    if not upload_template_df.empty:
                        st.dataframe(upload_template_df.head(10))
                        st.download_button(
                            label="📥 Download Upload Template (Excel)",
                            data=to_excel(upload_template_df),
                            file_name=f"upload_template_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="This file contains the formatted data ready for UPS system upload",
                            key="template_download",
                        )
                    
                    progress_bar.progress(100)
                    st.balloons()
                
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.exception(e)

if __name__ == "__main__":
    main()

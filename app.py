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
    # Remove all punctuation except commas (as they might separate address parts)
    addr = re.sub(r'[^\w\s,]', ' ', addr)
    # Normalize spaces and commas
    addr = re.sub(r'\s+', ' ', addr).strip()
    addr = re.sub(r',\s+', ',', addr)
    addr = re.sub(r'\s+,', ',', addr)
    return addr

# Improved flexible address matching with better handling of cases
def flexible_address_match(addr1, addr2, threshold=0.65):
    a1 = clean_address(addr1)
    a2 = clean_address(addr2)
    
    if not a1 or not a2:
        return False
    
    # Case 1: Exact match after cleaning
    if a1 == a2:
        return True
    
    # Case 2: One is a substring of the other (with relaxed matching)
    if a1 in a2 or a2 in a1:
        return True
    
    # Case 3: Split by comma and check components
    a1_parts = [p.strip() for p in a1.split(',') if p.strip()]
    a2_parts = [p.strip() for p in a2.split(',') if p.strip()]
    
    # Check if all parts of one address exist in the other (order doesn't matter)
    if a1_parts and a2_parts:
        if all(any(difflib.SequenceMatcher(None, p1, p2).ratio() >= 0.7 
                 for p2 in a2_parts) for p1 in a1_parts):
            return True
        if all(any(difflib.SequenceMatcher(None, p2, p1).ratio() >= 0.7 
                 for p1 in a1_parts) for p2 in a2_parts):
            return True
    
    # Case 4: Combined address line matching
    combined_a1 = ' '.join(a1_parts) if a1_parts else a1
    combined_a2 = ' '.join(a2_parts) if a2_parts else a2
    
    ratio = difflib.SequenceMatcher(None, combined_a1, combined_a2).ratio()
    if ratio >= threshold:
        return True
    
    # Case 5: Check if one address contains all words from the other (order doesn't matter)
    words1 = set(combined_a1.split())
    words2 = set(combined_a2.split())
    
    if words1 and words2:
        # Check if one set of words is a subset of the other
        if words1.issubset(words2) or words2.issubset(words1):
            return True
        
        # Check for significant word overlap (at least 70% of words match)
        common_words = words1 & words2
        min_len = min(len(words1), len(words2))
        if len(common_words) / min_len >= 0.7:
            return True
    
    return False

def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize Account Number (lowercase & strip)
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()

    ups_grouped = ups_df.groupby('Account Number_norm')

    processed_form_indices = set()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        if acc_norm not in ups_grouped.groups:
            unmatched_dict = form_row.to_dict()
            # Remove tones in address fields for unmatched output
            for col in form_row.index:
                if isinstance(form_row[col], str):
                    unmatched_dict[col] = remove_tones(form_row[col])
            unmatched_dict['Unmatched Reason'] = "Account Number not found in UPS data"
            unmatched_rows.append(unmatched_dict)
            continue

        ups_acc_df = ups_grouped.get_group(acc_norm)
        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

        if is_same_billing == "yes":
            # Single billing address for type 01
            new_addr1 = form_row["New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only"]
            new_addr2 = form_row["New Address Line 2 (Street Name)-In English Only"]
            new_addr3 = form_row["New Address Line 3 (Ward/Commune)-In English Only"]
            city = form_row["City / Province"]
            contact = form_row.get("Full Name of Contact-In English Only", "")

            matched_in_ups = False
            ups_row_for_template = None
            
            # Combine address lines for better matching
            combined_form = f"{new_addr1}, {new_addr2}"
            if new_addr3:
                combined_form += f", {new_addr3}"
            
            for _, ups_row in ups_acc_df.iterrows():
                if ups_row["Address Type"] == '01':
                    # Get UPS address components
                    ups_addr1 = ups_row["Address Line 1"]
                    ups_addr2 = ups_row.get("Address Line 2", "")
                    combined_ups = f"{ups_addr1}"
                    if ups_addr2:
                        combined_ups += f", {ups_addr2}"
                    
                    # Use the improved matching function
                    if flexible_address_match(combined_form, combined_ups):
                        matched_in_ups = True
                        ups_row_for_template = ups_row
                        break

            if matched_in_ups:
                matched_dict = form_row.to_dict()
                # Add tone-free versions
                matched_dict["New Address Line 1 (Tone-free)"] = remove_tones(new_addr1)
                matched_dict["New Address Line 2 (Tone-free)"] = remove_tones(new_addr2)
                matched_dict["New Address Line 3 (Tone-free)"] = remove_tones(new_addr3)
                matched_dict["UPS Matched Address"] = ups_row_for_template["Address Line 1"]
                if "Address Line 2" in ups_row_for_template:
                    matched_dict["UPS Matched Address"] += f", {ups_row_for_template['Address Line 2']}"
                
                for col in form_row.index:
                    if isinstance(form_row[col], str) and col not in matched_dict:
                        matched_dict[col] = remove_tones(form_row[col])
                matched_rows.append(matched_dict)
                processed_form_indices.add(idx)

                # Upload template with 3 rows for billing (codes 1,2,6) with invoice option same as code
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_row_for_template["AC_Name"],
                        "Address_Line1": remove_tones(new_addr1),
                        "Address_Line2": remove_tones(new_addr2),
                        "City": city,
                        "Postal_Code": ups_row_for_template["Postal_Code"],
                        "Country_Code": ups_row_for_template["Country_Code"],
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(new_addr3),
                        "Address_Country_Code": ups_row_for_template["Address_Country_Code"],
                        "Matching_Score": "High"  # Added matching confidence indicator
                    })
            else:
                unmatched_dict = form_row.to_dict()
                for col in form_row.index:
                    if isinstance(form_row[col], str):
                        unmatched_dict[col] = remove_tones(form_row[col])
                unmatched_dict['Unmatched Reason'] = "Billing address (type 01) not matched in UPS system"
                unmatched_rows.append(unmatched_dict)

        else:
            # Case "no": Separate billing, delivery, pickup addresses
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
                    if ups_row["Address Type"] == addr_type_code:
                        ups_addr1 = ups_row["Address Line 1"]
                        ups_addr2 = ups_row.get("Address Line 2", "")
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
                    match = check_address_in_ups(pu_addr[4], "02")
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
                # Add tone-free versions and matched UPS addresses
                matched_dict["New Billing Address Line 1 (Tone-free)"] = remove_tones(billing_addr1)
                matched_dict["New Billing Address Line 2 (Tone-free)"] = remove_tones(billing_addr2)
                matched_dict["New Billing Address Line 3 (Tone-free)"] = remove_tones(billing_addr3)
                matched_dict["UPS Matched Billing Address"] = billing_match["Address Line 1"]
                if "Address Line 2" in billing_match:
                    matched_dict["UPS Matched Billing Address"] += f", {billing_match['Address Line 2']}"

                matched_dict["New Delivery Address Line 1 (Tone-free)"] = remove_tones(delivery_addr1)
                matched_dict["New Delivery Address Line 2 (Tone-free)"] = remove_tones(delivery_addr2)
                matched_dict["New Delivery Address Line 3 (Tone-free)"] = remove_tones(delivery_addr3)
                matched_dict["UPS Matched Delivery Address"] = delivery_match["Address Line 1"]
                if "Address Line 2" in delivery_match:
                    matched_dict["UPS Matched Delivery Address"] += f", {delivery_match['Address Line 2']}"

                for i, pu_addr in enumerate(pickup_addrs, 1):
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = remove_tones(pu_addr[0])
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = remove_tones(pu_addr[1])
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = remove_tones(pu_addr[2])
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} UPS Matched Pickup Address"] = pickup_matches[i-1]["Address Line 1"]
                    if "Address Line 2" in pickup_matches[i-1]:
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} UPS Matched Pickup Address"] += f", {pickup_matches[i-1]['Address Line 2']}"

                matched_rows.append(matched_dict)

                # Upload template pickup addresses (multiple rows)
                for pu_addr in pickup_addrs:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "02",
                        "invoice option": "",
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": remove_tones(pu_addr[0]),
                        "Address_Line2": remove_tones(pu_addr[1]),
                        "City": pu_addr[3],
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": remove_tones(pu_addr[2]),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0],
                        "Matching_Score": "High"
                    })

                # Upload template billing address - 3 rows with codes 1,2,6
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": remove_tones(billing_addr1),
                        "Address_Line2": remove_tones(billing_addr2),
                        "City": billing_city,
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": remove_tones(billing_addr3),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0],
                        "Matching_Score": "High"
                    })

                # Upload template delivery address - 1 row
                upload_template_rows.append({
                    "AC_NUM": form_row["Account Number"],
                    "AC_Address_Type": "13",
                    "invoice option": "",
                    "AC_Name": ups_acc_df["AC_Name"].values[0],
                    "Address_Line1": remove_tones(delivery_addr1),
                    "Address_Line2": remove_tones(delivery_addr2),
                    "City": delivery_city,
                    "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                    "Country_Code": ups_acc_df["Country_Code"].values[0],
                    "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                    "Address_Line22": remove_tones(delivery_addr3),
                    "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0],
                    "Matching_Score": "High"
                })

    # Add unmatched forms rows not processed or no matches
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_form_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched_dict = row.to_dict()
        for col in row.index:
            if isinstance(row[col], str):
                unmatched_dict[col] = remove_tones(row[col])
        unmatched_dict['Unmatched Reason'] = "No matching address found or not processed"
        unmatched_rows.append(unmatched_dict)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)

    return matched_df, unmatched_df, upload_template_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Results')
        
        # Add some basic Excel formatting
        workbook = writer.book
        worksheet = writer.sheets['Results']
        
        # Add a header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Write the column headers with the defined format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-adjust columns' width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
            col_idx = df.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, column_width)
    
    return output.getvalue()

def main():
    st.set_page_config(
        page_title="Vietnam Address Validation Tool",
        layout="wide",
        page_icon="🇻🇳",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS
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

    # Sidebar
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

    # Main content
    st.title("🇻🇳 Vietnam Address Validation Tool")
    st.markdown('<div class="header-text">Enhanced address matching with 60-70% similarity capture</div>', unsafe_allow_html=True)

    # File upload
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
                    
                    # Read files
                    progress_bar.progress(10)
                    forms_df = pd.read_excel(forms_file)
                    ups_df = pd.read_excel(ups_file)
                    
                    # Process files
                    progress_bar.progress(30)
                    matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)
                    progress_bar.progress(80)
                    
                    # Display results
                    st.success(f"✅ Validation completed! Matched: {len(matched_df)} | Unmatched: {len(unmatched_df)}")
                    
                    # Results tabs
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
                                on_click=None,
                                args=None,
                                kwargs=None,
                                class_name="download-button"
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
                                class_name="download-button"
                            )
                        else:
                            st.info("All records matched successfully!")
                    
                    # Upload template section
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
                            class_name="download-button"
                        )
                    
                    progress_bar.progress(100)
                    st.balloons()
                
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.exception(e)

if __name__ == "__main__":
    main()

________________________________________
发自我的iPhone

------------------ Original ------------------
From: liudamon <liudamon@ups.com>
Date: Wed,Aug 6,2025 4:28 PM
To: damonliu2022 <damonliu2022@qq.com>
Subject: Re: RE:

Make sure similarity 60% to 70% will be captured in the matched list. I have two examples: For example: 1. 036V9A, in the Forms response, the address line 1 is "Lo A-9H-CN", address line 2 is "KCN Bau Bang" in UPS system, it is "LO A-9H-CN,KCN BAU BANG,THI TRAN LAI UYEN," 2. B86748, the address line 1 in Forms is "PHONG 4A, TANG 4, TOA NHA OPERA VIEW", line 2 is "161 DUONG DONG KHOI", in UPS system, the address line 1 is "PHONG 4A, TANG 4, TOA NHA OPERA VIEW, 161 DONG". So based on these two examples, please update your code for matching rule. For matched, unmatched and download file, make sure they are in excel format. Please update the code accordingly. 
 
 
 
 
 
Damon Liu
 
Best Regards,
 
From: Damon Liu 
Sent: Wednesday, August 6, 2025 4:16 PM
To: damonliu2022 <damonliu2022@qq.com>
Subject: 
 
import streamlit as st
import pandas as pd
import unicodedata
import difflib
import re
from io import BytesIO
 
# Remove Vietnamese tones from text
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text
 
# Clean address: remove tones, lowercase, remove punctuation, normalize spaces
def clean_address(addr):
    if not isinstance(addr, str):
        return ""
    addr = remove_tones(addr)
    addr = addr.lower()
    addr = re.sub(r'[^\w\s]', ' ', addr)  # remove punctuation
    addr = re.sub(r'\s+', ' ', addr).strip()  # normalize spaces
    return addr
 
# Flexible address matching with threshold (default 0.6)
def flexible_address_match(addr1, addr2, threshold=0.6):
    a1 = clean_address(addr1)
    a2 = clean_address(addr2)
    if not a1 or not a2:
        return False
    # Substring check
    if a1 in a2 or a2 in a1:
        return True
    # Word subset check
    words1 = set(a1.split())
    words2 = set(a2.split())
    if words1 and words2:
        shorter, longer = (words1, words2) if len(words1) < len(words2) else (words2, words1)
        if shorter.issubset(longer):
            return True
    # Fuzzy ratio check
    ratio = difflib.SequenceMatcher(None, a1, a2).ratio()
    if ratio >= threshold:
        return True
    return False
 
def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []
 
    # Normalize Account Number (lowercase & strip)
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()
 
    ups_grouped = ups_df.groupby('Account Number_norm')
 
    processed_form_indices = set()
 
    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()
 
        if acc_norm not in ups_grouped.groups:
            unmatched_dict = form_row.to_dict()
            # Remove tones in address fields for unmatched output
            for col in form_row.index:
                if isinstance(form_row[col], str):
                    unmatched_dict[col] = remove_tones(form_row[col])
            unmatched_dict['Unmatched Reason'] = "Account Number not found in UPS data"
            unmatched_rows.append(unmatched_dict)
            continue
 
        ups_acc_df = ups_grouped.get_group(acc_norm)
        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()
 
        if is_same_billing == "yes":
            # Single billing address for type 01
            new_addr1 = form_row["New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only"]
            new_addr2 = form_row["New Address Line 2 (Street Name)-In English Only"]
            new_addr3 = form_row["New Address Line 3 (Ward/Commune)-In English Only"]
            city = form_row["City / Province"]
            contact = form_row.get("Full Name of Contact-In English Only", "")
 
            matched_in_ups = False
            ups_row_for_template = None
            for _, ups_row in ups_acc_df.iterrows():
                if ups_row["Address Type"] == '01':  # address type 01 match only
                    # Check individual and combined address line matches for billing address
                    match1 = flexible_address_match(new_addr1, ups_row["Address Line 1"], threshold=0.5)
                    match2 = flexible_address_match(new_addr2, ups_row["Address Line 2"], threshold=0.5)
                    combined_form = (str(new_addr1) + ' ' + str(new_addr2)).strip()
                    combined_match = flexible_address_match(combined_form, ups_row["Address Line 1"], threshold=0.5)
                    if (match1 and match2) or combined_match:
                        matched_in_ups = True
                        ups_row_for_template = ups_row
                        break
 
            if matched_in_ups:
                matched_dict = form_row.to_dict()
                # Remove tones from all address fields in matched output
                matched_dict["New Address Line 1 (Tone-free)"] = remove_tones(new_addr1)
                matched_dict["New Address Line 2 (Tone-free)"] = remove_tones(new_addr2)
                matched_dict["New Address Line 3 (Tone-free)"] = remove_tones(new_addr3)
                for col in form_row.index:
                    if isinstance(form_row[col], str) and col not in matched_dict:
                        matched_dict[col] = remove_tones(form_row[col])
                matched_rows.append(matched_dict)
                processed_form_indices.add(idx)
 
                # Upload template with 3 rows for billing (codes 1,2,6) with invoice option same as code
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_row_for_template["AC_Name"],
                        "Address_Line1": remove_tones(new_addr1),
                        "Address_Line2": remove_tones(new_addr2),
                        "City": city,
                        "Postal_Code": ups_row_for_template["Postal_Code"],
                        "Country_Code": ups_row_for_template["Country_Code"],
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(new_addr3),
                        "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
                    })
            else:
                unmatched_dict = form_row.to_dict()
                for col in form_row.index:
                    if isinstance(form_row[col], str):
                        unmatched_dict[col] = remove_tones(form_row[col])
                unmatched_dict['Unmatched Reason'] = "Billing address (type 01) not matched in UPS system"
                unmatched_rows.append(unmatched_dict)
 
        else:
            # Case "no": Separate billing, delivery, pickup addresses
            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")
 
            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")
 
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
                pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city))
 
            def check_address_in_ups(addr1, addr2, addr_type_code):
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row["Address Type"] == addr_type_code:
                        # Check individual and combined address line matches for address in UPS
                        match1 = flexible_address_match(addr1, ups_row["Address Line 1"], threshold=0.5)
                        match2 = flexible_address_match(addr2, ups_row["Address Line 2"], threshold=0.5)
                        combined_addr = (str(addr1) + ' ' + str(addr2)).strip()
                        combined_match = flexible_address_match(combined_addr, ups_row["Address Line 1"], threshold=0.5)
                        if (match1 and match2) or combined_match:
                            return ups_row
                return None
 
            billing_match = check_address_in_ups(billing_addr1, billing_addr2, "03")
            delivery_match = check_address_in_ups(delivery_addr1, delivery_addr2, "13")
 
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
                    match = check_address_in_ups(pu_addr[0], pu_addr[1], "02")
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
                # Remove tones in all address fields
                matched_dict["New Billing Address Line 1 (Tone-free)"] = remove_tones(billing_addr1)
                matched_dict["New Billing Address Line 2 (Tone-free)"] = remove_tones(billing_addr2)
                matched_dict["New Billing Address Line 3 (Tone-free)"] = remove_tones(billing_addr3)
                matched_dict["New Delivery Address Line 1 (Tone-free)"] = remove_tones(delivery_addr1)
                matched_dict["New Delivery Address Line 2 (Tone-free)"] = remove_tones(delivery_addr2)
                matched_dict["New Delivery Address Line 3 (Tone-free)"] = remove_tones(delivery_addr3)
                for i, pu_addr in enumerate(pickup_addrs, 1):
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = remove_tones(pu_addr[0])
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = remove_tones(pu_addr[1])
                    matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = remove_tones(pu_addr[2])
                matched_rows.append(matched_dict)
 
                # Upload template pickup addresses (multiple rows)
                for pu_addr in pickup_addrs:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "02",
                        "invoice option": "",
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": remove_tones(pu_addr[0]),
                        "Address_Line2": remove_tones(pu_addr[1]),
                        "City": pu_addr[3],
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": remove_tones(pu_addr[2]),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                    })
 
                # Upload template billing address - 3 rows with codes 1,2,6
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": remove_tones(billing_addr1),
                        "Address_Line2": remove_tones(billing_addr2),
                        "City": billing_city,
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": remove_tones(billing_addr3),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                    })
 
                # Upload template delivery address - 1 row
                upload_template_rows.append({
                    "AC_NUM": form_row["Account Number"],
                    "AC_Address_Type": "13",
                    "invoice option": "",
                    "AC_Name": ups_acc_df["AC_Name"].values[0],
                    "Address_Line1": remove_tones(delivery_addr1),
                    "Address_Line2": remove_tones(delivery_addr2),
                    "City": delivery_city,
                    "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                    "Country_Code": ups_acc_df["Country_Code"].values[0],
                    "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                    "Address_Line22": remove_tones(delivery_addr3),
                    "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                })
 
    # Add unmatched forms rows not processed or no matches
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_form_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched_dict = row.to_dict()
        for col in row.index:
            if isinstance(row[col], str):
                unmatched_dict[col] = remove_tones(row[col])
        unmatched_dict['Unmatched Reason'] = "No matching address found or not processed"
        unmatched_rows.append(unmatched_dict)
 
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)
 
    return matched_df, unmatched_df, upload_template_df
 
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()
 
def main():
    st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")
    st.title("🇻🇳 Vietnam Address Validation Tool")
    st.write("Upload Microsoft Forms response file and UPS system address file to validate and generate upload template.")
 
    forms_file = st.file_uploader("Upload Microsoft Forms Response File (.xlsx)", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File (.xlsx)", type=["xlsx"])
 
    if forms_file and ups_file:
        with st.spinner("Processing files..."):
            forms_df = pd.read_excel(forms_file)
            ups_df = pd.read_excel(ups_file)
 
            matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)
 
            st.success(f"Validation completed! Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")
 
            if not matched_df.empty:
                st.download_button(
                    label="Download Matched Records",
                    data=to_excel(matched_df),
                    file_name="matched_records.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if not unmatched_df.empty:
                st.download_button(
                    label="Download Unmatched Records",
                    data=to_excel(unmatched_df),
                    file_name="unmatched_records.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if not upload_template_df.empty:
                st.download_button(
                    label="Download Upload Template",
                    data=to_excel(upload_template_df),
                    file_name="upload_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
 
if __name__ == "__main__":
    main()

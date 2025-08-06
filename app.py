import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO
from rapidfuzz import fuzz

# Remove Vietnamese tones
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text

# Normalize string columns (lowercase, strip spaces, remove tones)
def normalize_col(col):
    return col.astype(str).str.lower().str.strip().apply(remove_tones)

# Improved matching function – slightly loosened
def is_address_match(form_line1, form_line2, sys_line1, sys_line2):
    if not isinstance(form_line1, str) or not isinstance(sys_line1, str):
        return False
    if not isinstance(form_line2, str):
        form_line2 = ''
    if not isinstance(sys_line2, str):
        sys_line2 = ''

    # Combine address lines for better context
    form_full = f"{form_line1} {form_line2}".lower().strip()
    sys_full = f"{sys_line1} {sys_line2}".lower().strip()

    form_full = remove_tones(form_full)
    sys_full = remove_tones(sys_full)

    # Loose partial match threshold
    score = fuzz.partial_ratio(form_full, sys_full)

    return score >= 75  # lowered from 85–90 to 75 for looser match

# File upload
st.title("Vietnam Customer Address Validation Tool")
forms_file = st.file_uploader("Upload Microsoft Forms Response Excel", type=["xlsx"])
system_file = st.file_uploader("Upload UPS System Data Excel", type=["xlsx"])

if forms_file and system_file:
    forms_df = pd.read_excel(forms_file)
    system_df = pd.read_excel(system_file)

    # Normalize system data
    system_df['AC_NUM'] = normalize_col(system_df['AC_NUM'])
    system_df['Address_Line1'] = normalize_col(system_df['Address_Line1'])
    system_df['Address_Line2'] = normalize_col(system_df['Address_Line2'])

    matched_rows = []
    unmatched_rows = []

    for idx, row in forms_df.iterrows():
        acct = str(row['Account Number']).strip().lower()
        same_billing = str(row['Is Your New Billing Address the Same as Your Pickup and Delivery Address?']).strip().lower()

        sys_acct_df = system_df[system_df['AC_NUM'] == acct]

        def match_and_store(address_type_code, line1_col, line2_col):
            new_line1 = str(row.get(line1_col, '')).strip()
            new_line2 = str(row.get(line2_col, '')).strip()
            found = False
            for _, sys_row in sys_acct_df.iterrows():
                if is_address_match(new_line1, new_line2, sys_row['Address_Line1'], sys_row['Address_Line2']):
                    matched_rows.append({
                        "AC_NUM": acct.upper(),
                        "AC_Address_Type": address_type_code,
                        "AC_Name": sys_row.get("AC_Name", ""),
                        "Address_Line1": new_line1,
                        "Address_Line2": new_line2,
                        "City": row.get("City / Province", ""),
                        "Postal_Code": "",
                        "Country_Code": "VN",
                        "Attention_Name": row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": "",
                        "Address_Country_Code": "VN"
                    })
                    found = True
                    break
            return found

        matched = False
        if same_billing.lower() == 'yes':
            matched = match_and_store("01", "New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "New Address Line 2 (Street Name)-In English Only")
        else:
            pickup_matched = match_and_store("02", "New Pick Up Address Line 1 (Address No., etc)-In English Only", "New Pick Up Address Line 2 (Street Name)-In English Only")
            billing_matched = match_and_store("03", "New Billing Address Line 1 (Address No., etc)-In English Only", "New Billing Address Line 2 (Street Name)-In English Only")
            delivery_matched = match_and_store("13", "New Delivery Address Line 1 (Address No., etc)-In English Only", "New Delivery Address Line 2 (Street Name)-In English Only")
            matched = pickup_matched or billing_matched or delivery_matched

        if not matched:
            row['Unmatched Reason'] = "No close match in UPS system"
            unmatched_rows.append(row)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)

    def to_excel_download(df, filename):
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)
        st.download_button(label=f"Download {filename}", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.success(f"✅ Process complete. Matched: {len(matched_df)} | Unmatched: {len(unmatched_df)}")
    to_excel_download(matched_df, "matched_addresses.xlsx")
    to_excel_download(unmatched_df, "unmatched_forms_responses.xlsx")

import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO
from rapidfuzz import fuzz
import re

st.set_page_config(layout="wide")

# ------------------------ Utility Functions ------------------------

# Remove Vietnamese tones
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

# Normalize string for comparison
def normalize_text(text):
    if not isinstance(text, str):
        return ''
    text = remove_tones(text).lower().strip()
    return re.sub(r'\s+', ' ', text)  # remove extra spaces

# Token similarity: percentage of tokens from a in b
def token_similarity(a, b):
    if not a or not b:
        return 0
    a_tokens = set(a.split())
    b_tokens = set(b.split())
    match_count = len(a_tokens & b_tokens)
    return match_count / max(len(a_tokens), 1)

# Compare address lines
def is_address_match(new1, new2, old1, old2):
    n1, n2 = normalize_text(new1), normalize_text(new2)
    o1, o2 = normalize_text(old1), normalize_text(old2)

    # Combine lines for fuzzy match
    full_new = f"{n1} {n2}".strip()
    full_old = f"{o1} {o2}".strip()

    full_ratio = fuzz.token_set_ratio(full_new, full_old)
    token_match_1 = token_similarity(n1, o1)
    token_match_2 = token_similarity(n2, o2)

    return full_ratio >= 70 or token_match_1 >= 0.6 or token_match_2 >= 0.6

# ------------------------ Streamlit UI ------------------------

st.title("Vietnam Address Validation Tool")

uploaded_forms = st.file_uploader("Upload Microsoft Forms Excel File", type=["xlsx"], key="forms")
uploaded_system = st.file_uploader("Upload UPS System Address Excel File", type=["xlsx"], key="system")

if uploaded_forms and uploaded_system:
    forms_df = pd.read_excel(uploaded_forms)
    system_df = pd.read_excel(uploaded_system)

    # Validate required columns
    required_columns = ["Account Number", "Is Your New Billing Address the Same as Your Pickup and Delivery Address?",
                        "New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                        "New Address Line 2 (Street Name)-In English Only",
                        "City / Province"]
    
    if not all(col in forms_df.columns for col in required_columns):
        st.error("❌ One or more required columns are missing in the Forms file.")
    elif 'AC_NUM' not in system_df.columns:
        st.error("❌ The column 'AC_NUM' is missing in the UPS System file.")
    else:
        # Initialize outputs
        matched_rows = []
        unmatched_rows = []

        for _, row in forms_df.iterrows():
            acct = str(row["Account Number"]).strip()
            same_all = str(row["Is Your New Billing Address the Same as Your Pickup and Delivery Address?"]).strip().lower()
            new1 = row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            new2 = row.get("New Address Line 2 (Street Name)-In English Only", "")
            city = row.get("City / Province", "")
            
            # Normalize once
            new1_nf = normalize_text(new1)
            new2_nf = normalize_text(new2)

            system_acct_rows = system_df[system_df['AC_NUM'].astype(str).str.strip() == acct]

            if system_acct_rows.empty:
                unmatched_rows.append({**row, "Unmatched Reason": "Account not found in system"})
                continue

            matched = False
            for _, sys_row in system_acct_rows.iterrows():
                old1 = sys_row.get("Address_Line1", "")
                old2 = sys_row.get("Address_Line2", "")
                if is_address_match(new1, new2, old1, old2):
                    match_entry = {
                        "AC_NUM": acct,
                        "AC_Address_Type": "01" if same_all == "yes" else "02",
                        "AC_Name": sys_row.get("AC_Name", ""),
                        "Address_Line1": new1,
                        "Address_Line2": new2,
                        "City": city,
                        "Postal_Code": sys_row.get("Postal_Code", ""),
                        "Country_Code": sys_row.get("Country_Code", "VN"),
                        "Attention_Name": sys_row.get("Attention_Name", ""),
                        "Address_Line22": "",
                        "Address_Country_Code": sys_row.get("Country_Code", "VN"),
                    }
                    matched_rows.append(match_entry)
                    matched = True
                    break

            if not matched:
                unmatched_rows.append({**row, "Unmatched Reason": "Address did not match any record"})

        # Export buttons
        matched_df = pd.DataFrame(matched_rows)
        unmatched_df = pd.DataFrame(unmatched_rows)

        def convert_df(df):
            output = BytesIO()
            df.to_excel(output, index=False)
            return output.getvalue()

        st.success(f"✅ Matching complete. {len(matched_df)} matched, {len(unmatched_df)} unmatched.")

        st.download_button("Download Matched File", convert_df(matched_df), "matched.xlsx")
        st.download_button("Download Unmatched File", convert_df(unmatched_df), "unmatched.xlsx")

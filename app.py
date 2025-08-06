import streamlit as st
import pandas as pd
import unicodedata
from rapidfuzz import fuzz
from io import BytesIO

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")

# Function to remove Vietnamese tones
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize("NFD", text)
    text = "".join([c for c in text if unicodedata.category(c) != "Mn"])
    return text

# Normalize and remove tones
def normalize(text):
    if pd.isna(text):
        return ""
    return remove_tones(str(text)).strip().lower()

# Combine address line 1 and line 2 for matching
def combine_address_lines(row, line1_col, line2_col):
    line1 = str(row.get(line1_col, "")).strip()
    line2 = str(row.get(line2_col, "")).strip()
    return f"{line1}, {line2}" if line2 else line1

# Upload files
st.title("Vietnam Customer Address Validation Tool")
forms_file = st.file_uploader("Upload Microsoft Forms Response", type=["xlsx"])
system_file = st.file_uploader("Upload UPS System Data", type=["xlsx"])

if forms_file and system_file:
    # Read files
    forms_df = pd.read_excel(forms_file)
    system_df = pd.read_excel(system_file)

    # Drop empty rows
    forms_df.dropna(subset=["Account Number"], inplace=True)
    system_df.dropna(subset=["AC_NUM", "Address_Line1"], inplace=True)

    # Normalize system address for matching
    system_df["Norm_AC_NUM"] = system_df["AC_NUM"].apply(str).str.strip().str.upper()
    system_df["Normalized_Line1"] = system_df["Address_Line1"].apply(normalize)

    # Output DataFrames
    matched_rows = []
    unmatched_rows = []

    for _, form_row in forms_df.iterrows():
        acct = str(form_row["Account Number"]).strip().upper()

        form_line1 = normalize(form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""))
        form_line2 = normalize(form_row.get("New Address Line 2 (Street Name)-In English Only", ""))

        full_form_address = f"{form_line1}, {form_line2}".strip(", ")

        # Filter system addresses by account number
        system_candidates = system_df[system_df["Norm_AC_NUM"] == acct]

        best_score = 0
        best_match = None

        for _, sys_row in system_candidates.iterrows():
            sys_address = sys_row["Normalized_Line1"]
            score = fuzz.partial_ratio(full_form_address, sys_address)
            if score > best_score:
                best_score = score
                best_match = sys_row

        if best_score >= 85:  # Loosened threshold from 90 → 85
            result = form_row.to_dict()
            result["Match Score"] = best_score
            result["Matched Address"] = best_match["Address_Line1"]
            result["AC_NUM"] = best_match["AC_NUM"]
            result["AC_Name"] = best_match["AC_Name"]
            result["Address_Type"] = "01" if form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "").strip().lower() == "yes" else ""
            matched_rows.append(result)
        else:
            result = form_row.to_dict()
            result["Unmatched Reason"] = "No close address match found in UPS system."
            unmatched_rows.append(result)

    # Convert matched to DataFrame
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)

    # Show result counts
    st.success(f"✅ Matching complete. {len(matched_df)} matched, {len(unmatched_df)} unmatched.")

    # Export buttons
    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button("📥 Download Matched File", convert_df_to_excel(matched_df), file_name="matched.xlsx")
    with col2:
        st.download_button("📥 Download Unmatched File", convert_df_to_excel(unmatched_df), file_name="unmatched.xlsx")
    with col3:
        # Upload template generation
        if not matched_df.empty:
            upload_df = pd.DataFrame({
                "AC_NUM": matched_df["AC_NUM"],
                "AC_Address_Type": matched_df["Address_Type"],
                "AC_Name": matched_df["AC_Name"],
                "Address_Line1": matched_df["New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only"],
                "Address_Line2": matched_df["New Address Line 2 (Street Name)-In English Only"],
                "City": matched_df["City / Province"],
                "Postal_Code": "",
                "Country_Code": "VN",
                "Attention_Name": matched_df["Full Name of Contact-In English Only"],
                "Address_Line22": matched_df["New Address Line 3 (Ward/Commune)-In English Only"],
                "Address_Country_Code": "VN"
            })
            st.download_button("📥 Download Upload Template", convert_df_to_excel(upload_df), file_name="upload_template.xlsx")

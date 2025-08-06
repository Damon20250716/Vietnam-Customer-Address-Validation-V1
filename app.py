import streamlit as st
import pandas as pd
import unicodedata
from rapidfuzz import fuzz
from io import BytesIO

st.set_page_config(layout="wide")
st.title("🇻🇳 Vietnam Address Validation Tool")

# Function to remove Vietnamese tones
def remove_tones(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFD', text)
    return ''.join(c for c in text if unicodedata.category(c) != 'Mn')

# Normalize strings (remove tones, lowercase, strip spaces)
def normalize(text):
    if not isinstance(text, str):
        return ""
    text = remove_tones(text).lower().strip()
    return text.replace(",", "").replace(".", "")

# Check if address lines match
def address_match(form_line1, form_line2, ups_line1, ups_line2):
    form1 = normalize(form_line1)
    form2 = normalize(form_line2)
    ups1 = normalize(ups_line1)
    ups2 = normalize(ups_line2)

    # Direct line1/line2 match
    score1 = fuzz.partial_ratio(form1, ups1)
    score2 = fuzz.partial_ratio(form2, ups2)

    # Combined form lines to match UPS line 1
    combined_form = normalize(form_line1 + " " + form_line2)
    score_combined = fuzz.partial_ratio(combined_form, ups1)

    return (score1 > 85 and score2 > 85) or (score_combined > 85)

# File upload
forms_file = st.file_uploader("Upload Microsoft Forms Response", type=["xlsx"])
ups_file = st.file_uploader("Upload UPS System Export File", type=["xlsx"])

if forms_file and ups_file:
    with st.spinner("Processing..."):
        forms_df = pd.read_excel(forms_file)
        ups_df = pd.read_excel(ups_file)

        matched_rows = []
        unmatched_rows = []

        for _, form_row in forms_df.iterrows():
            acct = str(form_row.get("Account Number")).strip()
            form_line1 = str(form_row.get("New Address Line 1", "")).strip()
            form_line2 = str(form_row.get("New Address Line 2", "")).strip()
            form_city = str(form_row.get("City / Province", "")).strip()

            found_match = False
            ups_candidates = ups_df[ups_df["AC_NUM"].astype(str).str.strip() == acct]

            for _, ups_row in ups_candidates.iterrows():
                ups_line1 = str(ups_row.get("Address_Line1", "")).strip()
                ups_line2 = str(ups_row.get("Address_Line2", "")).strip()
                ups_city = str(ups_row.get("City", "")).strip()

                if address_match(form_line1, form_line2, ups_line1, ups_line2):
                    matched = form_row.to_dict()
                    matched["Tone-Free Line1"] = remove_tones(form_line1)
                    matched["Tone-Free Line2"] = remove_tones(form_line2)
                    matched_rows.append(matched)
                    found_match = True
                    break

            if not found_match:
                reason = "No address match found for account"
                unmatched = form_row.to_dict()
                unmatched["Unmatched Reason"] = reason
                unmatched_rows.append(unmatched)

        # Create output DataFrames
        matched_df = pd.DataFrame(matched_rows)
        unmatched_df = pd.DataFrame(unmatched_rows)

        # Show results
        st.success(f"✅ Matching complete. {len(matched_df)} matched, {len(unmatched_df)} unmatched.")
        st.download_button("📥 Download Matched File", BytesIO(matched_df.to_csv(index=False).encode()), "matched.csv", "text/csv")
        st.download_button("📥 Download Unmatched File", BytesIO(unmatched_df.to_csv(index=False).encode()), "unmatched.csv", "text/csv")

import streamlit as st
import pandas as pd
import unicodedata
import difflib
from io import BytesIO

# --- Utility functions ---

def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text

def normalize(text):
    if not isinstance(text, str):
        return ''
    return remove_tones(text.strip().lower())

def clean_address(addr):
    return normalize(addr)

def flexible_address_match(addr1, addr2, threshold=0.5):
    a1 = clean_address(addr1)
    a2 = clean_address(addr2)
    if not a1 or not a2:
        return False

    # Substring match
    if a1 in a2 or a2 in a1:
        return True

    # Token overlap
    words1 = set(a1.split())
    words2 = set(a2.split())
    if words1 and words2:
        intersection = words1.intersection(words2)
        match_ratio = len(intersection) / max(len(words1), len(words2))
        if match_ratio >= 0.6:
            return True

    # Fuzzy match
    ratio = difflib.SequenceMatcher(None, a1, a2).ratio()
    return ratio >= threshold

# --- App UI ---
st.title("Vietnam Address Validation Tool")
st.markdown("Upload Microsoft Forms responses and UPS system data")

forms_file = st.file_uploader("Upload Forms Responses Excel", type="xlsx")
system_file = st.file_uploader("Upload UPS System Address Excel", type="xlsx")
thresh = st.slider("Address Matching Similarity Threshold", 0.0, 1.0, 0.5, 0.01)

if forms_file and system_file:
    forms_df = pd.read_excel(forms_file)
    system_df = pd.read_excel(system_file)

    # Normalize system data
    system_df['Normalized Line 1'] = system_df['Address Line 1'].apply(clean_address)
    system_df['Normalized Line 2'] = system_df['Address Line 2'].apply(clean_address)

    matched_rows = []
    unmatched_rows = []

    for _, row in forms_df.iterrows():
        acct = str(row['Account Number']).strip()
        new_line1 = clean_address(row['New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
        new_line2 = clean_address(row['New Address Line 2 (Street Name)-In English Only'])

        # Filter system by account
        sys_acct_df = system_df[system_df['AC_NUM'].astype(str).str.strip() == acct]
        found = False

        for _, sys_row in sys_acct_df.iterrows():
            sys_line1 = sys_row['Normalized Line 1']
            sys_line2 = sys_row['Normalized Line 2']

            if flexible_address_match(new_line1, sys_line1, threshold=thresh) and \
               flexible_address_match(new_line2, sys_line2, threshold=thresh):
                matched_rows.append(row.to_dict())
                found = True
                break

        if not found:
            row_dict = row.to_dict()
            row_dict['Unmatched Reason'] = 'No matching address found'
            unmatched_rows.append(row_dict)

    # Output DataFrames
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)

    st.success(f"✅ Matching complete. {len(matched_df)} matched, {len(unmatched_df)} unmatched.")

    with st.expander("📥 Download Matched Results"):
        buf = BytesIO()
        matched_df.to_excel(buf, index=False)
        st.download_button("Download Matched File", data=buf.getvalue(), file_name="matched.xlsx")

    with st.expander("📤 Download Unmatched Results"):
        buf2 = BytesIO()
        unmatched_df.to_excel(buf2, index=False)
        st.download_button("Download Unmatched File", data=buf2.getvalue(), file_name="unmatched.xlsx")

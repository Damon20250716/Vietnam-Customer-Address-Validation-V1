
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")

def normalize_address(addr):
    return str(addr).strip().lower() if pd.notna(addr) else ""

def validate_addresses(forms_df, ups_df):
    matched = []
    unmatched = []

    # Try to locate key column dynamically
    col_match = [col for col in forms_df.columns if "is your new billing address" in col.lower()]
    if not col_match:
        st.error("Key question column not found in Forms file. Please check the column header.")
        return [], [], []
    billing_same_col = col_match[0]

    # Normalize UPS data for faster lookup
    ups_lookup = {}
    for _, row in ups_df.iterrows():
        acc = str(row['Account Number']).strip()
        addr1 = normalize_address(row['Address Line 1'])
        addr2 = normalize_address(row['Address Line 2'])
        addr3 = normalize_address(row['Address Line 3'])
        key = (acc, addr1, addr2, addr3)
        ups_lookup[key] = row

    for _, row in forms_df.iterrows():
        acc = str(row['Account Number']).strip()
        answer = str(row[billing_same_col]).strip().lower()

        addresses = []

        if answer == "yes":
            addr1 = normalize_address(row['New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
            addr2 = normalize_address(row['New Address Line 2 (Street Name)-In English Only'])
            addr3 = normalize_address(row['New Address Line 3 (Ward/Commune)-In English Only'])
            addresses.append((acc, addr1, addr2, addr3, "01"))
        else:
            # Billing
            b1 = normalize_address(row['New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
            b2 = normalize_address(row['New Billing Address Line 2 (Street Name)-In English Only'])
            b3 = normalize_address(row['New Billing Address Line 3 (Ward/Commune)-In English Only'])
            addresses.append((acc, b1, b2, b3, "03"))
            # Delivery
            d1 = normalize_address(row['New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
            d2 = normalize_address(row['New Delivery Address Line 2 (Street Name)-In English Only'])
            d3 = normalize_address(row['New Delivery Address Line 3 (Ward/Commune)-In English Only'])
            addresses.append((acc, d1, d2, d3, "13"))
            # Pickup (up to 3)
            for i in range(1, 4):
                col_prefix = f'{["First", "Second", "Third"][i-1]} New Pick Up Address'
                if f'{col_prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only' in row:
                    p1 = normalize_address(row[f'{col_prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
                    p2 = normalize_address(row[f'{col_prefix} Line 2 (Street Name)-In English Only'])
                    p3 = normalize_address(row[f'{col_prefix} Line 3 (Ward/Commune)-In English Only'])
                    if p1 or p2 or p3:
                        addresses.append((acc, p1, p2, p3, "02"))

        matched_rows = []
        for acc, a1, a2, a3, addr_type in addresses:
            key = (acc, a1, a2, a3)
            if key in ups_lookup:
                matched_rows.append({
                    "Account Number": acc,
                    "Address Type": addr_type,
                    "Address Line 1": a1,
                    "Address Line 2": a2,
                    "Address Line 3": a3,
                })
            else:
                unmatched.append({
                    "Account Number": acc,
                    "Address Type": addr_type,
                    "Address Line 1": a1,
                    "Address Line 2": a2,
                    "Address Line 3": a3,
                })

        if matched_rows:
            matched.extend(matched_rows)
        elif addresses:
            unmatched.append({
                "Account Number": acc,
                "Address Type": "N/A",
                "Address Line 1": "",
                "Address Line 2": "",
                "Address Line 3": "",
            })

    return matched, unmatched, convert_to_upload_template(matched)

def convert_to_upload_template(matched_rows):
    upload_rows = []
    for row in matched_rows:
        upload_rows.append({
            "Account Number": row["Account Number"],
            "Address Type": row["Address Type"],
            "Address Line 1": row["Address Line 1"],
            "Address Line 2": row["Address Line 2"],
            "Address Line 3": row["Address Line 3"],
        })
    return upload_rows

st.title("🇻🇳 Vietnam Customer Address Validation Tool")

forms_file = st.file_uploader("Upload Microsoft Forms Response Excel", type=['xlsx'])
ups_file = st.file_uploader("Upload UPS existing system info Excel file", type=['xlsx'])

if forms_file and ups_file:
    try:
        forms_df = pd.read_excel(forms_file)
        ups_df = pd.read_excel(ups_file)

        matched, unmatched, upload = validate_addresses(forms_df, ups_df)

        st.success(f"✅ Matched: {len(matched)} | ❌ Unmatched: {len(unmatched)}")

        with pd.ExcelWriter("vietnam_address_validation_output.xlsx", engine="xlsxwriter") as writer:
            pd.DataFrame(matched).to_excel(writer, sheet_name="Matched", index=False)
            pd.DataFrame(unmatched).to_excel(writer, sheet_name="Unmatched", index=False)
            pd.DataFrame(upload).to_excel(writer, sheet_name="Upload Template", index=False)
        with open("vietnam_address_validation_output.xlsx", "rb") as f:
            st.download_button("📥 Download Output Excel", f, file_name="vietnam_address_validation_output.xlsx")

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")

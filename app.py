import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------ Utility Functions ------------------

def normalize_col(col):
    return col.astype(str).str.lower().str.strip()

def load_excel_file(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '', regex=False)
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ------------------ Matching Logic ------------------

def process_forms_and_ups(forms_df, ups_df):
    matched, unmatched, upload_template = [], [], []
    
    forms_df.fillna("", inplace=True)
    ups_df.fillna("", inplace=True)
    
    ups_df["Account Number"] = normalize_col(ups_df["Account Number"].astype(str))
    
    for _, row in forms_df.iterrows():
        acc = str(row.get("Account Number", "")).strip()
        if not acc:
            continue

        is_same = str(row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()
        email = row.get("Please Provide Your Email Address-In English Only", "")
        contact = row.get("Full Name of Contact-In English Only", "")
        phone = row.get("Contact Phone Number", "")
        city = row.get("City / Province", "")

        if is_same == "yes":
            addr1 = row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            addr2 = row.get("New Address Line 2 (Street Name)-In English Only", "")
            addr3 = row.get("New Address Line 3 (Ward/Commune)-In English Only", "")

            for code in ["1", "2", "6"]:
                upload_template.append({
                    "Account Number": acc,
                    "Address Type": code,
                    "Address Line 1": addr1,
                    "Address Line 2": addr2,
                    "Address Line 3": addr3,
                    "City": city,
                    "Email": email,
                    "Contact Name": contact,
                    "Phone": phone
                })

            matched.append({
                "Account Number": acc,
                "Same for All": "Yes",
                "Line 1": addr1,
                "Line 2": addr2,
                "Line 3": addr3
            })

        else:
            ups_match = ups_df[ups_df["Account Number"] == acc]
            pickup_count = (ups_match["Address Type"] == "02").sum()

            def get_addr(prefix):
                return [
                    row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                    row.get(f"{prefix} Line 2 (Street Name)-In English Only", ""),
                    row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                ]

            def add_entry(code, addr):
                upload_template.append({
                    "Account Number": acc,
                    "Address Type": code,
                    "Address Line 1": addr[0],
                    "Address Line 2": addr[1],
                    "Address Line 3": addr[2],
                    "City": city,
                    "Email": email,
                    "Contact Name": contact,
                    "Phone": phone
                })

            add_entry("03", get_addr("New Billing Address"))
            add_entry("13", get_addr("New Delivery Address"))

            for i in range(1, pickup_count + 1):
                prefix = f"{i}st New Pick Up Address"
                add_entry("02", get_addr(prefix))

            matched.append({"Account Number": acc, "Same for All": "No"})

    matched_accs = {row["Account Number"] for row in matched}
    for _, row in forms_df.iterrows():
        if str(row.get("Account Number", "")).strip() not in matched_accs:
            unmatched.append(row)

    return pd.DataFrame(matched), pd.DataFrame(unmatched), pd.DataFrame(upload_template)

# ------------------ Streamlit UI ------------------

def main():
    st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")
    st.title("🇻🇳 Vietnam Address Validation Tool")

    forms_file = st.file_uploader("Upload Microsoft Forms Response File", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File", type=["xlsx"])

    if forms_file and ups_file:
        try:
            with st.spinner("Processing..."):
                forms_df = load_excel_file(forms_file)
                ups_df = load_excel_file(ups_file)

                matched_df, unmatched_df, upload_df = process_forms_and_ups(forms_df, ups_df)

                st.success("✅ Processing Complete!")
                st.write(f"🔗 Matched: {len(matched_df)}")
                st.write(f"🔍 Unmatched: {len(unmatched_df)}")

                st.download_button("📥 Download Matched File", to_excel(matched_df), "matched.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.download_button("📥 Download Unmatched File", to_excel(unmatched_df), "unmatched.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.download_button("📥 Download Upload Template", to_excel(upload_df), "upload_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"❌ An error occurred: {e}")

if __name__ == "__main__":
    main()

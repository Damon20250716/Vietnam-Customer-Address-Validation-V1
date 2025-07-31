import streamlit as st
import pandas as pd

# ------------------ Utility Functions ------------------

def normalize_col(col):
    return col.astype(str).str.lower().str.strip()

def load_excel_file(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '', regex=False)
    return df

# ------------------ Matching Logic ------------------

def is_address_equal(addr1, addr2):
    return normalize_col(addr1) == normalize_col(addr2)

def process_forms_and_ups(forms_df, ups_df):
    matched = []
    unmatched = []
    upload_template = []

    for _, form_row in forms_df.iterrows():
        acc = str(form_row.get("Account Number", "")).strip()
        is_same = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        addr1 = form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
        addr2 = form_row.get("New Address Line 2 (Street Name)-In English Only", "")
        addr3 = form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
        city = form_row.get("City / Province", "")
        email = form_row.get("Please Provide Your Email Address-In English Only", "")
        contact = form_row.get("Full Name of Contact-In English Only", "")
        phone = form_row.get("Contact Phone Number", "")

        if is_same == "yes":
            # Treat as address type 01 (all)
            matched.append({
                "Account Number": acc,
                "Address Type": "01",
                "Address Line 1": addr1,
                "Address Line 2": addr2,
                "Address Line 3": addr3,
                "City": city,
                "Email": email,
                "Contact Name": contact,
                "Phone": phone
            })
            # Upload template needs 3 rows with address types 1, 2, 6
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
        else:
            # Lookup pickup address count in UPS file
            ups_matches = ups_df[normalize_col(ups_df["Account Number"]) == acc]
            pickup_count = (ups_matches["Address Type"] == "02").sum()

            if pickup_count == 1:
                matched.extend([
                    {
                        "Account Number": acc,
                        "Address Type": "02",
                        "Address Line 1": form_row.get("New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)-In Vietnamese without accents", ""),
                        "Address Line 2": form_row.get("New Pick Up Address Line 2 (Street Name)-In Vietnamese without tone marks", ""),
                        "Address Line 3": form_row.get("New Pick Up Address Line 3 (Ward/Commune)-In Vietnamese without tone marks", ""),
                        "City": form_row.get("City / Province", ""),
                        "Email": email,
                        "Contact Name": contact,
                        "Phone": phone
                    },
                    {
                        "Account Number": acc,
                        "Address Type": "03",
                        "Address Line 1": form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                        "Address Line 2": form_row.get("New Billing Address Line 2 (Street Name)-In English Only", ""),
                        "Address Line 3": form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", ""),
                        "City": form_row.get("City / Province", ""),
                        "Email": email,
                        "Contact Name": contact,
                        "Phone": phone
                    },
                    {
                        "Account Number": acc,
                        "Address Type": "13",
                        "Address Line 1": form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                        "Address Line 2": form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", ""),
                        "Address Line 3": form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", ""),
                        "City": form_row.get("City / Province", ""),
                        "Email": email,
                        "Contact Name": contact,
                        "Phone": phone
                    }
                ])
            else:
                # Multiple pickup addresses: create rows per pickup
                for i in range(1, pickup_count + 1):
                    matched.append({
                        "Account Number": acc,
                        "Address Type": "02",
                        "Address Line 1": form_row.get(f"{i}st New Pick Up Address Line 1", ""),
                        "Address Line 2": form_row.get(f"{i}st New Pick Up Address Line 2", ""),
                        "Address Line 3": form_row.get(f"{i}st New Pick Up Address Line 3", ""),
                        "City": city,
                        "Email": email,
                        "Contact Name": contact,
                        "Phone": phone
                    })
                # Add billing & delivery
                matched.extend([
                    {
                        "Account Number": acc,
                        "Address Type": "03",
                        "Address Line 1": form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                        "Address Line 2": form_row.get("New Billing Address Line 2 (Street Name)-In English Only", ""),
                        "Address Line 3": form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", ""),
                        "City": city,
                        "Email": email,
                        "Contact Name": contact,
                        "Phone": phone
                    },
                    {
                        "Account Number": acc,
                        "Address Type": "13",
                        "Address Line 1": form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                        "Address Line 2": form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", ""),
                        "Address Line 3": form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", ""),
                        "City": city,
                        "Email": email,
                        "Contact Name": contact,
                        "Phone": phone
                    }
                ])

    # Forms accounts that never matched
    matched_accs = {row["Account Number"] for row in matched}
    for _, row in forms_df.iterrows():
        if str(row.get("Account Number", "")).strip() not in matched_accs:
            unmatched.append(row)

    return matched, unmatched, upload_template

# ------------------ Streamlit UI ------------------

def main():
    st.title("🇻🇳 Vietnam Address Validation Tool")

    forms_file = st.file_uploader("Upload Microsoft Forms Response File", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File", type=["xlsx"])

    if forms_file and ups_file:
        try:
            with st.spinner("Processing..."):
                forms_df = load_excel_file(forms_file)
                ups_df = load_excel_file(ups_file)

                st.subheader("🔍 Debug Info")
                st.write("Forms Columns:", forms_df.columns.tolist())
                st.write("UPS Columns:", ups_df.columns.tolist())

                matched, unmatched, template = process_forms_and_ups(forms_df, ups_df)

                matched_df = pd.DataFrame(matched)
                unmatched_df = pd.DataFrame(unmatched)
                upload_df = pd.DataFrame(template)

                st.success("✅ Processing Complete!")
                st.download_button("📥 Download Matched File", matched_df.to_excel(index=False), "matched.xlsx")
                st.download_button("📥 Download Unmatched File", unmatched_df.to_excel(index=False), "unmatched.xlsx")
                st.download_button("📥 Download Upload Template", upload_df.to_excel(index=False), "upload_template.xlsx")

        except Exception as e:
            st.error(f"❌ An error occurred: {e}")

if __name__ == "__main__":
    main()

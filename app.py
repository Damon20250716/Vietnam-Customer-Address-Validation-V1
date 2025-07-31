import streamlit as st
import pandas as pd

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")

st.title("🇻🇳 Vietnam Address Validation Tool")
st.write("Validate customer-submitted addresses from Microsoft Forms against UPS system records.")

def normalize_col(s):
    return s.astype(str).str.strip().str.lower().str.replace(" ", "")

def prepare_upload_template(df):
    # Expand billing addresses (code '03') into three codes: 1,2,6
    billing = df[df["Address Type"] == "03"].copy()
    others = df[df["Address Type"] != "03"].copy()

    billing_expanded = billing.loc[billing.index.repeat(3)].copy()
    billing_expanded["Code"] = [1, 2, 6] * len(billing)

    # For other address types, just map Address Type to Code directly
    others["Code"] = others["Address Type"].astype(str)

    # Concatenate billing expanded and others
    upload_df = pd.concat([billing_expanded, others], ignore_index=True)

    # Rename columns for upload format
    upload_df = upload_df.rename(columns={
        "Account Number": "Account Number",
        "New Address Line 1": "Line 1",
        "New Address Line 2": "Line 2",
        "New Address Line 3": "Line 3",
        "Code": "Code"
    })[["Account Number", "Code", "Line 1", "Line 2", "Line 3"]]

    return upload_df

def main():
    forms_file = st.file_uploader("Upload Microsoft Forms Response File", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File", type=["xlsx"])

    if forms_file and ups_file:
        forms_df = pd.read_excel(forms_file)
        ups_df = pd.read_excel(ups_file)

        # Normalize account numbers and address lines for comparison
        forms_df["Account Number"] = forms_df["Account Number"].astype(str).str.strip().str.lower()
        ups_df["Account Number"] = ups_df["Account Number"].astype(str).str.strip().str.lower()

        ups_df["Address Line 1_norm"] = normalize_col(ups_df["Address Line 1"])
        ups_df["Address Line 2_norm"] = normalize_col(ups_df["Address Line 2"])

        matched_rows = []
        unmatched_rows = []

        # Create dict of UPS addresses grouped by account and address type for quick lookup
        ups_grouped = ups_df.groupby("Account Number")

        for idx, form_row in forms_df.iterrows():
            acc = form_row["Account Number"]
            if acc not in ups_grouped.groups:
                unmatched_rows.append(form_row)
                continue

            ups_account_df = ups_grouped.get_group(acc)

            # Check if billing same as pickup and delivery
            same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower() == "yes"

            if same_billing:
                # Unified address, address type 01
                new_addr1 = form_row["New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only"]
                new_addr2 = form_row["New Address Line 2 (Street Name)-In English Only"]
                new_addr1_norm = str(new_addr1).strip().lower().replace(" ", "")
                new_addr2_norm = str(new_addr2).strip().lower().replace(" ", "")

                # Check if any UPS address line1 & line2 matches
                found = False
                for _, ups_addr in ups_account_df.iterrows():
                    if (ups_addr["Address Line 1_norm"] == new_addr1_norm and
                        ups_addr["Address Line 2_norm"] == new_addr2_norm):
                        matched_rows.append({
                            "Account Number": acc,
                            "Address Type": "01",
                            "New Address Line 1": new_addr1,
                            "New Address Line 2": new_addr2,
                            "New Address Line 3": form_row["New Address Line 3 (Ward/Commune)-In English Only"]
                        })
                        found = True
                        break
                if not found:
                    unmatched_rows.append(form_row)
            else:
                # When not same billing, check Billing(03), Delivery(13), Pickups(02)
                matched_flag = False
                # Define the prefixes for the different address types
                addr_map = {
                    "03": "New Billing Address",
                    "13": "New Delivery Address",
                    "02": ["First New Pick Up Address", "Second New Pick Up Address", "Third New Pick Up Address"]
                }

                # Check Billing and Delivery
                for addr_type in ["03", "13"]:
                    prefix = addr_map[addr_type]
                    new_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                    new_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                    if not new_addr1.strip():
                        continue
                    new_addr1_norm = str(new_addr1).strip().lower().replace(" ", "")
                    new_addr2_norm = str(new_addr2).strip().lower().replace(" ", "")

                    ups_sub = ups_account_df[ups_account_df["Address Type"] == addr_type]
                    found = False
                    for _, ups_addr in ups_sub.iterrows():
                        if (ups_addr["Address Line 1_norm"] == new_addr1_norm and
                            ups_addr["Address Line 2_norm"] == new_addr2_norm):
                            matched_rows.append({
                                "Account Number": acc,
                                "Address Type": addr_type,
                                "New Address Line 1": new_addr1,
                                "New Address Line 2": new_addr2,
                                "New Address Line 3": form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                            })
                            found = True
                            matched_flag = True
                            break
                    if not found:
                        unmatched_rows.append(form_row)

                # Check pickups (can be up to 3)
                ups_pickups = ups_account_df[ups_account_df["Address Type"] == "02"]
                pickup_prefixes = addr_map["02"]
                pickups_found = 0

                for prefix in pickup_prefixes:
                    new_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                    if not new_addr1.strip():
                        continue
                    new_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                    new_addr1_norm = str(new_addr1).strip().lower().replace(" ", "")
                    new_addr2_norm = str(new_addr2).strip().lower().replace(" ", "")

                    found = False
                    for _, ups_addr in ups_pickups.iterrows():
                        if (ups_addr["Address Line 1_norm"] == new_addr1_norm and
                            ups_addr["Address Line 2_norm"] == new_addr2_norm):
                            matched_rows.append({
                                "Account Number": acc,
                                "Address Type": "02",
                                "New Address Line 1": new_addr1,
                                "New Address Line 2": new_addr2,
                                "New Address Line 3": form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                            })
                            found = True
                            pickups_found += 1
                            matched_flag = True
                            break
                    if not found:
                        unmatched_rows.append(form_row)

                # Verify number of pickup addresses match between Forms and UPS
                if pickups_found != len(ups_pickups):
                    unmatched_rows.append(form_row)

        # Prepare DataFrames to return
        matched_df = pd.DataFrame(matched_rows)
        unmatched_df = pd.DataFrame(unmatched_rows)

        # Prepare upload template DataFrame (expanding billing)
        upload_template_df = prepare_upload_template(matched_df)

        st.success(f"Processed: {len(matched_df)} matched rows, {len(unmatched_df)} unmatched rows.")

        st.download_button("Download Matched CSV", data=matched_df.to_csv(index=False), file_name="matched.csv", mime="text/csv")
        st.download_button("Download Unmatched CSV", data=unmatched_df.to_csv(index=False), file_name="unmatched.csv", mime="text/csv")
        st.download_button("Download Upload Template CSV", data=upload_template_df.to_csv(index=False), file_name="upload_template.csv", mime="text/csv")

if __name__ == "__main__":
    main()

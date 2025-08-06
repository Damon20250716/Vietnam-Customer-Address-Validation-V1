# Vietnam Address Validation Tool

import pandas as pd
import unicodedata
import re
from difflib import SequenceMatcher
from io import BytesIO

# ---------------------- Helper Functions ----------------------

def remove_vietnamese_tones(text):
    if pd.isnull(text):
        return ""
    text = str(text)
    text = unicodedata.normalize('NFD', text)
    text = ''.join([c for c in text if unicodedata.category(c) != 'Mn'])
    text = re.sub(r'[^A-Za-z0-9\s,.-]', '', text)
    return text.lower().strip()

def address_similarity(addr1, addr2):
    addr1 = remove_vietnamese_tones(addr1)
    addr2 = remove_vietnamese_tones(addr2)
    return SequenceMatcher(None, addr1, addr2).ratio()


def combine_address(row, prefix):
    parts = [row.get(f"{prefix} Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
             row.get(f"{prefix} Address Line 2 (Street Name)-In English Only", ""),
             row.get(f"{prefix} Address Line 3 (Ward/Commune)-In English Only", "")]
    return ", ".join(filter(None, parts))

# ---------------------- Main Validation Logic ----------------------

def validate_addresses(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    forms_df['Account Number'] = forms_df['Account Number'].astype(str).str.strip()
    ups_df['Account Number'] = ups_df['Account Number'].astype(str).str.strip()

    ups_grouped = ups_df.groupby("Account Number")

    for _, form_row in forms_df.iterrows():
        acc = str(form_row['Account Number']).strip()
        if acc not in ups_grouped.groups:
            unmatched_rows.append(form_row)
            continue

        ups_rows = ups_grouped.get_group(acc)
        billing_same = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower() == 'yes'

        if billing_same:
            addr_form = combine_address(form_row, "New")
            match_found = False
            for _, ups_row in ups_rows.iterrows():
                if ups_row['Address Type'] != '01':
                    continue
                addr_ups = remove_vietnamese_tones(str(ups_row.get('Address Line 1', '')))
                sim_score = address_similarity(addr_form, addr_ups)
                if sim_score >= 0.65:
                    match_found = True
                    matched_rows.append(form_row)
                    upload_template_rows.extend([{
                        'Account Number': acc,
                        'Customer Name': form_row.get('Full Name of Contact-In English Only', ''),
                        'Change Type': change_type,
                        'New Address Line 1': form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                        'New Address Line 2': form_row.get("New Address Line 2 (Street Name)-In English Only", ""),
                        'New Address Line 3': form_row.get("New Address Line 3 (Ward/Commune)-In English Only", ""),
                        'New City': form_row.get("City / Province", "")
                    } for change_type in [1, 2, 6]])
                    break

            if not match_found:
                unmatched_rows.append(form_row)

        else:
            match_found = False
            for addr_type, ups_code in zip(["New Billing", "New Delivery"], ["03", "13"]):
                addr_form = combine_address(form_row, addr_type)
                for _, ups_row in ups_rows.iterrows():
                    if ups_row['Address Type'] != ups_code:
                        continue
                    addr_ups = remove_vietnamese_tones(str(ups_row.get('Address Line 1', '')))
                    sim_score = address_similarity(addr_form, addr_ups)
                    if sim_score >= 0.65:
                        match_found = True
                        matched_rows.append(form_row)
                        upload_template_rows.append({
                            'Account Number': acc,
                            'Customer Name': form_row.get('Full Name of Contact-In English Only', ''),
                            'Change Type': int(ups_code),
                            'New Address Line 1': form_row.get(f"{addr_type} Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                            'New Address Line 2': form_row.get(f"{addr_type} Address Line 2 (Street Name)-In English Only", ""),
                            'New Address Line 3': form_row.get(f"{addr_type} Address Line 3 (Ward/Commune)-In English Only", ""),
                            'New City': form_row.get(f"{addr_type} City / Province", "")
                        })
                        break

            for i in range(1, 4):
                addr_type = f"First New Pick Up" if i == 1 else f"Second New Pick Up" if i == 2 else f"Third New Pick Up"
                addr_form = combine_address(form_row, addr_type)
                for _, ups_row in ups_rows.iterrows():
                    if ups_row['Address Type'] != '02':
                        continue
                    addr_ups = remove_vietnamese_tones(str(ups_row.get('Address Line 1', '')))
                    sim_score = address_similarity(addr_form, addr_ups)
                    if sim_score >= 0.65:
                        matched_rows.append(form_row)
                        upload_template_rows.append({
                            'Account Number': acc,
                            'Customer Name': form_row.get('Full Name of Contact-In English Only', ''),
                            'Change Type': 2,
                            'New Address Line 1': form_row.get(f"{addr_type} Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                            'New Address Line 2': form_row.get(f"{addr_type} Address Line 2 (Street Name)-In English Only", ""),
                            'New Address Line 3': form_row.get(f"{addr_type} Address Line 3 (Ward/Commune)-In English Only", ""),
                            'New City': form_row.get(f"{addr_type} City / Province", "")
                        })
                        match_found = True
                        break

            if not match_found:
                unmatched_rows.append(form_row)

    matched_df = pd.DataFrame(matched_rows).drop_duplicates(subset=['Account Number'])
    unmatched_df = pd.DataFrame(unmatched_rows).drop_duplicates(subset=['Account Number'])
    upload_template_df = pd.DataFrame(upload_template_rows)

    return matched_df, unmatched_df, upload_template_df

# ---------------------- End of Code ----------------------

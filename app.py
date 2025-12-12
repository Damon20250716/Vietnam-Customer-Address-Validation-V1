import streamlit as st
import pandas as pd
from unidecode import unidecode

st.title("Vietnam Address Transformation Tool")
st.write("Upload MS Forms file and Testing template to generate final output.")

# Remove diacritics for Vietnamese text
def remove_diacritics(text):
    if pd.isna(text):
        return ""
    return unidecode(str(text)).upper().strip()

# ========================
# Main transformation logic
# ========================
def process_files(forms_df, template_df):

    output_rows = []  # list to collect all generated rows

    for _, row in forms_df.iterrows():

        account = row["Account Number"]

        same_addr = str(row["Column B"]).strip().lower() == "yes"

        # Case 1: Address = Same (Only 1 row in Testing)
        if same_addr:
            new_row = template_df.iloc[0].copy()
            new_row["Account Number"] = account
            new_row["Addr Type"] = "01"

            new_row["New Address Line 1"] = remove_diacritics(row["Column C"])
            new_row["New Address Line 2"] = remove_diacritics(row["Column D"])
            new_row["New Address Line 3 - Ward/Commune"] = remove_diacritics(row["Column E"])
            new_row["New Five Digits Postal Code"] = remove_diacritics(row["Column G"])

            output_rows.append(new_row)

        # Case 2: Address = NOT same → create billing + delivery + pickup(s)
        else:
            # -----------------------------
            # (1) Billing Address (Addr Type 03)
            # -----------------------------
            bill = template_df.iloc[0].copy()
            bill["Account Number"] = account
            bill["Addr Type"] = "03"

            bill["New Billing Address Line 1"] = remove_diacritics(row["Column H"])
            bill["New Billing Address Line 2"] = remove_diacritics(row["Column I"])
            bill["New Billing Address Line 3"] = remove_diacritics(row["Column J"])
            bill["New Billing Address Line 4"] = remove_diacritics(row["Column K"])
            bill["New Five Digits Postal Code2"] = remove_diacritics(row["Column L"])

            output_rows.append(bill)

            # -----------------------------
            # (2) Delivery Address (Addr Type 13)
            # -----------------------------
            deli = template_df.iloc[0].copy()
            deli["Account Number"] = account
            deli["Addr Type"] = "13"

            deli["New Delivery Address Line 1"] = remove_diacritics(row["Column M"])
            deli["New Delivery Address Line 2"] = remove_diacritics(row["Column N"])
            deli["New Delivery Address Line 3"] = remove_diacritics(row["Column O"])
            deli["New Delivery Address Line 4"] = remove_diacritics(row["Column P"])
            deli["New Five Digits Postal Code3"] = remove_diacritics(row["Column Q"])

            output_rows.append(deli)

            # -----------------------------
            # (3) Pickup Address — dynamic rows
            # -----------------------------
            pickup_count = str(row["Column R"]).strip().lower()

            def create_pickup_row(source_cols):
                pr = template_df.iloc[0].copy()
                pr["Account Number"] = account
                pr["Addr Type"] = "02"
                pr["New Pickup Address Line 1"] = remove_diacritics(row[source_cols[0]])
                pr["New Pickup Address Line 2"] = remove_diacritics(row[source_cols[1]])
                pr["New Pickup Address Line 3"] = remove_diacritics(row[source_cols[2]])
                pr["New Pickup Address Line 4"] = remove_diacritics(row[source_cols[3]])
                pr["New Five Digits Postal Code4"] = remove_diacritics(row[source_cols[4]])
                output_rows.append(pr)

            if pickup_count == "one":
                create_pickup_row(["Column S","Column T","Column U","Column V","Column W"])

            elif pickup_count == "two":
                create_pickup_row(["Column X","Column Y","Column Z","Column AA","Column AB"])
                create_pickup_row(["Column AC","Column AD","Column AE","Column AF","Column AG"])

            elif pickup_count == "three":
                create_pickup_row(["Column AH","Column AI","Column AJ","Column AK","Column AL"])
                create_pickup_row(["Column AM","Column AN","Column AO","Column AP","Column AQ"])
                create_pickup_row(["Column AR","Column AS","Column AT","Column AU","Column AV"])

    return pd.DataFrame(output_rows)


# ========================
# Streamlit UI
# ========================
forms_file = st.file_uploader("Upload MS Forms Excel", type=["xlsx"])
template_file = st.file_uploader("Upload Testing Template Excel", type=["xlsx"])

if forms_file and template_file:
    forms_df = pd.read_excel(forms_file)
    template_df = pd.read_excel(template_file)

    if st.button("Generate Testing File"):
        result = process_files(forms_df, template_df)

        st.success("Processing completed! Click below to download.")

        # Export file
        result_file = "Testing_Output.xlsx"
        result.to_excel(result_file, index=False)

        with open(result_file, "rb") as f:
            st.download_button(
                label="Download Testing File",
                data=f,
                file_name="Testing_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


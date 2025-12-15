import streamlit as st
import pandas as pd
import unicodedata

st.set_page_config(page_title="Vietnamese Accent Remover", layout="centered")

st.title("🇻🇳 Vietnamese Accent Remover (Excel)")
st.write("Upload an Excel file. All Vietnamese accents will be removed.")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"]
)

def remove_vietnamese_accents(text):
    if pd.isna(text):
        return text
    text = str(text)
    text = unicodedata.normalize('NFD', text)
    text = ''.join(
        char for char in text
        if unicodedata.category(char) != 'Mn'
    )
    return text.replace('Đ', 'D').replace('đ', 'd')

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # 只处理文本列（object 类型）
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(remove_vietnamese_accents)

        st.success("✅ Vietnamese accents removed successfully!")

        st.subheader("Preview")
        st.dataframe(df.head(20))

        output_file = "excel_without_vietnamese_accents.xlsx"
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download cleaned Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")


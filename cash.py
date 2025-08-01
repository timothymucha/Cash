import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Preview", layout="wide")
st.title("🧾 Excel Data Preview - From Row 17")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        # Read the Excel file starting from row 17 (index 16, zero-based)
        df_raw = pd.read_excel(uploaded_file, header=None, skiprows=16)

        # Display the first 10 rows for preview
        st.subheader("🔍 First 10 Rows of Raw Data (from row 17)")
        st.dataframe(df_raw.head(10))
        
        st.success("✅ File loaded and previewed successfully.")

    except Exception as e:
        st.error(f"❌ Failed to read file: {e}")

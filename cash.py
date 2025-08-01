import streamlit as st
import pandas as pd
from io import StringIO
from datetime import datetime

st.set_page_config(page_title="Mnarani Excel to QuickBooks IIF", layout="wide")
st.title("📄 Convert Mnarani Excel Statement to QuickBooks IIF")

uploaded_file = st.file_uploader("Upload Excel (.xlsx) File", type=["xlsx"])

def generate_iif(df):
    output = StringIO()

    # IIF headers
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        date = row["Bill Date"].strftime("%m/%d/%Y")
        docnum = row["Bill No."]
        amount = float(row["Amount"])
        memo = f"Till: {row['Till#']} - Bill No: {docnum}"
        name = "Walk In"

        output.write(f"TRNS\tRECEIPT\t{date}\tUndeposited Funds\t{name}\t{memo}\t{amount:.2f}\t{docnum}\n")
        output.write(f"SPL\tRECEIPT\t{date}\tRevenue:Cash Sales\t{name}\t{memo}\t{-amount:.2f}\t\t\n")
        output.write("ENDTRNS\n")

    return output.getvalue()

if uploaded_file:
    try:
        # Read Excel skipping to row 17 (skiprows=16)
        raw_df = pd.read_excel(uploaded_file, header=None, skiprows=16)

        # Extract relevant columns
        df = raw_df.iloc[:, [4, 9, 15, 25]]  # E, J, P, Z
        df.columns = ["Till#", "Bill Date", "Bill No.", "Amount"]

        # Stop when Till# becomes NaN (end of cash)
        df = df[df["Till#"].notna()]

        # Clean dates
        df["Bill Date"] = pd.to_datetime(df["Bill Date"], errors="coerce")
        df = df[df["Bill Date"].notna()]

        # Preview first 10 rows
        st.subheader("🔍 Data Preview (First 10 Rows)")
        st.dataframe(df.head(10), use_container_width=True)

        # Generate IIF file
        iif_data = generate_iif(df)
        st.subheader("⬇️ Download IIF")
        st.download_button("Download IIF File", iif_data, file_name="mnarani_cash_sales.iif", mime="text/plain")

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
else:
    st.info("📤 Upload an Excel file to begin.")

import streamlit as st
import pandas as pd
from io import StringIO
from datetime import datetime

st.set_page_config(page_title="Cash Sales to IIF Converter", layout="wide")
st.title("🧾 Convert Cash Sales Excel to QuickBooks IIF")

uploaded_file = st.file_uploader("Upload the Excel statement", type=["xlsx"])

def parse_custom_date(date_str):
    try:
        return datetime.strptime(date_str, "%d-%b-%Y %I.%M.%S %p")
    except Exception:
        return pd.NaT

def convert_to_iif(df):
    output = StringIO()
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        date = row["Bill Date"].strftime("%m/%d/%Y")
        bill_no = str(row["Bill No."])
        amount = float(row["Amount"])
        till = str(row["Till#"])
        memo = f"Till: {till} Bill#: {bill_no}"

        output.write(f"TRNS\tPAYMENT\t{date}\tCash in Drawer\tWalk In\t{memo}\t{amount:.2f}\t{bill_no}\n")
        output.write(f"SPL\tPAYMENT\t{date}\tAccounts Receivable\tWalk In\t{memo}\t{-amount:.2f}\t\t\n")
        output.write("ENDTRNS\n")

    return output.getvalue()

if uploaded_file:
    try:
        # Read Excel file skipping rows before headers
        df = pd.read_excel(uploaded_file, skiprows=15)

        # Rename columns based on row structure
        df = df.rename(columns={
            "Till# Till#": "Till#",
            "Bill Date Bill Date": "Bill Date",
            "Bill No. Bill No.": "Bill No.",
            "Amount Amount": "Amount"
        })

        # Drop rows missing required data
        df = df[["Till#", "Bill Date", "Bill No.", "Amount"]].dropna()

        # Parse dates
        df["Bill Date"] = df["Bill Date"].astype(str).apply(parse_custom_date)
        df = df.dropna(subset=["Bill Date"])

        # Convert to IIF
        iif_content = convert_to_iif(df)
        st.download_button("📥 Download IIF File", iif_content, file_name="cash_sales.iif", mime="text/plain")

        st.success("✅ File processed successfully!")

    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")

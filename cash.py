# streamlit_app.py

import streamlit as st
import pandas as pd
from io import StringIO
from datetime import datetime

st.set_page_config(page_title="Cash Sales to QuickBooks IIF", layout="wide")
st.title("🧾 Convert Cash Sales Excel to QuickBooks IIF")

uploaded_file = st.file_uploader("Upload Excel file with cash sales", type=["xlsx"])

if uploaded_file:
    try:
        # Read with header rows 12 and 13
        df_raw = pd.read_excel(uploaded_file, header=[11, 12], skiprows=0)

        # Remove rows before row 17 (0-indexed as 16)
        df = df_raw.iloc[16:].copy()
        df.columns = [' '.join(col).strip() for col in df.columns.values]

        # Extract only Cash transactions
        cash_rows = df[df.iloc[:, 0].astype(str).str.contains("cash", case=False, na=False)]

        # Extract relevant columns
        till_col = 'Till# Till#'         # Column E
        date_col = 'Bill Date Bill Date' # Column J
        bill_col = 'Bill No. Bill No.'   # Column P
        amount_col = 'Amount Amount'     # Column Z

        df_cash = cash_rows[[till_col, date_col, bill_col, amount_col]].copy()
        df_cash.columns = ['Till', 'Bill Date', 'Bill No', 'Amount']

        # Format date
        df_cash['Date'] = pd.to_datetime(df_cash['Bill Date'], errors='coerce').dt.strftime('%m/%d/%Y')

        # Filter only rows with valid amounts
        df_cash = df_cash[pd.to_numeric(df_cash['Amount'], errors='coerce').notnull()]

        # === Generate IIF ===
        output = StringIO()
        output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
        output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
        output.write("!ENDTRNS\n")

        for _, row in df_cash.iterrows():
            memo = f"Till: {row['Till']} | Bill: {row['Bill No']}"
            output.write(f"TRNS\tCASH SALE\t{row['Date']}\tCash in Drawer\tWalk In\t{memo}\t{row['Amount']}\t{row['Bill No']}\n")
            output.write(f"SPL\tCASH SALE\t{row['Date']}\tAccounts Receivable\tWalk In\t{memo}\t{-float(row['Amount'])}\t\t\n")
            output.write("ENDTRNS\n")

        st.success("✅ IIF file generated successfully.")
        st.download_button("Download .IIF file", output.getvalue(), file_name="cash_sales.iif", mime="text/plain")

    except Exception as e:
        st.error(f"⚠️ Failed to process file: {e}")

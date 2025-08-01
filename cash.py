# cash_sales_to_iif.py

import streamlit as st
import pandas as pd
from io import StringIO
from datetime import datetime

st.set_page_config(page_title="Cash Sales to QuickBooks IIF", layout="wide")
st.title("🧾 Convert Cash Sales Statement to QuickBooks IIF")

uploaded_file = st.file_uploader("Upload Excel Statement", type=["xlsx"])

if uploaded_file:
    try:
        # Read from row 17 onward
        df = pd.read_excel(uploaded_file, header=None, skiprows=16)

        # Manually assign key columns
        df = df.rename(columns={
            4: 'Till#',       # Column E
            9: 'Bill Date',   # Column J
            15: 'Bill No.',   # Column P
            25: 'Amount'      # Column Z
        })

        # Drop rows with missing values in required fields
        df = df[['Till#', 'Bill Date', 'Bill No.', 'Amount']].dropna()

        # Clean and convert date
        df['Bill Date'] = pd.to_datetime(df['Bill Date'], errors='coerce')
        df = df.dropna(subset=['Bill Date'])

        # Sort and convert amount
        df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
        df = df[df['Amount'] > 0]

        # Format IIF
        output = StringIO()
        output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
        output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
        output.write("!ENDTRNS\n")

        for _, row in df.iterrows():
            date = row['Bill Date'].strftime('%m/%d/%Y')
            memo = f"Till {row['Till#']}"
            amount = round(row['Amount'], 2)
            docnum = row['Bill No.']

            output.write(f"TRNS\tCASH\t{date}\tCash in Drawer\tWalk In\t{memo}\t{amount}\t{docnum}\n")
            output.write(f"SPL\tCASH\t{date}\tAccounts Receivable\tWalk In\t{memo}\t{-amount}\t\t\n")
            output.write("ENDTRNS\n")

        st.download_button("⬇️ Download IIF File", output.getvalue(), file_name="cash_sales.iif")

        st.success("✅ File processed successfully!")

    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")

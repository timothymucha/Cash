import streamlit as st
import pandas as pd
from io import StringIO

st.set_page_config(page_title="Cash Sales to QuickBooks IIF", layout="centered")
st.title("🧾 Convert Cash Sales to QuickBooks IIF")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load raw Excel
        raw_df = pd.read_excel(uploaded_file, header=[11, 12])  # Rows 12 and 13 as header (Python is 0-based)

        # Combine multilevel headers
        raw_df.columns = [' '.join(col).strip() for col in raw_df.columns.values]

        # Remove first rows before actual data (up to row 15)
        data_df = raw_df.iloc[4:].copy()  # Starts from row 16 (zero-indexed)

        # Clean column names
        bill_date_col = [col for col in data_df.columns if 'Bill Date' in col][0]
        bill_no_col = [col for col in data_df.columns if 'Bill No' in col][0]
        amount_col = [col for col in data_df.columns if 'Amount' in col][0]

        # Filter out rows with missing required data
        data_df = data_df[[bill_date_col, bill_no_col, amount_col]]
        data_df.columns = ['Bill Date', 'Bill No', 'Amount']
        data_df.dropna(subset=['Bill Date', 'Bill No', 'Amount'], inplace=True)

        # Remove non-cash if present — assuming we’re already in "Cash" section
        data_df['Bill Date'] = pd.to_datetime(data_df['Bill Date'])

        # Create IIF output
        output = StringIO()
        output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
        output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
        output.write("!ENDTRNS\n")

        for _, row in data_df.iterrows():
            date = row['Bill Date'].strftime('%m/%d/%Y')
            docnum = str(row['Bill No'])
            amount = float(row['Amount'])

            # IIF lines
            output.write(f"TRNS\tCASH\t{date}\tCash in Drawer\tWalk In\tCash Sale\t{amount}\t{docnum}\n")
            output.write(f"SPL\tCASH\t{date}\tAccounts Receivable\tWalk In\tCash Sale\t{-amount}\t\t\n")
            output.write("ENDTRNS\n")

        st.success("✅ File converted. Download your IIF below:")
        st.download_button("📥 Download IIF File", data=output.getvalue(), file_name="cash_sales.iif", mime="text/plain")

    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")

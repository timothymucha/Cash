# streamlit_app.py

import streamlit as st
import pandas as pd
from io import StringIO
from datetime import datetime

st.set_page_config(page_title="Excel to IIF - Cash Sales", layout="wide")
st.title("🧾 Convert Excel Cash Sales to QuickBooks IIF")

uploaded_file = st.file_uploader("Upload your Excel (.xlsx) sales statement", type="xlsx")

def clean_excel_sheet(df_raw):
    # Unmerge-style fix: fill down merged rows
    df_clean = df_raw.ffill()

    # Try to detect the header row (example: it might start with 'Bill No' or similar)
    header_row_index = df_clean[df_clean.iloc[:, 0].astype(str).str.contains("Bill", case=False, na=False)].index.min()
    if pd.isna(header_row_index):
        st.error("Could not detect header row. Please ensure 'Bill No' or similar appears in a column header.")
        return None

    df_clean.columns = df_clean.iloc[header_row_index]
    df_clean = df_clean.iloc[header_row_index + 1:]

    # Drop empty columns and rows
    df_clean = df_clean.dropna(axis=1, how='all').dropna(axis=0, how='all')

    return df_clean

def convert_to_iif(df):
    output = StringIO()

    # IIF Headers
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        try:
            bill_date = pd.to_datetime(row['Bill Date']).strftime('%m/%d/%Y')
            amount = float(row['Amount'])
            docnum = str(row['Bill No'])
            memo = str(row.get('Memo', 'Cash Sale'))

            output.write(f"TRNS\tPAYMENT\t{bill_date}\tCash in Drawer\tWalk In\t{memo}\t{amount:.2f}\t{docnum}\n")
            output.write(f"SPL\tPAYMENT\t{bill_date}\tAccounts Receivable\tWalk In\t{memo}\t{-amount:.2f}\t\t\n")
            output.write("ENDTRNS\n")
        except Exception as e:
            st.warning(f"Skipped a row due to error: {e}")
    
    return output.getvalue()

if uploaded_file:
    try:
        # Load and clean Excel sheet
        df_raw = pd.read_excel(uploaded_file, header=None)
        df_clean = clean_excel_sheet(df_raw)

        if df_clean is not None:
            # Only cash payments
            df_cash = df_clean[df_clean['Payment Mode'].astype(str).str.lower() == 'cash']

            required_columns = ['Bill No', 'Bill Date', 'Amount']
            if not all(col in df_cash.columns for col in required_columns):
                st.error(f"Missing required columns. Ensure your file contains: {required_columns}")
            else:
                iif_data = convert_to_iif(df_cash)

                st.download_button(
                    label="📥 Download IIF file",
                    data=iif_data,
                    file_name="cash_sales.iif",
                    mime="text/plain"
                )

                st.success("✅ Conversion complete! Click above to download the .IIF file.")
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

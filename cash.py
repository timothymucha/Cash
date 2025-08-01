import streamlit as st
import pandas as pd
from io import StringIO

st.set_page_config(page_title="Cash Sales to IIF Converter", layout="wide")
st.title("📄 Convert Excel Cash Sales to QuickBooks IIF")

uploaded_file = st.file_uploader("📤 Upload Cash Sales Excel File", type=["xlsx"])

def truncate_at_blank(df):
    """Stop reading at the first completely blank row"""
    for i, row in df.iterrows():
        if row.isnull().all():
            return df.iloc[:i]
    return df

def generate_iif(df):
    output = StringIO()

    # IIF Headers
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        trns_date = row['Date'].strftime("%m/%d/%Y")
        amount = float(row['Amount'])
        docnum = str(row['Bill No.'])
        till = str(row['Till No'])

        memo = f"Till {till} | Invoice {docnum}"

        # Write the transaction
        output.write(f"TRNS\tCASH\t{trns_date}\tCash in Drawer\tWalk In\t{memo}\t{amount}\t{docnum}\n")
        output.write(f"SPL\tCASH\t{trns_date}\tRevenue:Cash Sales\tWalk In\t{memo}\t{-amount}\t\t\n")
        output.write("ENDTRNS\n")

    return output.getvalue()

if uploaded_file:
    try:
        # Read raw data from row 17 onwards
        df_raw = pd.read_excel(uploaded_file, header=None, skiprows=16)

        # Rename relevant columns
        df_raw.rename(columns={
            4: "Till No",
            9: "Date",
            15: "Bill No.",
            25: "Amount"
        }, inplace=True)

        # Keep only the relevant columns
        df = df_raw[["Till No", "Date", "Bill No.", "Amount"]].copy()

        # Truncate at the first blank row
        df = truncate_at_blank(df)

        # Drop rows where essential fields are missing
        df.dropna(subset=["Date", "Amount"], inplace=True)

        # Convert Date
        df["Date"] = pd.to_datetime(df["Date"], format="%d-%b-%Y %I.%M.%S %p", errors="coerce")

        st.subheader("🧾 Preview: First 10 Cleaned Cash Sales")
        st.dataframe(df.head(10))

        # Generate IIF
        iif_text = generate_iif(df)

        st.subheader("📥 Download .IIF File")
        st.download_button("Download IIF", iif_text, file_name="cash_sales.iif", mime="text/plain")

    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")

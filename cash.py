import streamlit as st
import pandas as pd
from io import StringIO
from openpyxl import load_workbook

st.set_page_config(page_title="Cash Sales to IIF", layout="wide")
st.title("🧾 Convert Cash Sales XLSX to QuickBooks IIF")

def read_excel_unmerged(file):
    wb = load_workbook(file, data_only=True)
    ws = wb.active

    # Fill merged cell values
    for merged_cell_range in ws.merged_cells.ranges:
        merged_cell = ws[merged_cell_range.coord]
        first_cell_value = merged_cell[0][0].value
        for row in merged_cell:
            for cell in row:
                cell.value = first_cell_value
        ws.unmerge_cells(merged_cell_range.coord)

    # Convert to dataframe from row 17 onwards
    data = []
    for row in ws.iter_rows(min_row=17, values_only=True):
        data.append(row)
    df = pd.DataFrame(data)

    # Manually assign expected columns based on known structure
    df.columns = [f"col_{i}" for i in range(len(df.columns))]
    df = df.rename(columns={
        "col_4": "Till",        # Column E
        "col_9": "Bill Date",   # Column J
        "col_15": "Bill No.",   # Column P
        "col_25": "Amount"      # Column Z
    })

    # Drop rows until 'Cash' ends
    df = df[df["Till"].notna()]
    df = df[df["Till"].astype(str).str.lower().str.contains("cash", na=False)]

    # Keep relevant columns
    df = df[["Till", "Bill Date", "Bill No.", "Amount"]]
    df = df.dropna(subset=["Bill Date", "Bill No.", "Amount"])

    return df

def convert_to_iif(df):
    output = StringIO()
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        try:
            date_str = pd.to_datetime(row["Bill Date"]).strftime("%m/%d/%Y")
        except:
            continue  # Skip bad date rows

        docnum = str(row["Bill No."]).strip()
        memo = f"Till: {row['Till']}"
        amount = float(row["Amount"])

        output.write(f"TRNS\tCASH\t{date_str}\tCash in Drawer\tWalk In\t{memo}\t{amount:.2f}\t{docnum}\n")
        output.write(f"SPL\tCASH\t{date_str}\tAccounts Receivable\tWalk In\t{memo}\t{-amount:.2f}\t\t\n")
        output.write("ENDTRNS\n")

    return output.getvalue()

uploaded_file = st.file_uploader("📂 Upload Sales XLSX File", type=["xlsx"])

if uploaded_file:
    try:
        df = read_excel_unmerged(uploaded_file)

        if df.empty:
            st.error("🚫 No cash transactions found or failed to extract data. Check file format.")
        else:
            st.success("✅ Data extracted successfully.")
            st.subheader("🔍 Preview (First 10 rows)")
            st.dataframe(df.head(10))

            iif_content = convert_to_iif(df)
            st.download_button(
                label="⬇️ Download IIF File",
                data=iif_content,
                file_name="cash_sales.iif",
                mime="text/plain"
            )
    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")

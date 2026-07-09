import streamlit as st
import pandas as pd
from io import StringIO, BytesIO

st.set_page_config(page_title="Payment Table Converter", layout="wide")

st.title("Payment Table Converter")

paste_text = st.text_area("Paste copied Outlook table here", height=300)

OUTPUT_COLUMNS = [
    "Swift Code",
    "IBAN Number",
    "Description",
    "Payment Amount",
    "Branch"
]

def normalize_header(x):
    return str(x).strip().lower().replace(" ", "").replace("_", "")

COLUMN_MAP = {
    "employeeid": "Employee ID",
    "employeecode": "Employee ID",
    "empid": "Employee ID",

    "employeename": "Employee Name",
    "name": "Employee Name",

    "swiftcode": "Swift Code",
    "swift": "Swift Code",

    "bank": "Bank",
    "bankname": "Bank",

    "ibannumber": "IBAN Number",
    "iban": "IBAN Number",

    "description": "Description",
    "details": "Description",
    "remarks": "Description",

    "paymentamount": "Payment Amount",
    "amount": "Payment Amount",
    "payamount": "Payment Amount",

    "branch": "Branch",
    "location": "Branch"
}

def read_pasted_table(text):
    text = text.strip()
    if not text:
        return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str)
        if df.shape[1] > 1:
            return df.fillna("")
    except:
        pass

    try:
        df = pd.read_csv(StringIO(text), dtype=str)
        if df.shape[1] > 1:
            return df.fillna("")
    except:
        pass

    return pd.DataFrame()

def clean_and_map(df):
    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    new_cols = {}
    for col in df.columns:
        key = normalize_header(col)
        if key in COLUMN_MAP:
            new_cols[col] = COLUMN_MAP[key]
        else:
            new_cols[col] = str(col).strip()

    df = df.rename(columns=new_cols)

    for col in df.columns:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    if "Employee Name" in df.columns:
        df["Employee Name"] = df["Employee Name"].str[:35]

    if "Description" not in df.columns:
        df["Description"] = ""

    if "Employee ID" in df.columns:
        df["Description"] = (
            df["Description"].astype(str).str.strip()
            + " "
            + df["Employee ID"].astype(str).str.strip()
        ).str.strip()

    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    final_df = df[OUTPUT_COLUMNS].copy()

    return final_df

if st.button("Convert"):
    df = read_pasted_table(paste_text)

    if df.empty:
        st.error("Could not read table. Please copy the full table from Outlook and paste again.")
    else:
        final_df = clean_and_map(df)

        st.success("Converted to fixed output model.")
        edited_df = st.data_editor(final_df, use_container_width=True, num_rows="dynamic")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Payment")

        st.download_button(
            "Download Excel",
            output.getvalue(),
            "payment_output.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

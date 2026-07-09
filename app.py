import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
import re

st.set_page_config(page_title="Payment Excel Converter", layout="wide")

st.title("Payment Excel Converter")

paste_text = st.text_area("Paste copied Outlook table here", height=300)

OUTPUT_COLUMNS = [
    "BRANCH",
    "BENEFICIARY NAME",
    "BENEFICIARY ACCOUNT NUMBER",
    "SWIFT CODE",
    "DESCRIPTION"
]

def normalize_header(col):
    col = str(col).strip().lower()
    col = re.sub(r"[^a-z0-9]", "", col)
    return col

COLUMN_MAP = {
    "employeeid": "Employee ID",
    "employeecode": "Employee ID",
    "empid": "Employee ID",
    "empcode": "Employee ID",

    "employeename": "Employee Name",
    "employeenameenglish": "Employee Name",
    "beneficiaryname": "Employee Name",
    "beneficaryname": "Employee Name",
    "name": "Employee Name",

    "swiftcode": "Swift Code",
    "swift": "Swift Code",

    "ibannumber": "IBAN Number",
    "iban": "IBAN Number",
    "beneficiaryaccountnumber": "IBAN Number",
    "beneficaryaccountnumber": "IBAN Number",
    "accountnumber": "IBAN Number",

    "description": "Description",
    "remarks": "Description",
    "remark": "Description",

    "branch": "Branch",
    "location": "Branch",
}

def read_pasted_table(text):
    text = text.strip()
    if not text:
        return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str)
        if df.shape[1] > 1:
            return df.fillna("")
    except Exception:
        pass

    try:
        df = pd.read_csv(StringIO(text), dtype=str)
        if df.shape[1] > 1:
            return df.fillna("")
    except Exception:
        pass

    return pd.DataFrame()

def process_data(df):
    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    rename_dict = {}
    for col in df.columns:
        key = normalize_header(col)
        rename_dict[col] = COLUMN_MAP.get(key, col)

    df = df.rename(columns=rename_dict)

    for col in df.columns:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    if "Employee ID" not in df.columns:
        df["Employee ID"] = ""

    if "Employee Name" not in df.columns:
        df["Employee Name"] = ""

    if "IBAN Number" not in df.columns:
        df["IBAN Number"] = ""

    if "Swift Code" not in df.columns:
        df["Swift Code"] = ""

    if "Branch" not in df.columns:
        df["Branch"] = ""

    if "Description" not in df.columns:
        df["Description"] = ""

    df["Employee Name"] = (
        df["Employee Name"]
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str[:35]
    )

    df["Description"] = (
        df["Description"].astype(str).str.strip()
        + " "
        + df["Employee ID"].astype(str).str.strip()
    ).str.strip()

    output_df = pd.DataFrame()
    output_df["BRANCH"] = df["Branch"]
    output_df["BENEFICIARY NAME"] = df["Employee Name"]
    output_df["BENEFICIARY ACCOUNT NUMBER"] = df["IBAN Number"]
    output_df["SWIFT CODE"] = df["Swift Code"]
    output_df["DESCRIPTION"] = df["Description"]

    return output_df

if st.button("Convert"):
    df = read_pasted_table(paste_text)

    if df.empty:
        st.error("Could not read table. Please copy the full Outlook table and paste again.")
    else:
        final_df = process_data(df)

        st.success("Converted to required output model.")
        edited_df = st.data_editor(
            final_df,
            use_container_width=True,
            num_rows="dynamic"
        )

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Payment")

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="payment_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

import streamlit as st
import pandas as pd
from io import StringIO, BytesIO

st.set_page_config(page_title="Copy Paste Table to Excel", layout="wide")

st.title("Copy Paste Table to Excel")

paste_text = st.text_area("Paste copied table here", height=300)


def read_pasted_table(text):
    text = text.strip()
    if not text:
        return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str)
        if df.shape[1] > 1:
            df.columns = df.columns.astype(str).str.strip()
            return df.fillna("")
    except Exception:
        pass

    try:
        df = pd.read_csv(StringIO(text), dtype=str)
        if df.shape[1] > 1:
            df.columns = df.columns.astype(str).str.strip()
            return df.fillna("")
    except Exception:
        pass

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    rows = [line.split() for line in lines]
    return pd.DataFrame(rows).fillna("")


def process_data(df):
    df = df.copy()

    # Trim header spaces
    df.columns = df.columns.astype(str).str.strip()

    # Trim all cell spaces
    for col in df.columns:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    # Employee Name = TRIM + LEFT(35)
    if "Employee Name" in df.columns:
        df["Employee Name"] = df["Employee Name"].str[:35]

    # Description = Description + space + Employee ID
    if "Description" in df.columns and "Employee ID" in df.columns:
        df["Description"] = (
            df["Description"].astype(str).str.strip()
            + " "
            + df["Employee ID"].astype(str).str.strip()
        ).str.strip()

    return df


if st.button("Convert"):
    df = read_pasted_table(paste_text)

    if df.empty:
        st.error("No data found. Please paste table data.")
    else:
        final_df = process_data(df)

        st.success("Table converted. Please verify before download.")

        edited_df = st.data_editor(
            final_df,
            use_container_width=True,
            num_rows="dynamic"
        )

        # Remove Employee ID and Bank only while saving
        save_df = edited_df.copy()
        save_df = save_df.drop(
            columns=[c for c in ["Employee ID", "Bank"] if c in save_df.columns],
            errors="ignore"
        )

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            save_df.to_excel(writer, index=False, sheet_name="Data")

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="converted_table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
